#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json, os, sys, subprocess, socket
import psycopg2
import psycopg2.extras
from datetime import datetime
from flask import Flask, request, render_template, redirect, url_for, jsonify, send_file

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

BASE         = os.path.dirname(os.path.abspath(__file__))
DATA_DIR     = os.environ.get("DATA_DIR", BASE)
DATABASE_URL = os.environ.get("DATABASE_URL")
JSON_OUT     = os.path.join(DATA_DIR, "datos_social_icetex.json")
HTML_OUT     = os.path.join(DATA_DIR, "reporte_social_icetex.html")
XLSX_OUT     = os.path.join(DATA_DIR, "reporte_social_icetex.xlsx")

PLATAFORMAS = ["Facebook", "Instagram", "X (Twitter)", "LinkedIn", "TikTok", "YouTube"]

app = Flask(__name__, template_folder="templates")
app.secret_key = os.environ.get("SECRET_KEY", "dev-key-local")

# ── Base de datos
def get_db():
    url = DATABASE_URL
    if url and url.startswith("postgres://"):
        url = url.replace("postgres://", "postgresql://", 1)
    conn = psycopg2.connect(url, cursor_factory=psycopg2.extras.RealDictCursor)
    return conn

def init_db():
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS envios (
                    id        SERIAL PRIMARY KEY,
                    analista  TEXT,
                    fb_cuenta TEXT,
                    ig_cuenta TEXT,
                    datos     TEXT,
                    sla       TEXT,
                    fecha     TEXT
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS metricas_plataforma (
                    id            SERIAL PRIMARY KEY,
                    plataforma    TEXT NOT NULL,
                    mes           INTEGER NOT NULL,
                    anio          INTEGER NOT NULL,
                    seguidores    BIGINT DEFAULT 0,
                    seguidores_ant BIGINT DEFAULT 0,
                    visualizaciones BIGINT DEFAULT 0,
                    likes         INTEGER DEFAULT 0,
                    comentarios   INTEGER DEFAULT 0,
                    compartidos   INTEGER DEFAULT 0,
                    guardados     INTEGER DEFAULT 0,
                    publicaciones INTEGER DEFAULT 0,
                    respondidas   INTEGER DEFAULT 0,
                    alcance_ext_pct REAL DEFAULT 0,
                    sent_pos      INTEGER DEFAULT 0,
                    sent_neg      INTEGER DEFAULT 0,
                    sent_neu      INTEGER DEFAULT 0,
                    sent_crit     INTEGER DEFAULT 0,
                    UNIQUE(plataforma, mes, anio)
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS publicaciones_dashboard (
                    id            SERIAL PRIMARY KEY,
                    plataforma    TEXT NOT NULL,
                    mes           INTEGER NOT NULL,
                    anio          INTEGER NOT NULL,
                    texto         TEXT,
                    likes         INTEGER DEFAULT 0,
                    comentarios   INTEGER DEFAULT 0,
                    compartidos   INTEGER DEFAULT 0,
                    guardados     INTEGER DEFAULT 0,
                    visualizaciones INTEGER DEFAULT 0,
                    alcance       INTEGER DEFAULT 0,
                    tipo          TEXT DEFAULT 'normal'
                )
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS temas_dashboard (
                    id        SERIAL PRIMARY KEY,
                    mes       INTEGER NOT NULL,
                    anio      INTEGER NOT NULL,
                    tema      TEXT NOT NULL,
                    menciones INTEGER DEFAULT 0,
                    UNIQUE(tema, mes, anio)
                )
            """)
        conn.commit()
    finally:
        conn.close()

# ── Conclusiones automáticas
def generar_conclusiones(plataformas, sentimiento, temas):
    conclusiones = []

    total_coms = sum(p.get("total_comentarios", 0) for p in plataformas.values())
    total_resp = sum(p.get("respondidos", 0) for p in plataformas.values())
    if total_coms > 0:
        tasa = total_resp / total_coms * 100
        if tasa < 50:
            conclusiones.append({"tipo": "alerta", "texto": f"La tasa de respuesta global es del {tasa:.0f}%, por debajo del mínimo recomendado (70%). Se sugiere reforzar los turnos de atención digital."})
        elif tasa >= 80:
            conclusiones.append({"tipo": "exito", "texto": f"Excelente tasa de respuesta global del {tasa:.0f}%. La entidad mantiene una atención digital oportuna y efectiva."})
        else:
            conclusiones.append({"tipo": "info", "texto": f"La tasa de respuesta global es del {tasa:.0f}%. Existe oportunidad de mejora para alcanzar el objetivo del 80%."})

    for plat, m in plataformas.items():
        evol = m.get("evolucion_pct", 0)
        if evol > 20:
            conclusiones.append({"tipo": "exito", "texto": f"{plat} muestra un incremento del {evol:.1f}% en volumen de comentarios frente al mes anterior."})
        elif evol < -20:
            conclusiones.append({"tipo": "alerta", "texto": f"{plat} registra una caída del {abs(evol):.1f}% en comentarios. Se recomienda revisar la estrategia de contenido."})

    total_sent = sum(sentimiento.values())
    if total_sent > 0:
        neg  = (sentimiento.get("Negativo", 0) + sentimiento.get("Crítico", 0)) / total_sent * 100
        pos  = sentimiento.get("Positivo", 0) / total_sent * 100
        crit = sentimiento.get("Crítico", 0) / total_sent * 100
        if crit > 20:
            conclusiones.append({"tipo": "alerta", "texto": f"El {crit:.0f}% de los comentarios son críticos. Se recomienda activar protocolo de gestión de crisis digital."})
        elif neg > 40:
            conclusiones.append({"tipo": "alerta", "texto": f"El {neg:.0f}% de los comentarios presentan tono negativo. Se sugiere reforzar la calidad de las respuestas."})
        elif pos > 60:
            conclusiones.append({"tipo": "exito", "texto": f"El {pos:.0f}% de los comentarios son positivos, reflejando una percepción favorable de la entidad."})

    if temas:
        top = max(temas, key=temas.get)
        conclusiones.append({"tipo": "info", "texto": f"El tema con mayor volumen de menciones es '{top}' con {temas[top]} interacciones. Se recomienda priorizar contenido informativo sobre este tema."})

    if not conclusiones:
        conclusiones.append({"tipo": "info", "texto": "Ingresa datos desde el formulario de comentarios para generar conclusiones automáticas."})

    return conclusiones

# ── Detección de sentimiento server-side
_EMOCIONES = {
    "Crítico": [
        "queja","reclamo","denuncia","tutela","demanda","acción legal","accion legal",
        "derecho de petición","derecho de peticion","superintendencia","defensoría","defensoria",
        "procuraduría","procuraduria","contraloría","contraloria","personería","personeria",
        "estafa","fraude","robo","hurto","corrupción","corrupcion","malversación","malversacion",
        "inaceptable","indignante","vergüenza","verguenza","escándalo","escandalo","aberración",
        "aberracion","injusticia","ilegal","inconstitucional","arbitrario","arbitraria","impunidad",
        "irresponsable","irresponsabilidad","incompetente","incompetencia",
        "abuso","negligencia","perjuicio","incumplimiento","engaño","engano","mentira","falsedad",
        "manipulación","manipulacion","discriminación","discriminacion","acoso","amenaza",
        "no hacen lo que les corresponde","no cumplen su función","no cumplen su funcion",
        "ya basta","hasta cuándo","hasta cuando","basta ya","no aguanto más","no aguanto mas",
        "indignado","indignada","furioso","furiosa","desesperado","desesperada","impotente",
        "rabia","asco","repudio","nunca más","nunca mas","no vuelvo","los denuncio",
        "voy a demandar","tomaré acciones","tomare acciones",
        "de qué sirve","de que sirve","para qué sirve","para que sirve",
        "educación digna","educacion digna","derecho a la educación","derecho a la educacion",
        "paguen el subsidio","paguen el dinero","no pagan","siguen sin pagar",
        "proceso fallido","burocracia excesiva","trabas burocráticas","trabas burocraticas",
        "sin como estudiar","sin poder estudiar",
    ],
    "Negativo": [
        "malo","mala","mal ","terrible","pésimo","pesimo","horrible","fatal","deficiente",
        "deplorable","lamentable","desastroso","desastrosa","pobre","mediocre","nefasto","nefasta",
        "malísimo","malisimo","muy malo","muy mala","paupérrimo","pauperrimo",
        "molesto","molesta","frustrado","frustrada","decepcionado","decepcionada","decepción",
        "decepcion","triste","angustiado","angustiada","preocupado","preocupada",
        "cansado","cansada","harto","harta","agotado","agotada","estresado","estresada",
        "desmotivado","desmotivada","confundido","confundida","perdido","perdida",
        "insatisfecho","insatisfecha","inconformidad","malestar",
        "no funciona","no sirve","no carga","no responden","nadie responde","sin respuesta",
        "ignorando","ignorado","ignorada","abandonado","abandonada","olvidado","olvidada",
        "no me atienden","no me ayudan","no me solucionan","no dan razón","no dan razon",
        "no dan información","no dan informacion","no explican",
        "demora","demorado","retraso","retrasado","tardanza","lento","lenta",
        "hace días","hace dias","hace semanas","hace meses","hace un año",
        "llevo esperando","sigo esperando","aún no","aun no","todavía no","todavia no",
        "sin novedad","sin solución","sin solucion","semanas sin","meses sin",
        "bloqueado","bloqueada","suspendido","suspendida","cancelado","cancelada",
        "rechazado","rechazada","negado","negada","no han desembolsado","no llegó","no llego",
        "cobro incorrecto","cobro errado","me cobraron mal","deuda incorrecta","saldo incorrecto",
        "beca no llegó","beca no llego","subsidio no llegó","subsidio no llego",
        "no me renovaron","no me legalizaron","me negaron",
        "nadie sabe","nadie me dice","no hay información","no hay informacion",
        "información confusa","informacion confusa","muy confuso","muy confusa",
        "no es claro","no es clara","cambian los requisitos","cambian las reglas",
        "portal caído","portal caido","sistema caído","sistema caido","error en el sistema",
        "no puedo ingresar","no puedo acceder","no me deja","me bloquea",
    ],
    "Positivo": [
        "gracias","muchas gracias","mil gracias","infinitas gracias",
        "agradecido","agradecida","agradezco","agradecemos","aprecio","apreciamos",
        "feliz","felices","contento","contenta","satisfecho","satisfecha",
        "excelente","genial","perfecto","perfecta","bueno","buena","bien ","muy bien",
        "maravilloso","maravillosa","increíble","increible","fantástico","fantastico",
        "estupendo","estupenda","espectacular","extraordinario","extraordinaria",
        "sobresaliente","notable","impecable",
        "encantado","encantada","felicitaciones","felicito","bravo","bien hecho",
        "muy profesional","muy amable","muy cordial","muy atento","muy atenta",
        "resolvieron","solucionaron","atendieron","respondieron rápido","respondieron rapido",
        "me ayudaron","me orientaron","lo lograron","cumplieron","eficiente","eficientes",
        "eficaz","buen servicio","excelente servicio","buen trabajo","excelente atención",
        "excelente atencion","gran apoyo","muy eficaz","lo solucionaron","lo resolvieron",
        "me dieron solución","me dieron solucion","me respondieron a tiempo",
        "rápido","rapido","ágil","agil","inmediato","sin demora","a tiempo","sin problemas",
        "sin inconvenientes","todo bien","fácil","facil","sencillo","sencilla","claro","clara",
        "bien explicado","muy claro","muy clara",
        "me aprobaron","me desembolsaron","recibí el dinero","recibi el dinero",
        "llegó el desembolso","llego el desembolso","recibí el subsidio","recibi el subsidio",
        "llegó la beca","llego la beca","me renovaron","me legalizaron","proceso exitoso",
        "sin complicaciones","me apoyaron","me asesoraron bien",
    ],
}

def detectar_emocion_srv(texto: str) -> str:
    t = (texto or "").lower()
    conteos = {}
    for em in ["Crítico", "Negativo", "Positivo"]:
        conteos[em] = sum(1 for p in _EMOCIONES[em] if p in t)
    conteos["Crítico"] *= 2  # peso doble para casos graves
    mejor = max(conteos, key=conteos.get)
    return mejor if conteos[mejor] > 0 else "Neutro"

def reprocesar_emociones(datos: dict) -> dict:
    for plat, pubs in datos.items():
        for pub in pubs:
            for c in pub.get("comentarios", []):
                c["emocion"] = detectar_emocion_srv(c.get("texto", ""))
    return datos

# ── Helpers originales
def get_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "localhost"

def consolidar_datos():
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM envios ORDER BY fecha DESC")
            rows = cur.fetchall()
    finally:
        conn.close()
    if not rows:
        return None, []
    consolidado = {"twitter": [], "facebook": [], "instagram": [],
                   "linkedin": [], "tiktok": [], "youtube": []}
    analistas   = []
    sla_final   = {"facebook": 8, "instagram": 12, "twitter": 4,
                   "linkedin": 8, "tiktok": 24, "youtube": 24}
    for row in rows:
        try:
            d = json.loads(row["datos"])
            analistas.append(row["analista"] or "Anónimo")
            if row["sla"]:
                s = json.loads(row["sla"])
                sla_final.update(s)
            consolidado["facebook"]  += d.get("facebook",  [])
            consolidado["instagram"] += d.get("instagram", [])
            consolidado["twitter"]   += d.get("twitter",   [])
            consolidado["linkedin"]  += d.get("linkedin",  [])
            consolidado["tiktok"]    += d.get("tiktok",    [])
            consolidado["youtube"]   += d.get("youtube",   [])
        except Exception:
            continue
    return consolidado, analistas, sla_final

# ── Rutas originales
@app.route("/")
def index():
    return redirect(url_for("formulario"))

@app.route("/formulario")
def formulario():
    config_path = os.path.join(BASE, "config.json")
    config = {}
    if os.path.exists(config_path):
        with open(config_path, encoding="utf-8") as f:
            config = json.load(f)
    sla = {
        "facebook":  config.get("facebook",  {}).get("sla_horas", 8),
        "instagram": config.get("instagram", {}).get("sla_horas", 12),
        "twitter":   config.get("twitter",   {}).get("sla_horas", 4),
        "linkedin":  config.get("linkedin",  {}).get("sla_horas", 8),
        "tiktok":    config.get("tiktok",    {}).get("sla_horas", 24),
        "youtube":   config.get("youtube",   {}).get("sla_horas", 24),
    }
    return render_template("formulario.html", sla=sla)

@app.route("/enviar", methods=["POST"])
def enviar():
    body = request.get_json(force=True, silent=True) or {}
    meta      = body.get("_meta", {})
    analista  = meta.get("analista", "Anónimo")
    fb_cuenta = meta.get("fb_cuenta", "")
    ig_cuenta = meta.get("ig_cuenta", "")
    sla       = json.dumps(meta.get("sla", {}))
    datos_raw = {
        "twitter":   body.get("twitter",   []),
        "facebook":  body.get("facebook",  []),
        "instagram": body.get("instagram", []),
        "linkedin":  body.get("linkedin",  []),
        "tiktok":    body.get("tiktok",    []),
        "youtube":   body.get("youtube",   []),
    }
    datos_str = json.dumps(reprocesar_emociones(datos_raw), ensure_ascii=False)
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO envios (analista, fb_cuenta, ig_cuenta, datos, sla, fecha) VALUES (%s,%s,%s,%s,%s,%s)",
                (analista, fb_cuenta, ig_cuenta, datos_str, sla, datetime.now().isoformat())
            )
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True, "mensaje": f"Datos de '{analista}' guardados correctamente."})

@app.route("/admin")
def admin():
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT id, analista, fb_cuenta, ig_cuenta, fecha, datos FROM envios ORDER BY fecha DESC")
            rows = cur.fetchall()
    finally:
        conn.close()
    envios = []
    for r in rows:
        try:
            d = json.loads(r["datos"])
            n_fb = len(d.get("facebook", []))
            n_ig = len(d.get("instagram", []))
            total_coms = sum(len(p["comentarios"]) for plat in d.values() for p in plat)
        except Exception:
            n_fb = n_ig = total_coms = 0
        envios.append({
            "id":       r["id"],
            "analista": r["analista"] or "Anónimo",
            "fb":       r["fb_cuenta"],
            "ig":       r["ig_cuenta"],
            "fecha":    r["fecha"][:16].replace("T", " "),
            "n_fb":     n_fb,
            "n_ig":     n_ig,
            "coms":     total_coms,
        })
    reporte_existe = os.path.exists(HTML_OUT)
    return render_template("admin.html", envios=envios, reporte_existe=reporte_existe)

@app.route("/eliminar/<int:eid>", methods=["POST"])
def eliminar(eid):
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM envios WHERE id=%s", (eid,))
        conn.commit()
    finally:
        conn.close()
    return redirect(url_for("admin"))

@app.route("/generar", methods=["POST"])
def generar():
    result_data = consolidar_datos()
    if result_data is None or result_data[0] is None:
        return jsonify({"ok": False, "mensaje": "No hay envíos para consolidar."})
    consolidado, analistas, sla_final = result_data
    config_path = os.path.join(BASE, "config.json")
    try:
        with open(config_path, encoding="utf-8") as f:
            config = json.load(f)
        config.setdefault("facebook",  {})["sla_horas"] = sla_final.get("facebook",  8)
        config.setdefault("instagram", {})["sla_horas"] = sla_final.get("instagram", 12)
        config.setdefault("twitter",   {})["sla_horas"] = sla_final.get("twitter",   4)
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    with open(JSON_OUT, "w", encoding="utf-8") as f:
        json.dump(consolidado, f, ensure_ascii=False, indent=2, default=str)
    monitor_py = os.path.join(BASE, "monitor.py")
    r = subprocess.run([sys.executable, monitor_py, "--desde-json"], capture_output=True, text=True)
    if r.returncode == 0:
        total = sum(len(v) for v in consolidado.values())
        total_coms = sum(len(c["comentarios"]) for plat in consolidado.values() for c in plat)
        return jsonify({"ok": True, "mensaje": f"Reporte generado con {total} publicaciones y {total_coms} comentarios de: {', '.join(analistas)}"})
    return jsonify({"ok": False, "mensaje": f"Error: {r.stderr[:300]}"})

@app.route("/reporte")
def reporte():
    if not os.path.exists(HTML_OUT):
        return """<!DOCTYPE html><html lang='es'><head><meta charset='UTF-8'>
        <style>body{font-family:'Segoe UI',Arial,sans-serif;background:#f0f2f5;display:flex;
        align-items:center;justify-content:center;min-height:100vh;margin:0}
        .box{background:#fff;border-radius:16px;padding:40px 48px;box-shadow:0 4px 24px rgba(0,0,0,.1);
        text-align:center;max-width:460px}
        h2{color:#1a2340;margin-bottom:12px;font-size:1.3rem}
        p{color:#666;font-size:.9rem;margin-bottom:24px;line-height:1.6}
        a{display:inline-block;background:#1976d2;color:#fff;padding:12px 28px;border-radius:10px;
        text-decoration:none;font-weight:700;font-size:.95rem}
        a:hover{background:#1565c0}</style></head>
        <body><div class='box'>
        <div style='font-size:48px;margin-bottom:16px'>📊</div>
        <h2>El reporte aún no ha sido generado</h2>
        <p>Para generar el reporte ve al panel de administrador,<br>
        asegúrate de tener envíos cargados y haz clic en<br>
        <strong>"📊 Generar reporte ahora"</strong>.</p>
        <a href='/admin'>Ir al Administrador</a>
        </div></body></html>"""
    return send_file(HTML_OUT)

@app.route("/descargar-excel")
def descargar_excel():
    if not os.path.exists(XLSX_OUT):
        return "Excel no disponible", 404
    return send_file(XLSX_OUT, as_attachment=True)

# ── Dashboard
@app.route("/dashboard")
def dashboard():
    return render_template("dashboard.html")

@app.route("/dashboard/ingresar", methods=["GET", "POST"])
def ingresar_metricas():
    mensaje = None
    error   = None
    if request.method == "POST":
        accion = request.form.get("accion", "metricas")
        mes    = int(request.form.get("mes",  datetime.now().month))
        anio   = int(request.form.get("anio", datetime.now().year))
        try:
            conn = get_db()
            try:
                with conn.cursor() as cur:
                    if accion == "metricas":
                        plat = request.form.get("plataforma")
                        cur.execute("""
                            INSERT INTO metricas_plataforma
                            (plataforma,mes,anio,seguidores,seguidores_ant,visualizaciones,
                             likes,comentarios,compartidos,guardados,publicaciones,respondidas,
                             alcance_ext_pct,sent_pos,sent_neg,sent_neu,sent_crit)
                            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                            ON CONFLICT (plataforma,mes,anio) DO UPDATE SET
                            seguidores=EXCLUDED.seguidores,
                            seguidores_ant=EXCLUDED.seguidores_ant,
                            visualizaciones=EXCLUDED.visualizaciones,
                            likes=EXCLUDED.likes,
                            comentarios=EXCLUDED.comentarios,
                            compartidos=EXCLUDED.compartidos,
                            guardados=EXCLUDED.guardados,
                            publicaciones=EXCLUDED.publicaciones,
                            respondidas=EXCLUDED.respondidas,
                            alcance_ext_pct=EXCLUDED.alcance_ext_pct,
                            sent_pos=EXCLUDED.sent_pos,
                            sent_neg=EXCLUDED.sent_neg,
                            sent_neu=EXCLUDED.sent_neu,
                            sent_crit=EXCLUDED.sent_crit
                        """, (
                            plat, mes, anio,
                            int(request.form.get("seguidores", 0) or 0),
                            int(request.form.get("seguidores_ant", 0) or 0),
                            int(request.form.get("visualizaciones", 0) or 0),
                            int(request.form.get("likes", 0) or 0),
                            int(request.form.get("comentarios", 0) or 0),
                            int(request.form.get("compartidos", 0) or 0),
                            int(request.form.get("guardados", 0) or 0),
                            int(request.form.get("publicaciones", 0) or 0),
                            int(request.form.get("respondidas", 0) or 0),
                            float(request.form.get("alcance_ext_pct", 0) or 0),
                            int(request.form.get("sent_pos", 0) or 0),
                            int(request.form.get("sent_neg", 0) or 0),
                            int(request.form.get("sent_neu", 0) or 0),
                            int(request.form.get("sent_crit", 0) or 0),
                        ))
                        mensaje = f"✅ Métricas de {plat} guardadas para {mes}/{anio}."

                    elif accion == "publicacion":
                        cur.execute("""
                            INSERT INTO publicaciones_dashboard
                            (plataforma,mes,anio,texto,likes,comentarios,compartidos,guardados,visualizaciones,alcance,tipo)
                            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                        """, (
                            request.form.get("plataforma"),
                            mes, anio,
                            request.form.get("texto_pub", ""),
                            int(request.form.get("pub_likes", 0) or 0),
                            int(request.form.get("pub_comentarios", 0) or 0),
                            int(request.form.get("pub_compartidos", 0) or 0),
                            int(request.form.get("pub_guardados", 0) or 0),
                            int(request.form.get("pub_visualizaciones", 0) or 0),
                            int(request.form.get("pub_alcance", 0) or 0),
                            request.form.get("pub_tipo", "normal"),
                        ))
                        mensaje = "✅ Publicación guardada."

                    elif accion == "tema":
                        cur.execute("""
                            INSERT INTO temas_dashboard (mes,anio,tema,menciones)
                            VALUES (%s,%s,%s,%s)
                            ON CONFLICT (tema,mes,anio) DO UPDATE SET menciones=EXCLUDED.menciones
                        """, (
                            mes, anio,
                            request.form.get("tema_nombre", "").strip(),
                            int(request.form.get("tema_menciones", 0) or 0),
                        ))
                        mensaje = "✅ Tema guardado."

                conn.commit()
            finally:
                conn.close()
        except Exception as e:
            error = f"Error al guardar: {str(e)}"

    return render_template("ingresar_metricas.html",
                           plataformas=PLATAFORMAS,
                           mensaje=mensaje,
                           error=error,
                           mes_actual=datetime.now().month,
                           anio_actual=datetime.now().year)

@app.route("/admin/reprocesar-sentimiento", methods=["POST"])
def reprocesar_sentimiento():
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT id, datos FROM envios")
            rows = cur.fetchall()
        actualizados = 0
        for row in rows:
            try:
                datos = reprocesar_emociones(json.loads(row["datos"]))
                with conn.cursor() as cur:
                    cur.execute("UPDATE envios SET datos=%s WHERE id=%s",
                                (json.dumps(datos, ensure_ascii=False), row["id"]))
                actualizados += 1
            except Exception:
                continue
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True, "mensaje": f"Sentimiento reprocesado en {actualizados} envíos."})

@app.route("/admin/envio/<int:eid>")
def get_envio(eid):
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM envios WHERE id=%s", (eid,))
            row = cur.fetchone()
    finally:
        conn.close()
    if not row:
        return jsonify({"ok": False}), 404
    try:
        datos = json.loads(row["datos"])
    except Exception:
        datos = {}
    return jsonify({"ok": True, "id": eid, "analista": row["analista"], "datos": datos})

@app.route("/admin/envio/<int:eid>/guardar", methods=["POST"])
def save_envio(eid):
    body  = request.get_json(force=True, silent=True) or {}
    datos = body.get("datos", {})
    conn  = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("UPDATE envios SET datos=%s WHERE id=%s",
                        (json.dumps(datos, ensure_ascii=False), eid))
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True, "mensaje": "Cambios guardados correctamente."})

@app.route("/dashboard/pub/eliminar/<int:pid>", methods=["POST"])
def eliminar_pub(pid):
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM publicaciones_dashboard WHERE id=%s", (pid,))
        conn.commit()
    finally:
        conn.close()
    return redirect(url_for("ingresar_metricas"))

@app.route("/dashboard/tema/eliminar/<int:tid>", methods=["POST"])
def eliminar_tema(tid):
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM temas_dashboard WHERE id=%s", (tid,))
        conn.commit()
    finally:
        conn.close()
    return redirect(url_for("ingresar_metricas"))

@app.route("/api/dashboard")
def api_dashboard():
    mes  = request.args.get("mes",  datetime.now().month, type=int)
    anio = request.args.get("anio", datetime.now().year,  type=int)
    mes_ant  = mes - 1 if mes > 1 else 12
    anio_ant = anio if mes > 1 else anio - 1

    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT datos, fecha FROM envios ORDER BY fecha DESC")
            rows = cur.fetchall()
    finally:
        conn.close()

    def parse_mes_anio(s):
        try:
            dt = datetime.fromisoformat((s or "")[:10])
            return dt.month, dt.year
        except Exception:
            return None, None

    def extraer(filas):
        plats, pubs_list = {}, []
        temas = {}
        sent  = {"Positivo": 0, "Negativo": 0, "Neutro": 0, "Crítico": 0}
        for row in filas:
            try:
                d = json.loads(row["datos"])
            except Exception:
                continue
            for plat, pub_list in d.items():
                for pub in pub_list:
                    coms   = pub.get("comentarios", [])
                    n_coms = len(coms)
                    n_resp = sum(1 for c in coms if c.get("respondido"))
                    tiempos = [c["tiempo_respuesta_min"] for c in coms
                               if c.get("tiempo_respuesta_min") is not None]
                    if plat not in plats:
                        plats[plat] = {"total_comentarios": 0, "respondidos": 0,
                                       "publicaciones": 0, "tiempos": [],
                                       "alcance": 0, "impresiones": 0,
                                       "nuevos_seg": 0, "ctrs": [], "t_vision": []}
                    plats[plat]["total_comentarios"] += n_coms
                    plats[plat]["respondidos"]       += n_resp
                    plats[plat]["publicaciones"]     += 1
                    plats[plat]["tiempos"].extend(tiempos)
                    likes       = pub.get("likes", 0) or 0
                    shares      = pub.get("shares", 0) or 0
                    alcance     = pub.get("alcance", 0) or 0
                    impresiones = pub.get("impresiones", 0) or 0
                    ctr         = pub.get("ctr", 0) or 0
                    nuevos_seg  = pub.get("nuevos_seguidores", 0) or 0
                    t_vision    = pub.get("tiempo_visionado_seg", 0) or 0
                    guardados   = pub.get("guardados", 0) or 0
                    tasa = round(n_resp / n_coms * 100, 1) if n_coms > 0 else 0
                    interacciones = likes + shares + n_coms + guardados
                    tasa_interaccion = round(interacciones / alcance * 100, 2) if alcance > 0 else 0
                    viral_score = (n_coms * 3) + (shares * 2) + (likes * 1) + (guardados * 2)
                    riesgo = round((100 - tasa) * 0.7 + max(0, 10 - n_coms) * 3, 1)
                    plats[plat]["alcance"]     = plats[plat].get("alcance", 0) + alcance
                    plats[plat]["impresiones"] = plats[plat].get("impresiones", 0) + impresiones
                    plats[plat]["nuevos_seg"]  = plats[plat].get("nuevos_seg", 0) + nuevos_seg
                    plats[plat]["ctrs"].append(ctr) if ctr > 0 else None
                    plats[plat]["t_vision"].append(t_vision) if t_vision > 0 else None
                    pubs_list.append({
                        "plataforma":       plat,
                        "texto":            pub.get("texto", ""),
                        "url":              pub.get("url", ""),
                        "fecha":            pub.get("fecha", "")[:10] if pub.get("fecha") else "",
                        "n_coms":           n_coms,
                        "n_resp":           n_resp,
                        "likes":            likes,
                        "shares":           shares,
                        "guardados":        guardados,
                        "alcance":          alcance,
                        "impresiones":      impresiones,
                        "ctr":              ctr,
                        "nuevos_seguidores":nuevos_seg,
                        "tiempo_visionado_seg": t_vision,
                        "tasa":             tasa,
                        "tasa_interaccion": tasa_interaccion,
                        "viral_score":      viral_score,
                        "riesgo":           riesgo,
                    })
                    for c in coms:
                        t = c.get("tema", "otro")
                        temas[t] = temas.get(t, 0) + 1
                        em = c.get("emocion", "Neutro")
                        sent[em] = sent.get(em, 0) + 1
        return plats, pubs_list, temas, sent

    curr, prev = [], []
    for row in rows:
        m, a = parse_mes_anio(row["fecha"])
        if m == mes and a == anio:
            curr.append(row)
        elif m == mes_ant and a == anio_ant:
            prev.append(row)
        elif m is None:
            curr.append(row)

    plats_c, pubs_c, temas_c, sent_c = extraer(curr)
    plats_p, _, _, _                  = extraer(prev)

    plataformas = {}
    for plat, st in plats_c.items():
        total = st["total_comentarios"]
        resp  = st["respondidos"]
        tms   = st["tiempos"]
        prev_total = plats_p.get(plat, {}).get("total_comentarios", 0)
        evol  = round((total - prev_total) / prev_total * 100, 1) if prev_total > 0 else 0
        plataformas[plat] = {
            "total_comentarios":   total,
            "respondidos":         resp,
            "publicaciones":       st["publicaciones"],
            "tasa_respuesta":      round(resp / total * 100, 1) if total > 0 else 0,
            "tiempo_promedio_min": round(sum(tms) / len(tms), 0) if tms else None,
            "comentarios_mes_ant": prev_total,
            "evolucion_pct":       evol,
        }

    # ── Clasificación ±30% respecto al promedio histórico
    if pubs_c:
        avg_score = sum(p["viral_score"] for p in pubs_c) / len(pubs_c)
        umbral_alto = avg_score * 1.30
        umbral_bajo = avg_score * 0.70
        for p in pubs_c:
            if p["viral_score"] >= umbral_alto:
                p["clasificacion"] = "viral"
            elif p["viral_score"] <= umbral_bajo:
                p["clasificacion"] = "bajo"
            else:
                p["clasificacion"] = "promedio"
        avg_score_r = round(avg_score, 1)
    else:
        avg_score_r = 0

    virales = sorted([p for p in pubs_c if p.get("clasificacion")=="viral"],
                     key=lambda x: x["viral_score"], reverse=True)[:5]
    if not virales:
        virales = sorted(pubs_c, key=lambda x: x["viral_score"], reverse=True)[:3]

    textos_virales = {p["texto"] for p in virales}
    bajas = sorted(
        [p for p in pubs_c if p["texto"] not in textos_virales and p.get("clasificacion") in ("bajo", "promedio")],
        key=lambda x: x["riesgo"], reverse=True
    )[:5]

    # ── Métricas globales avanzadas
    total_alcance     = sum(p.get("alcance", 0) for p in pubs_c)
    total_impresiones = sum(p.get("impresiones", 0) for p in pubs_c)
    total_nuevos_seg  = sum(p.get("nuevos_seguidores", 0) for p in pubs_c)
    ctrs_validos      = [p["ctr"] for p in pubs_c if p.get("ctr", 0) > 0]
    t_vision_validos  = [p["tiempo_visionado_seg"] for p in pubs_c if p.get("tiempo_visionado_seg", 0) > 0]
    total_interacciones = sum(p.get("likes",0)+p.get("shares",0)+p.get("n_coms",0)+p.get("guardados",0) for p in pubs_c)
    avg_ctr           = round(sum(ctrs_validos)/len(ctrs_validos), 2) if ctrs_validos else 0
    avg_t_vision      = round(sum(t_vision_validos)/len(t_vision_validos), 0) if t_vision_validos else 0
    tasa_interaccion_global = round(total_interacciones / total_alcance * 100, 2) if total_alcance > 0 else 0

    n_virales  = sum(1 for p in pubs_c if p.get("clasificacion")=="viral")
    n_promedio = sum(1 for p in pubs_c if p.get("clasificacion")=="promedio")
    n_bajos    = sum(1 for p in pubs_c if p.get("clasificacion")=="bajo")

    # ── Recomendaciones automáticas
    recomendaciones = []
    total_pubs = len(pubs_c)
    if total_pubs > 0:
        pct_bajos = n_bajos / total_pubs * 100
        pct_virales = n_virales / total_pubs * 100
        total_resp_global = sum(p.get("n_resp",0) for p in pubs_c)
        total_coms_global = sum(p.get("n_coms",0) for p in pubs_c)
        tasa_resp_global  = round(total_resp_global / total_coms_global * 100, 1) if total_coms_global > 0 else 0

        if pct_bajos > 40:
            recomendaciones.append({"tipo":"alerta","icono":"📉","titulo":"Alto % de publicaciones con bajo rendimiento",
                "texto":f"{n_bajos} de {total_pubs} posts ({pct_bajos:.0f}%) están por debajo del 70% del promedio. Revisar horarios de publicación, formato del contenido y temáticas menos efectivas."})
        if pct_virales > 30:
            mejores = sorted(pubs_c, key=lambda x: x["viral_score"], reverse=True)[:2]
            temas_virales = ", ".join(set(p.get("texto","")[:40] for p in mejores))
            recomendaciones.append({"tipo":"exito","icono":"🔥","titulo":"Contenido viral identificado",
                "texto":f"{n_virales} posts superan el 130% del promedio. Replicar el formato y temática de estos contenidos: \"{temas_virales}...\""})
        if tasa_resp_global < 60:
            recomendaciones.append({"tipo":"alerta","icono":"💬","titulo":"Tasa de respuesta insuficiente",
                "texto":f"Solo el {tasa_resp_global}% de los comentarios tienen respuesta. Se recomienda establecer turnos de atención y metas de respuesta del 80% en menos de 4 horas."})
        if avg_ctr > 0 and avg_ctr < 1:
            recomendaciones.append({"tipo":"alerta","icono":"🔗","titulo":"CTR bajo — mejorar call-to-action",
                "texto":f"El CTR promedio es {avg_ctr}%. Incluir llamados a la acción claros (\"Más información aquí\", \"Solicita tu crédito\"), usar botones de enlace y optimizar los primeros 3 segundos del contenido."})
        if avg_ctr >= 3:
            recomendaciones.append({"tipo":"exito","icono":"🎯","titulo":"Excelente CTR",
                "texto":f"CTR promedio de {avg_ctr}%, por encima del benchmark sectorial (1-2%). Continuar con este estilo de copys y formatos."})
        if avg_t_vision > 0 and avg_t_vision < 15:
            recomendaciones.append({"tipo":"alerta","icono":"⏱️","titulo":"Tiempo de visionado bajo en videos",
                "texto":f"Promedio de visionado: {avg_t_vision}s. Capturar la atención en los primeros 3 segundos, usar subtítulos y mantener videos cortos (30-60s) con mensaje directo."})
        if avg_t_vision >= 30:
            recomendaciones.append({"tipo":"exito","icono":"▶️","titulo":"Buen engagement en videos",
                "texto":f"Tiempo de visionado promedio: {avg_t_vision}s. El formato de video está funcionando bien. Aumentar la frecuencia de publicación de videos."})
        if total_nuevos_seg > 0:
            recomendaciones.append({"tipo":"info","icono":"👥","titulo":"Nuevos seguidores captados",
                "texto":f"{total_nuevos_seg} nuevos seguidores en el período. Para aumentar este número: publicar contenido educativo sobre ICETEX, hacer preguntas directas a la audiencia y usar hashtags relevantes."})
        if not recomendaciones:
            recomendaciones.append({"tipo":"info","icono":"📊","titulo":"Ingresa métricas completas",
                "texto":"Para generar recomendaciones detalladas, completa los campos de alcance, impresiones, CTR y tiempo de visionado en el formulario de ingreso de datos."})

    metricas_avanzadas = {
        "total_alcance":      total_alcance,
        "total_impresiones":  total_impresiones,
        "total_nuevos_seg":   total_nuevos_seg,
        "avg_ctr":            avg_ctr,
        "avg_t_vision_seg":   avg_t_vision,
        "tasa_interaccion":   tasa_interaccion_global,
        "total_interacciones":total_interacciones,
        "n_virales":          n_virales,
        "n_promedio":         n_promedio,
        "n_bajos":            n_bajos,
        "avg_score":          avg_score_r,
        "total_pubs":         total_pubs,
    }

    temas_s = dict(sorted(temas_c.items(), key=lambda x: x[1], reverse=True))
    conclusiones = generar_conclusiones(plataformas, sent_c, temas_s)

    return jsonify({
        "mes": mes, "anio": anio,
        "plataformas":       plataformas,
        "virales":           virales,
        "bajas":             bajas,
        "todos_pubs":        sorted(pubs_c, key=lambda x: x["viral_score"], reverse=True),
        "temas":             temas_s,
        "sentimiento":       sent_c,
        "conclusiones":      conclusiones,
        "metricas_avanzadas":metricas_avanzadas,
        "recomendaciones":   recomendaciones,
    })

init_db()

if __name__ == "__main__":
    ip = get_ip()
    print(f"\n  → http://{ip}:5000/formulario")
    print(f"  → http://{ip}:5000/admin")
    print(f"  → http://{ip}:5000/dashboard")
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
