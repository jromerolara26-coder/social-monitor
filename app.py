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

    total_pub  = sum(p.get("publicaciones", 0) for p in plataformas.values())
    total_resp = sum(p.get("respondidas", 0) for p in plataformas.values())
    if total_pub > 0:
        tasa = total_resp / total_pub * 100
        if tasa < 50:
            conclusiones.append({"tipo": "alerta", "texto": f"La tasa de respuesta global es del {tasa:.0f}%, por debajo del mínimo recomendado (70%). Se sugiere reforzar los turnos de atención digital."})
        elif tasa >= 80:
            conclusiones.append({"tipo": "exito", "texto": f"Excelente tasa de respuesta global del {tasa:.0f}%. La entidad mantiene una atención digital oportuna y efectiva."})
        else:
            conclusiones.append({"tipo": "info", "texto": f"La tasa de respuesta global es del {tasa:.0f}%. Existe oportunidad de mejora para alcanzar el objetivo del 80%."})

    for plat, m in plataformas.items():
        crec = m.get("crecimiento_pct", 0)
        if crec > 10:
            conclusiones.append({"tipo": "exito", "texto": f"{plat} muestra un crecimiento de seguidores del {crec:.1f}% frente al mes anterior, indicando mayor visibilidad institucional."})
        elif crec < -2:
            conclusiones.append({"tipo": "alerta", "texto": f"{plat} registra una disminución del {abs(crec):.1f}% en seguidores. Se recomienda revisar la frecuencia y relevancia del contenido publicado."})

    total_sent = sum(sentimiento.values())
    if total_sent > 0:
        neg  = (sentimiento.get("Negativo", 0) + sentimiento.get("Crítico", 0)) / total_sent * 100
        pos  = sentimiento.get("Positivo", 0) / total_sent * 100
        crit = sentimiento.get("Crítico", 0) / total_sent * 100
        if crit > 20:
            conclusiones.append({"tipo": "alerta", "texto": f"El {crit:.0f}% de los comentarios son críticos. Se recomienda activar protocolo de gestión de crisis digital y revisar casos urgentes."})
        elif neg > 40:
            conclusiones.append({"tipo": "alerta", "texto": f"El {neg:.0f}% de los comentarios presentan tono negativo. Se sugiere reforzar la calidad de las respuestas y reducir tiempos de atención."})
        elif pos > 60:
            conclusiones.append({"tipo": "exito", "texto": f"El {pos:.0f}% de los comentarios son positivos, reflejando una percepción favorable de la entidad en redes sociales."})

    if temas:
        top = max(temas, key=temas.get)
        conclusiones.append({"tipo": "info", "texto": f"El tema con mayor volumen de menciones es '{top}' con {temas[top]} interacciones. Se recomienda priorizar contenido informativo sobre este tema."})

    engs = [(plat, m.get("engagement_rate", 0)) for plat, m in plataformas.items() if m.get("engagement_rate", 0) > 0]
    if engs:
        mejor = max(engs, key=lambda x: x[1])
        conclusiones.append({"tipo": "info", "texto": f"{mejor[0]} registra el mayor engagement rate con {mejor[1]:.1f}%, siendo la plataforma con mayor interacción relativa del período."})

    if not conclusiones:
        conclusiones.append({"tipo": "info", "texto": "Ingresa métricas para este período para generar conclusiones automáticas."})

    return conclusiones

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
    consolidado = {"twitter": [], "facebook": [], "instagram": []}
    analistas   = []
    sla_final   = {"facebook": 8, "instagram": 12, "twitter": 4}
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
    datos_str = json.dumps({
        "twitter":   body.get("twitter",   []),
        "facebook":  body.get("facebook",  []),
        "instagram": body.get("instagram", []),
    }, ensure_ascii=False)
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
        return "<h2 style='font-family:sans-serif;padding:40px'>Aún no se ha generado el reporte.</h2>"
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

    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT * FROM metricas_plataforma WHERE mes=%s AND anio=%s", (mes, anio))
            met_rows = cur.fetchall()

            cur.execute("""
                SELECT * FROM publicaciones_dashboard
                WHERE mes=%s AND anio=%s
                ORDER BY (likes+comentarios+compartidos+guardados) DESC
            """, (mes, anio))
            pub_rows = cur.fetchall()

            cur.execute("""
                SELECT * FROM temas_dashboard WHERE mes=%s AND anio=%s
                ORDER BY menciones DESC
            """, (mes, anio))
            tema_rows = cur.fetchall()

            cur.execute("SELECT datos FROM envios")
            envio_rows = cur.fetchall()
    finally:
        conn.close()

    plataformas = {}
    for row in met_rows:
        plat = row["plataforma"]
        seg  = row["seguidores"] or 0
        ant  = row["seguidores_ant"] or 0
        total_int = (row["likes"] or 0) + (row["comentarios"] or 0) + (row["compartidos"] or 0) + (row["guardados"] or 0)
        vis  = row["visualizaciones"] or 0
        eng  = round(total_int / vis * 100, 2) if vis > 0 else (round(total_int / seg * 100, 2) if seg > 0 else 0)
        crec = round((seg - ant) / ant * 100, 1) if ant > 0 else 0
        pub  = row["publicaciones"] or 0
        resp = row["respondidas"] or 0
        plataformas[plat] = {
            "seguidores":       seg,
            "seguidores_ant":   ant,
            "crecimiento_pct":  crec,
            "visualizaciones":  vis,
            "likes":            row["likes"] or 0,
            "comentarios":      row["comentarios"] or 0,
            "compartidos":      row["compartidos"] or 0,
            "guardados":        row["guardados"] or 0,
            "publicaciones":    pub,
            "respondidas":      resp,
            "tasa_respuesta":   round(resp / pub * 100, 1) if pub > 0 else 0,
            "alcance_ext_pct":  row["alcance_ext_pct"] or 0,
            "engagement_rate":  eng,
            "sentimiento": {
                "Positivo": row["sent_pos"] or 0,
                "Negativo": row["sent_neg"] or 0,
                "Neutro":   row["sent_neu"] or 0,
                "Crítico":  row["sent_crit"] or 0,
            },
        }

    virales, bajas, normales = [], [], []
    for pub in pub_rows:
        seg_plat  = plataformas.get(pub["plataforma"], {}).get("seguidores", 1) or 1
        vis       = pub["visualizaciones"] or 0
        total_int = (pub["likes"] or 0) + (pub["comentarios"] or 0) + (pub["compartidos"] or 0) + (pub["guardados"] or 0)
        base      = vis if vis > 0 else seg_plat
        eng       = round(total_int / base * 100, 2) if base > 0 else 0
        alc_pct   = round((pub["alcance"] or 0) / seg_plat * 100, 1) if seg_plat > 0 else 0
        entry = {
            "id": pub["id"], "plataforma": pub["plataforma"], "texto": pub["texto"] or "",
            "likes": pub["likes"] or 0, "comentarios": pub["comentarios"] or 0,
            "compartidos": pub["compartidos"] or 0, "guardados": pub["guardados"] or 0,
            "visualizaciones": vis, "engagement_rate": eng,
            "alcance_no_seg_pct": alc_pct, "tipo": pub["tipo"],
        }
        if pub["tipo"] == "viral":
            virales.append(entry)
        elif pub["tipo"] == "bajo":
            bajas.append(entry)
        else:
            normales.append(entry)

    normales_sorted = sorted(normales, key=lambda x: x["engagement_rate"], reverse=True)
    if not virales:
        virales = normales_sorted[:3]
    if not bajas:
        bajas = normales_sorted[-3:][::-1] if len(normales_sorted) > 3 else []

    temas = {row["tema"]: row["menciones"] for row in tema_rows}

    sentimiento = {"Positivo": 0, "Negativo": 0, "Neutro": 0, "Crítico": 0}
    for row in envio_rows:
        try:
            datos = json.loads(row["datos"])
            for plat_pubs in datos.values():
                for pub in plat_pubs:
                    for com in pub.get("comentarios", []):
                        em = com.get("emocion", "Neutro")
                        sentimiento[em] = sentimiento.get(em, 0) + 1
        except Exception:
            pass
    for plat_data in plataformas.values():
        for k, v in plat_data.get("sentimiento", {}).items():
            sentimiento[k] = sentimiento.get(k, 0) + v

    conclusiones = generar_conclusiones(plataformas, sentimiento, temas)

    return jsonify({
        "mes": mes, "anio": anio,
        "plataformas": plataformas,
        "virales": virales[:5],
        "bajas": bajas[:5],
        "temas": temas,
        "sentimiento": sentimiento,
        "conclusiones": conclusiones,
    })

init_db()

if __name__ == "__main__":
    ip = get_ip()
    print(f"\n  → http://{ip}:5000/formulario")
    print(f"  → http://{ip}:5000/admin")
    print(f"  → http://{ip}:5000/dashboard")
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
