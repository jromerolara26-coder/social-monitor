#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json, os, sys, subprocess, socket, re, unicodedata
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
# Patrones regex para tiempo de espera con números ("llevo 3 semanas", "hace 6 meses", etc.)
_PATRONES_ESPERA = [
    re.compile(r"llevo\s+\d+\s*(d[íi]as?|semanas?|meses?|a[ñn]os?)", re.IGNORECASE),
    re.compile(r"hace\s+\d+\s*(d[íi]as?|semanas?|meses?|a[ñn]os?)", re.IGNORECASE),
    re.compile(r"\d+\s*(d[íi]as?|semanas?|meses?|a[ñn]os?)\s+(esperando|sin\s+respuesta|sin\s+soluci)", re.IGNORECASE),
    re.compile(r"desde\s+hace\s+\d+", re.IGNORECASE),
]

_NEGACIONES_ES = frozenset({
    "no","nunca","jamás","jamas","tampoco","ni","ningún","ningun","ninguna",
    "nada","nadie","sin","imposible","difícil","dificil","apenas","todavía","todavia",
})

_EMOJIS_NEG = set("😡🤬😤👎😠😞😢😭😩😫💔🙄😑😒🤦🤷")
_EMOJIS_POS = set("❤😊😍🥰👍🙌🎉✅💯😃😄🤩🥳💪⭐🌟✨🙏")

def _norm_texto(texto: str) -> str:
    """Normaliza para comparación: minúsculas, sin tildes, sin puntuación, con espacios límite."""
    t = (texto or "").lower()
    # Eliminar tildes/diacríticos: café→cafe, caído→caido, ñ→n, etc.
    t = unicodedata.normalize("NFD", t)
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    t = re.sub(r"[^\w]", " ", t)
    return " " + re.sub(r"\s+", " ", t).strip() + " "

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
        "ladrón","ladron","sinvergüenza","sinverguenza","sinvergüenzas","sinverguenzas",
        "deberían cerrar","deberian cerrar","entidad corrupta","inútiles","inutiles",
        "no sirven para nada","para qué existen","para que existen",
        "roban a los estudiantes","roban a la gente","se roban la plata",
        "cómo es posible esto","como es posible esto",
        "denunciaré","denunciare","los voy a denunciar","les voy a poner queja",
        "acudiré","acudire a","buscaré un abogado","buscare un abogado",
        "reportaré esto","reportare esto","escalaré el caso","escalare el caso",
        "no les importamos","nos tienen abandonados","nos tienen de último","nos tienen de ultimo",
        "vergüenza de entidad","verguenza de entidad","esto no puede seguir así","esto no puede seguir asi",
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
        "demorado","retraso","retrasado","tardanza","lento","lenta",
        "mucha demora","tanta demora","la demora","hay demora","con demora",
        "hace días","hace dias","hace semanas","hace meses","hace un año","hace un mes",
        "llevo esperando","llevo dias","llevo semanas","llevo meses","llevo un mes",
        "sigo esperando","aún no","aun no","todavía no","todavia no",
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
        "desastre","un desastre","pesadilla","una pesadilla","ridículo","ridiculo",
        "ridícula","ridicula","caos","un caos","es un caos","qué caos","que caos",
        "pésima gestión","pesima gestion","mala gestión","mala gestion",
        "pésima experiencia","pesima experiencia","muy mala experiencia","mala experiencia",
        "sin noticias","no me contactaron","nunca me llamaron","no me notificaron",
        "imposible comunicarse","imposible contactarlos","imposible hablar con alguien",
        "me tienen esperando","nos tienen esperando","llevo días esperando","llevamos semanas",
        "llevamos meses","hace más de un mes","hace mas de un mes",
        "siguen sin responder","sigo sin respuesta","sigo sin solución","sigo sin solucion",
        "da error","da errores","falla constantemente","siempre falla",
        "nadie da información","nadie da informacion","sin información clara","sin claridad",
        "tanta demora","demasiada demora","muy deficiente","totalmente deficiente",
        "no cumplieron","no han cumplido","incumplieron","prometieron y no cumplieron",
        "qué asco","que asco de servicio","qué desastre","que desastre",
        "no funciona nada","todo está mal","todo esta mal","nada funciona",
        "me dejaron solo","me dejaron sola","sin acompañamiento","sin apoyo",
        "no hay respuesta","no tengo respuesta","sin obtener respuesta","no he recibido respuesta",
        "no me respondieron","aun sin respuesta","sin respuesta alguna",
        "perdi la beca","perdí la beca","perdi el beneficio","perdí el beneficio",
        "no me notifico","no me notificó","no nos notificaron","sin notificarme",
        "no puedo pagar","no pude pagar","no me deja pagar","no me dejan pagar",
        "hace 1 mes","hace 2 meses","hace 3 meses","hace 4 meses","hace 5 meses","hace 6 meses",
        "hace varios meses","hace bastante tiempo","desde hace meses","desde hace semanas",
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
        "recibí el desembolso","recibi el desembolso","recibí la beca","recibi la beca",
        "llegó la beca","llego la beca","me renovaron","me legalizaron","proceso exitoso",
        "sin complicaciones","me apoyaron","me asesoraron bien","fue aprobado","fue aprobada",
        "desembolsaron","el desembolso llegó","el desembolso llego",
    ],
}

def detectar_emocion_srv(texto: str) -> str:
    if not texto or not texto.strip():
        return "Neutro"

    t_norm = _norm_texto(texto)
    conteos = {"Crítico": 0.0, "Negativo": 0.0, "Positivo": 0.0}

    # Patrones de tiempo de espera con números ("llevo 3 semanas", "hace 6 meses", etc.)
    for patron in _PATRONES_ESPERA:
        if patron.search(texto):
            conteos["Negativo"] += 3
            break

    # Emojis en texto original (no se normalizan)
    for ch in texto:
        if ch in _EMOJIS_NEG:
            conteos["Negativo"] += 1.5
        elif ch in _EMOJIS_POS:
            conteos["Positivo"] += 1.5

    for em in ["Crítico", "Negativo", "Positivo"]:
        for frase in _EMOCIONES[em]:
            f_norm = _norm_texto(frase)         # conservar espacios límite para word-boundary
            if not f_norm.strip() or f_norm not in t_norm:
                continue
            # Frases más largas = mayor peso (más específicas y confiables)
            peso = max(1, len(f_norm.strip().split()))
            if em == "Positivo":
                # Si hay una negación justo antes del término positivo, se invierte a Negativo
                idx = t_norm.find(f_norm)
                ctx = t_norm[max(0, idx - 35): idx]
                if frozenset(ctx.split()) & _NEGACIONES_ES:
                    conteos["Negativo"] += peso
                else:
                    conteos["Positivo"] += peso
            else:
                conteos[em] += peso

    # Crítico recibe triple peso (casos graves deben prevalecer)
    conteos["Crítico"] *= 3

    mejor = max(conteos, key=conteos.get)
    return mejor if conteos[mejor] > 0 else "Neutro"

def reprocesar_emociones(datos: dict) -> dict:
    for plat, pubs in datos.items():
        for pub in pubs:
            for c in pub.get("comentarios", []):
                c["emocion"] = detectar_emocion_srv(c.get("texto", ""))
                c["tema"]    = detectar_tema_srv(c.get("texto", ""))
    return datos

# ── Detección automática de temas ICETEX
_TEMAS_ICETEX = {
    "Crédito / Préstamo": [
        "crédito","credito","préstamo","prestamo","financiación","financiacion",
        "financiamiento","solicitar crédito","solicitud de crédito","aplicar crédito",
        "modalidad","largo plazo","corto plazo","acces","icetex fácil","icetex facil",
        "tasa de interés","intereses","amortización","amortizacion","cuota mensual",
        "cuota del crédito","monto del crédito","valor del crédito",
    ],
    "Desembolso / Giro": [
        "desembolso","desembolsar","desembolsaron","no desembolsan","sin desembolso",
        "giro","giraron","no han girado","no me han girado","cuándo giran","cuando giran",
        "cuándo desembolsan","cuando desembolsan","pendiente de giro","pendiente desembolso",
        "transferencia","no llegó el dinero","no llego el dinero","no llegó la plata",
        "no llego la plata","plata no llega","dinero no llega","no llegó","no llego",
        "giro pendiente","giro no realizado","demora en el giro","demora desembolso",
    ],
    "Matrícula / Pago": [
        "matrícula","matricula","pago de matrícula","pagar matrícula","pago matricula",
        "factura","recibo de matrícula","liquidación","liquidacion","valor matrícula",
        "pago semestre","semestre","cuota matrícula","cobro matrícula","recibo pago",
        "aval de matrícula","aval matricula","fecha de pago","vencimiento pago",
    ],
    "Subsidio / Beca": [
        "subsidio","subsidios","beca","becas","apoyo económico","apoyo economico",
        "auxilio","beneficio económico","beneficio economico","gratuidad",
        "generación e","generacion e","ser pilo paga","fondos especiales",
        "fondo especial","beca de posgrado","beca exterior","beca internacional",
        "beca de excelencia","beneficio educativo","apoyo académico",
    ],
    "Documentos / Requisitos": [
        "documento","documentos","certificado","paz y salvo","constancia",
        "soporte","carta","estado de cuenta","extracto","comprobante",
        "documentación","documentacion","requisito","requisitos","cargar documentos",
        "subir documentos","adjuntar","adjunto","formulario","formularios",
        "documentos requeridos","documentos faltantes","falta documentos",
        "solicitar certificado","solicitar constancia","papeles","papelería",
    ],
    "Portal / Plataforma": [
        "portal","plataforma","página web","pagina web","sitio web","aplicativo",
        "app icetex","sistema","sistema caído","sistema caido","caído","caido",
        "error","error sistema","no carga","no funciona","falla","fallo","no abre",
        "login","contraseña","clave","usuario","acceso","no puedo ingresar",
        "no puedo entrar","no me deja entrar","problema con el portal",
        "página no carga","el portal no sirve","plataforma caída",
    ],
    "Refinanciación / Mora": [
        "refinanciación","refinanciacion","refinanciar","reestructurar","reestructuración",
        "reestructuracion","mora","moroso","morosa","vencida","vencido","deuda vencida",
        "cobro","cobros","jurídico","juridico","embargo","cobranza","cartera vencida",
        "carta de cobro","acuerdo de pago","plan de pagos","cuota atrasada","atraso",
        "atraso en pago","intereses de mora","sanción","sanciones",
    ],
    "Condonación": [
        "condonación","condonacion","condonar","perdonar deuda","exonerar",
        "exoneración","exoneracion","perdón de deuda","perdon de deuda",
        "cancelar deuda","quitar deuda","eliminar deuda","saldo cero",
        "quitar intereses","eliminar intereses","exonerar deuda","borrar deuda",
    ],
    "Atención al cliente": [
        "atención","atencion","servicio al cliente","asesor","asesora","agente",
        "call center","línea","linea","teléfono","telefono","chat","whatsapp",
        "no responden","no atienden","mala atención","demora","tarda","espera",
        "turno","nadie responde","me ignoraron","ignoraron","mala atención",
        "pésimo servicio","pesimo servicio","mal servicio","no me ayudan",
        "no dan solución","no dan solucion","sin respuesta","correo sin respuesta",
    ],
    "Posgrado / Exterior": [
        "posgrado","maestría","maestria","doctorado","especialización","especializacion",
        "postgrado","exterior","internacional","estudiar fuera","estudiar afuera",
        "otro país","otro pais","extranjero","beca exterior","crédito exterior",
        "estudio en el exterior","programa en el exterior",
    ],
    "Estado de solicitud": [
        "estado","en estudio","en revisión","revision","en proceso","proceso de estudio",
        "aprobado","aprobada","rechazado","rechazada","negado","negada","resultado",
        "cuándo me responden","cuando responden","respuesta solicitud","seguimiento",
        "radicado","radicacion","número de radicado","número de caso","consultar estado",
        "cómo saber","como saber","cómo va","como va mi solicitud","saber si fue aprobado",
    ],
    "Renovación / Legalización": [
        "renovación","renovacion","renovar","legalización","legalizacion","legalizar",
        "renovar crédito","actualizar datos","actualización","actualizacion",
        "próximo semestre","proximo semestre","continuar beneficio","continuar crédito",
        "renovar beca","renovar subsidio","renovar apoyo","mantener beneficio",
    ],
    "Cobro / Facturación": [
        "me están cobrando","me cobran","cobro incorrecto","cobro erróneo","cobro erroneo",
        "me cobran de más","me cobran de mas","error en cobro","cobro duplicado",
        "doble cobro","cobro sin aviso","cobro inesperado","factura incorrecta",
        "saldo incorrecto","saldo equivocado","no debo eso","no reconozco el cobro",
    ],
}

def detectar_tema_srv(texto: str) -> str:
    if not texto:
        return "Otros"
    t = texto.lower()
    scores = {}
    for tema, palabras in _TEMAS_ICETEX.items():
        score = sum(1 for p in palabras if p in t)
        if score > 0:
            scores[tema] = score
    if not scores:
        return "Otros"
    return max(scores, key=scores.get)

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
            fecha_carga  = str(row["fecha"])[:10] if row["fecha"] else ""
            analista_row = row["analista"] or "Anónimo"
            if row["sla"]:
                s = json.loads(row["sla"])
                sla_final.update(s)
            for plat in ["facebook", "instagram", "twitter", "linkedin", "tiktok", "youtube"]:
                for pub in d.get(plat, []):
                    pub["fecha_carga"]  = fecha_carga
                    pub["analista_carga"] = analista_row
                consolidado[plat] += d.get(plat, [])
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

def _pub_fingerprint(pub: dict) -> str:
    """Clave única de una publicación: URL si existe, si no los primeros 100 chars del texto."""
    url = (pub.get("url") or "").strip()
    if url and url.startswith("http"):
        return url.lower()
    return (pub.get("texto") or "").strip().lower()[:100]

def _fingerprints_existentes() -> set:
    """Devuelve el conjunto de fingerprints de todas las publicaciones ya guardadas."""
    conn = get_db()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT datos FROM envios")
            rows = cur.fetchall()
    finally:
        conn.close()
    fps = set()
    for row in rows:
        try:
            d = json.loads(row["datos"])
            for pubs in d.values():
                for pub in pubs:
                    fp = _pub_fingerprint(pub)
                    if fp:
                        fps.add(fp)
        except Exception:
            pass
    return fps

@app.route("/enviar", methods=["POST"])
def enviar():
    body = request.get_json(force=True, silent=True) or {}
    meta      = body.get("_meta", {})
    analista  = meta.get("analista", "Anónimo")
    fb_cuenta = meta.get("fb_cuenta", "")
    ig_cuenta = meta.get("ig_cuenta", "")
    sla       = json.dumps(meta.get("sla", {}))
    forzar    = body.get("_forzar", False)
    datos_raw = {
        "twitter":   body.get("twitter",   []),
        "facebook":  body.get("facebook",  []),
        "instagram": body.get("instagram", []),
        "linkedin":  body.get("linkedin",  []),
        "tiktok":    body.get("tiktok",    []),
        "youtube":   body.get("youtube",   []),
    }

    # ── Detección de duplicados (omitir si el usuario forzó el envío)
    if not forzar:
        existentes = _fingerprints_existentes()
        duplicados = []
        for plat, pubs in datos_raw.items():
            for pub in pubs:
                fp = _pub_fingerprint(pub)
                if fp and fp in existentes:
                    label = pub.get("url") or (pub.get("texto", "")[:60] + "…")
                    duplicados.append({"plataforma": plat, "label": label})
        if duplicados:
            return jsonify({
                "ok":        False,
                "duplicados": duplicados,
                "mensaje":   f"⚠️ {len(duplicados)} publicación(es) ya fueron registradas anteriormente.",
            }), 409

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
        temas_detalle = {}
        sent  = {"Positivo": 0, "Negativo": 0, "Neutro": 0, "Crítico": 0}
        sent_msgs = {"Positivo": [], "Negativo": [], "Neutro": [], "Crítico": []}
        ranking_usuarios = {}
        for row in filas:
            try:
                d = json.loads(row["datos"])
            except Exception:
                continue
            for plat, pub_list in d.items():
                for pub in pub_list:
                    coms    = pub.get("comentarios", [])
                    n_coms  = len(coms)
                    n_resp  = sum(1 for c in coms if c.get("respondido"))
                    tiempos = [c["tiempo_respuesta_min"] for c in coms
                               if c.get("tiempo_respuesta_min") is not None]
                    if plat not in plats:
                        plats[plat] = {"total_comentarios": 0, "respondidos": 0,
                                       "publicaciones": 0, "tiempos": [],
                                       "alcance": 0, "impresiones": 0,
                                       "nuevos_seg": 0, "ctrs": [], "t_vision": [],
                                       "likes": 0, "shares": 0, "guardados": 0}
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
                    interacciones    = likes + shares + n_coms + guardados
                    tasa_interaccion = round(interacciones / alcance * 100, 2) if alcance > 0 else 0
                    viral_score = (n_coms * 3) + (shares * 2) + (likes * 1) + (guardados * 2)
                    riesgo      = round((100 - tasa) * 0.7 + max(0, 10 - n_coms) * 3, 1)
                    plats[plat]["alcance"]     = plats[plat].get("alcance", 0) + alcance
                    plats[plat]["impresiones"] = plats[plat].get("impresiones", 0) + impresiones
                    plats[plat]["nuevos_seg"]  = plats[plat].get("nuevos_seg", 0) + nuevos_seg
                    plats[plat]["likes"]       = plats[plat].get("likes", 0) + likes
                    plats[plat]["shares"]      = plats[plat].get("shares", 0) + shares
                    plats[plat]["guardados"]   = plats[plat].get("guardados", 0) + guardados
                    if ctr > 0:   plats[plat]["ctrs"].append(ctr)
                    if t_vision > 0: plats[plat]["t_vision"].append(t_vision)
                    pubs_list.append({
                        "plataforma":            plat,
                        "texto":                 pub.get("texto", ""),
                        "url":                   pub.get("url", ""),
                        "fecha":                 pub.get("fecha", "")[:10] if pub.get("fecha") else "",
                        "n_coms":                n_coms,
                        "n_resp":                n_resp,
                        "likes":                 likes,
                        "shares":                shares,
                        "guardados":             guardados,
                        "alcance":               alcance,
                        "impresiones":           impresiones,
                        "ctr":                   ctr,
                        "nuevos_seguidores":     nuevos_seg,
                        "tiempo_visionado_seg":  t_vision,
                        "tasa":                  tasa,
                        "tasa_interaccion":      tasa_interaccion,
                        "viral_score":           viral_score,
                        "riesgo":                riesgo,
                    })
                    for c in coms:
                        texto_c = c.get("texto", "")
                        tema_c  = detectar_tema_srv(texto_c) if texto_c.strip() else (c.get("tema") or "Otros")
                        em      = c.get("emocion", "Neutro")
                        temas[tema_c] = temas.get(tema_c, 0) + 1
                        sent[em]      = sent.get(em, 0) + 1
                        # Mensajes por sentimiento (max 20 por categoría)
                        if len(sent_msgs.get(em, [])) < 20:
                            sent_msgs.setdefault(em, []).append({
                                "usuario":    c.get("usuario", "Usuario"),
                                "texto":      c.get("texto", ""),
                                "plataforma": plat,
                                "fecha":      c.get("fecha", "")[:10] if c.get("fecha") else "",
                                "respondido": c.get("respondido", False),
                                "texto_resp": c.get("texto_respuesta", "") or "",
                                "pub_url":    pub.get("url", ""),
                                "pub_texto":  pub.get("texto", "")[:70],
                            })
                        # Ranking de usuarios
                        usr = c.get("usuario") or "Desconocido"
                        if usr not in ranking_usuarios:
                            ranking_usuarios[usr] = {"total": 0, "respondidos": 0,
                                                     "plataformas": [], "emociones": []}
                        ranking_usuarios[usr]["total"] += 1
                        if c.get("respondido"):
                            ranking_usuarios[usr]["respondidos"] += 1
                        if plat not in ranking_usuarios[usr]["plataformas"]:
                            ranking_usuarios[usr]["plataformas"].append(plat)
                        ranking_usuarios[usr]["emociones"].append(em)
                        # Detalle por tema
                        if tema_c not in temas_detalle:
                            temas_detalle[tema_c] = {
                                "count": 0,
                                "sent": {"Positivo":0,"Negativo":0,"Crítico":0,"Neutro":0},
                                "ejemplos": [],
                                "respondidos": 0,
                            }
                        temas_detalle[tema_c]["count"] += 1
                        temas_detalle[tema_c]["sent"][em] = temas_detalle[tema_c]["sent"].get(em, 0) + 1
                        if c.get("respondido"): temas_detalle[tema_c]["respondidos"] += 1
                        if len(temas_detalle[tema_c]["ejemplos"]) < 5:
                            temas_detalle[tema_c]["ejemplos"].append({
                                "texto":      c.get("texto", ""),
                                "emocion":    em,
                                "plataforma": plat,
                                "usuario":    c.get("usuario", ""),
                                "pub_url":    pub.get("url", ""),
                                "pub_texto":  pub.get("texto", "")[:60],
                                "respondido": c.get("respondido", False),
                                "texto_resp": c.get("texto_respuesta", "") or "",
                            })
        ranking_list = sorted([
            {"usuario": u, "total": d["total"], "respondidos": d["respondidos"],
             "sin_resp": d["total"] - d["respondidos"],
             "plataformas": d["plataformas"],
             "tasa_resp": round(d["respondidos"] / d["total"] * 100) if d["total"] else 0,
             "emocion_dom": max(set(d["emociones"]), key=d["emociones"].count) if d["emociones"] else "Neutro"}
            for u, d in ranking_usuarios.items()
        ], key=lambda x: x["total"], reverse=True)[:20]
        return plats, pubs_list, temas, sent, sent_msgs, temas_detalle, ranking_list

    curr, prev = [], []
    for row in rows:
        m, a = parse_mes_anio(row["fecha"])
        if m == mes and a == anio:
            curr.append(row)
        elif m == mes_ant and a == anio_ant:
            prev.append(row)
        elif m is None:
            curr.append(row)

    plats_c, pubs_c, temas_c, sent_c, sent_msgs_c, temas_det_c, ranking_c = extraer(curr)
    plats_p, _, _, _, _, _, _                                              = extraer(prev)

    plataformas = {}
    for plat, st in plats_c.items():
        total      = st["total_comentarios"]
        resp       = st["respondidos"]
        tms        = st["tiempos"]
        ctrs_p     = st.get("ctrs", [])
        tv_p       = st.get("t_vision", [])
        prev_total = plats_p.get(plat, {}).get("total_comentarios", 0)
        evol       = round((total - prev_total) / prev_total * 100, 1) if prev_total > 0 else 0
        tasa_r     = round(resp / total * 100, 1) if total > 0 else 0
        alcance_p  = st.get("alcance", 0)
        inter_p    = st.get("likes",0)+st.get("shares",0)+total+st.get("guardados",0)
        plataformas[plat] = {
            "total_comentarios":   total,
            "respondidos":         resp,
            "no_respondidos":      total - resp,
            "publicaciones":       st["publicaciones"],
            "tasa_respuesta":      tasa_r,
            "tiempo_promedio_min": round(sum(tms) / len(tms), 0) if tms else None,
            "comentarios_mes_ant": prev_total,
            "evolucion_pct":       evol,
            "alcance":             alcance_p,
            "impresiones":         st.get("impresiones", 0),
            "likes":               st.get("likes", 0),
            "shares":              st.get("shares", 0),
            "guardados":           st.get("guardados", 0),
            "nuevos_seguidores":   st.get("nuevos_seg", 0),
            "avg_ctr":             round(sum(ctrs_p)/len(ctrs_p),2) if ctrs_p else 0,
            "avg_t_vision":        round(sum(tv_p)/len(tv_p),0) if tv_p else 0,
            "tasa_interaccion":    round(inter_p / alcance_p * 100, 2) if alcance_p > 0 else 0,
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
        "temas_detalle":     temas_det_c,
        "sentimiento":       sent_c,
        "sent_msgs":         sent_msgs_c,
        "ranking_usuarios":  ranking_c,
        "conclusiones":      conclusiones,
        "metricas_avanzadas":metricas_avanzadas,
        "recomendaciones":   recomendaciones,
    })

_FRASES_GENERICAS_V = [
    "gracias por contactarnos","gracias por escribirnos","gracias por comunicarte",
    "te invitamos a ingresar","visita nuestra página","ingresa a",
    "por favor escríbenos","escríbenos al correo","comunícate al","llámanos al",
    "lamentamos el inconveniente","entendemos tu situación",
    "con gusto te ayudamos","estamos para servirte","quedamos atentos",
    "bienvenido","hola! gracias","hola, gracias","con mucho gusto",
    "en breve","a la mayor brevedad","nos pondremos en contacto",
    "mediante mensaje directo","por mensaje privado","por mp","al número",
]
_INDICADORES_RESOLUCION_V = [
    "hemos revisado","hemos verificado","revisamos tu caso","verificamos tu caso",
    "tu caso fue","procedimos a","ya fue procesado","ya fue actualizado","ya fue corregido",
    "el desembolso","el subsidio","la beca","el crédito","tu solicitud","tu radicado",
    "número de radicado","el proceso de","los requisitos son","los pasos son",
    "debes","debes ingresar","debes adjuntar","el plazo es","la fecha límite",
    "tienes hasta","te informamos que","te comunicamos que","te confirmamos que",
    "fue aprobado","fue aprobada","está en proceso","en revisión","en trámite",
]
_STOPWORDS_V = {"el","la","los","las","un","una","y","o","de","en","que","a","por","con",
                "su","al","del","es","se","me","te","nos","hay","para","pero","como",
                "más","mas","muy","ya","si","no","mi","tu","le","lo"}

def evaluar_calidad_api(comentario: str, respuesta: str) -> dict:
    if not respuesta or not respuesta.strip():
        return {"score":0,"nivel":"Sin respuesta","motivo":"ICETEX no respondió","relevancia_pct":0,
                "sugerencia":"Responder el comentario del usuario lo antes posible."}
    c = comentario.lower()
    r = respuesta.lower()
    longitud      = len(respuesta.strip())
    n_genericas   = sum(1 for f in _FRASES_GENERICAS_V if f in r)
    n_resolucion  = sum(1 for f in _INDICADORES_RESOLUCION_V if f in r)
    palabras_com  = {w for w in c.split() if len(w) > 3 and w not in _STOPWORDS_V}
    n_palabras_en_resp = sum(1 for p in palabras_com if p in r)
    relevancia_pct = round(n_palabras_en_resp / len(palabras_com) * 100, 1) if palabras_com else 0
    redirige = any(f in r for f in ["escríbenos al","escribenos al","por mensaje privado",
                                     "mediante mensaje directo","al correo","al número","por mp"])
    if n_genericas >= 2 and n_resolucion == 0 and n_palabras_en_resp < 2:
        return {"score":1,"nivel":"Genérica","relevancia_pct":relevancia_pct,
                "motivo":"Respuesta automática: no aborda el tema del usuario.",
                "sugerencia":"Personalizar la respuesta mencionando el tema específico del comentario."}
    if redirige and n_resolucion == 0 and longitud < 200:
        return {"score":1,"nivel":"Genérica","relevancia_pct":relevancia_pct,
                "motivo":"Solo redirige a otro canal sin dar información útil.",
                "sugerencia":"Dar al menos una respuesta parcial antes de redirigir al usuario."}
    if n_resolucion >= 2 and n_palabras_en_resp >= 3 and longitud > 200:
        return {"score":4,"nivel":"Excelente","relevancia_pct":relevancia_pct,
                "motivo":f"Respuesta resuelve y confirma acción tomada. Relevancia: {relevancia_pct}%.",
                "sugerencia":"Mantener este nivel de detalle y personalización."}
    if (n_palabras_en_resp >= 3 or n_resolucion >= 1) and longitud > 150:
        return {"score":3,"nivel":"Específica","relevancia_pct":relevancia_pct,
                "motivo":f"Respuesta aborda el tema del usuario. Relevancia: {relevancia_pct}%.",
                "sugerencia":"Agregar confirmación de la acción tomada para alcanzar nivel Excelente."}
    if n_palabras_en_resp >= 2 or (longitud > 200 and n_genericas == 0):
        return {"score":3,"nivel":"Específica","relevancia_pct":relevancia_pct,
                "motivo":f"Respuesta personalizada. Relevancia: {relevancia_pct}%.",
                "sugerencia":"Mencionar explícitamente la solución o próximo paso para el usuario."}
    if n_genericas >= 1 and longitud < 120:
        return {"score":1,"nivel":"Genérica","relevancia_pct":relevancia_pct,
                "motivo":"Respuesta corta y genérica sin abordar la solicitud.",
                "sugerencia":"Ampliar la respuesta abordando el tema puntual del comentario."}
    return {"score":2,"nivel":"Parcial","relevancia_pct":relevancia_pct,
            "motivo":f"Reconoce el tema pero no da solución concreta. Relevancia: {relevancia_pct}%.",
            "sugerencia":"Incluir pasos concretos o información específica que resuelva la consulta."}

@app.route("/api/validar-calidad", methods=["POST"])
def api_validar_calidad():
    body       = request.get_json(force=True, silent=True) or {}
    comentario = body.get("comentario", "")
    respuesta  = body.get("respuesta", "")
    if not comentario:
        return jsonify({"error": "Falta el comentario"}), 400
    return jsonify(evaluar_calidad_api(comentario, respuesta))

init_db()

if __name__ == "__main__":
    ip = get_ip()
    print(f"\n  → http://{ip}:5000/formulario")
    print(f"  → http://{ip}:5000/admin")
    print(f"  → http://{ip}:5000/dashboard")
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
