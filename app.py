#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Social Monitor — Servidor web local
Ejecutar:  python app.py
Acceder:   http://TU_IP:5000
"""

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
        conn.commit()
    finally:
        conn.close()

# ── Helpers
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

# ── Rutas
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
    meta    = body.get("_meta", {})
    analista = meta.get("analista", "Anónimo")
    fb_cuenta = meta.get("fb_cuenta", "")
    ig_cuenta = meta.get("ig_cuenta", "")
    sla      = json.dumps(meta.get("sla", {}))
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
            n_fb = len(d.get("facebook",  []))
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

    # Actualizar config con SLA
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

    # Guardar JSON consolidado
    with open(JSON_OUT, "w", encoding="utf-8") as f:
        json.dump(consolidado, f, ensure_ascii=False, indent=2, default=str)

    # Generar reporte
    monitor_py = os.path.join(BASE, "monitor.py")
    r = subprocess.run([sys.executable, monitor_py, "--desde-json"],
                       capture_output=True, text=True)
    if r.returncode == 0:
        total = sum(len(v) for v in consolidado.values())
        total_coms = sum(len(c["comentarios"]) for plat in consolidado.values() for c in plat)
        return jsonify({
            "ok": True,
            "mensaje": f"Reporte generado con {total} publicaciones y {total_coms} comentarios de: {', '.join(analistas)}"
        })
    return jsonify({"ok": False, "mensaje": f"Error: {r.stderr[:300]}"})

@app.route("/reporte")
def reporte():
    if not os.path.exists(HTML_OUT):
        return "<h2 style='font-family:sans-serif;padding:40px'>Aún no se ha generado el reporte. Ve al panel de administrador.</h2>"
    return send_file(HTML_OUT)

@app.route("/descargar-excel")
def descargar_excel():
    if not os.path.exists(XLSX_OUT):
        return "Excel no disponible", 404
    return send_file(XLSX_OUT, as_attachment=True)

init_db()  # Siempre inicializar al importar (necesario para gunicorn)

if __name__ == "__main__":
    ip = get_ip()
    print("=" * 58)
    print("  SOCIAL MONITOR — Servidor Web")
    print("=" * 58)
    print(f"\n  Formulario para el equipo:")
    print(f"  → http://{ip}:5000/formulario")
    print(f"\n  Panel de administrador:")
    print(f"  → http://{ip}:5000/admin")
    print(f"\n  Reporte:")
    print(f"  → http://{ip}:5000/reporte")
    print(f"\n  Comparte la URL del formulario con tu equipo.")
    print(f"  Presiona Ctrl+C para detener el servidor.\n")
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
