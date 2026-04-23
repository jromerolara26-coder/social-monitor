#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ingreso manual de publicaciones y comentarios reales.
- Usa hora real del comentario y respuesta (DD/MM/AAAA HH:MM)
- Calcula tiempo de respuesta automáticamente
- Permite configurar SLA por plataforma
"""

import json
import os
import sys
import subprocess
from datetime import datetime, timezone, timedelta

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

BASE     = os.path.dirname(os.path.abspath(__file__))
JSON_OUT = os.path.join(BASE, "datos_social_icetex.json")

# ──────────────────────────────────────────────────────────
# Helpers de entrada
# ──────────────────────────────────────────────────────────
def leer(prompt: str, default: str = "") -> str:
    try:
        val = input(prompt).strip()
        return val if val else default
    except (EOFError, KeyboardInterrupt):
        print(); return default

def leer_int(prompt: str, default: int = 0) -> int:
    try:
        val = input(prompt).strip()
        return int(val) if val else default
    except (ValueError, EOFError):
        return default

def leer_float(prompt: str, default: float = 0.0) -> float:
    try:
        val = input(prompt).strip()
        return float(val) if val else default
    except (ValueError, EOFError):
        return default

def leer_fecha(prompt: str) -> datetime:
    """Pide una fecha y hora. Acepta DD/MM/AAAA HH:MM o DD/MM/AAAA."""
    while True:
        val = input(prompt).strip()
        if not val:
            return datetime.now(timezone.utc)
        for fmt in ["%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y"]:
            try:
                dt = datetime.strptime(val, fmt)
                return dt.replace(tzinfo=timezone.utc)
            except ValueError:
                continue
        print("  ⚠️  Formato inválido. Ejemplo correcto:  21/04/2026 14:30")

def sep(char="─", n=54):
    print(char * n)

# ──────────────────────────────────────────────────────────
# Clasificador de temas
# ──────────────────────────────────────────────────────────
TEMAS_KW = {
    "crédito / pago":            ["crédito","credito","préstamo","prestamo","cuota","pago","deuda","saldo","mora","cobro","factura","abono"],
    "beca / condonación":        ["beca","condonación","condonacion","generación e","generacion e","subsidio","perdí la beca"],
    "trámites / portal":         ["portal","trámite","tramite","solicitud","formulario","sistema","acceso","contraseña","login"],
    "atención al cliente":       ["atención","atencion","servicio","demora","espera","respuesta","chat","línea","linea","asesor"],
    "requisitos / elegibilidad": ["requisito","elegibilidad","aplica","cumplir","condición","condicion","puntaje","sisbén","sisben","requisito"],
    "información general":       ["información","informacion","cómo","como","cuándo","cuando","dónde","donde","qué","que","cuál","cual"],
    "reclamo / queja":           ["reclamo","queja","problema","error","fallo","mal","pésimo","pesimo","terrible","injusto","cancelaron"],
    "posgrado / internacional":  ["posgrado","maestría","maestria","doctorado","exterior","internacional","convenio"],
}

def clasificar(texto: str) -> str:
    t = texto.lower()
    for tema, kws in TEMAS_KW.items():
        if any(k in t for k in kws):
            return tema
    return "otro"

def evaluar_calidad(respuesta: str) -> dict:
    if not respuesta:
        return {"score": 0, "nivel": "Sin respuesta", "motivo": "No hubo respuesta"}
    palabras = len(respuesta.split())
    if palabras >= 25:
        return {"score": 3, "nivel": "Específica",  "motivo": "Respuesta detallada"}
    if palabras >= 10:
        return {"score": 2, "nivel": "Parcial",     "motivo": "Respuesta presente"}
    return     {"score": 1, "nivel": "Genérica",    "motivo": "Respuesta muy corta"}

# ──────────────────────────────────────────────────────────
# Ingreso de comentarios
# ──────────────────────────────────────────────────────────
def ingresar_comentarios(plataforma: str, sla_horas: float) -> list:
    comentarios = []
    print()
    print("  Ingresa los comentarios de esta publicación.")
    print("  → Deja el campo 'Usuario' en blanco para terminar.\n")

    while True:
        sep("·", 44)
        usuario = leer("  Nombre/usuario de quien comentó (Enter = terminar): ")
        if not usuario:
            break

        texto = leer("  Texto del comentario: ")
        if not texto:
            print("  (sin texto, omitido)")
            continue

        print("  Fecha y hora del comentario  (DD/MM/AAAA HH:MM)", end="")
        print("  → Ejemplo: 22/04/2026 09:15")
        fecha_com = leer_fecha("  Fecha comentario: ")

        respondido_s = leer("  ¿Respondiste este comentario? (s/n): ", "n").lower()
        respondido   = respondido_s.startswith("s")

        t_resp_min = None
        txt_resp   = None
        cumple_sla = None

        if respondido:
            print("  Fecha y hora de TU respuesta (DD/MM/AAAA HH:MM)")
            fecha_resp = leer_fecha("  Fecha respuesta: ")
            txt_resp   = leer("  Texto de tu respuesta: ")

            diff = (fecha_resp - fecha_com).total_seconds() / 60
            if diff < 0:
                print("  ⚠️  La respuesta es anterior al comentario. Revisa las fechas.")
                diff = abs(diff)
            t_resp_min = round(diff, 1)
            cumple_sla = t_resp_min <= sla_horas * 60

            hh = int(t_resp_min // 60)
            mm = int(t_resp_min % 60)
            estado = "✅ dentro del SLA" if cumple_sla else "❌ fuera del SLA"
            print(f"  ⏱️  Tiempo de respuesta: {hh}h {mm}min  →  {estado}")

        comentarios.append({
            "usuario":              usuario,
            "texto":                texto,
            "fecha":                fecha_com.isoformat(),
            "respondido":           respondido,
            "tiempo_respuesta_min": t_resp_min,
            "texto_respuesta":      txt_resp,
            "tema":                 clasificar(texto),
            "calidad":              evaluar_calidad(txt_resp or ""),
            "cumple_sla":           cumple_sla,
        })
        print(f"  ✓ Comentario de '{usuario}' guardado  (tema: {clasificar(texto)})")

    return comentarios

# ──────────────────────────────────────────────────────────
# Ingreso de publicaciones por plataforma
# ──────────────────────────────────────────────────────────
def ingresar_plataforma(plataforma: str, nombre_cuenta: str, sla_horas: float) -> list:
    publicaciones = []
    icono = "📘" if plataforma == "facebook" else "📸"

    sep("═")
    print(f"{icono}  {plataforma.upper()} — Cuenta: @{nombre_cuenta}  |  SLA: {sla_horas}h")
    sep("═")

    n = leer_int(f"¿Cuántas publicaciones de {plataforma} quieres ingresar? (0 para omitir): ", 0)
    if n == 0:
        return []

    for i in range(1, n + 1):
        sep("─")
        print(f"  Publicación {i} de {n}")
        sep("─", 40)

        texto  = leer("  Texto/descripción del post: ")
        print("  Fecha y hora de la publicación (DD/MM/AAAA HH:MM)  → Ejemplo: 20/04/2026 10:00")
        fecha_pub = leer_fecha("  Fecha publicación: ")
        likes  = leer_int("  Likes / reacciones: ", 0)
        shares = leer_int("  Compartidos: ", 0)
        url    = leer(f"  URL del post (opcional, Enter para omitir): ",
                      f"https://www.{plataforma}.com/{nombre_cuenta}")

        comentarios = ingresar_comentarios(plataforma, sla_horas)

        post_id = f"{plataforma[:2].upper()}_{i:03d}_{int(fecha_pub.timestamp())}"
        publicaciones.append({
            "id":            post_id,
            "texto":         texto or "(Sin texto)",
            "fecha":         fecha_pub.isoformat(),
            "likes":         likes,
            "shares":        shares,
            "replies_total": len(comentarios),
            "comentarios":   comentarios,
            "url":           url,
        })
        print(f"\n  ✓ Publicación {i} guardada con {len(comentarios)} comentario(s).")

    return publicaciones

# ──────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────
def main():
    print()
    sep("═")
    print("  INGRESO DE DATOS REALES — Social Monitor")
    sep("═")
    print("  Copia y pega directamente desde tu Facebook / Instagram.")
    print("  Formato de fechas:  DD/MM/AAAA HH:MM   (ej: 22/04/2026 14:30)\n")

    # Cargar config existente
    config_path = os.path.join(BASE, "config.json")
    config = {}
    if os.path.exists(config_path):
        with open(config_path, encoding="utf-8") as f:
            config = json.load(f)

    # ── Nombres de cuenta
    fb_nombre = config.get("facebook",  {}).get("page_id",  "MiPagina")
    ig_nombre = config.get("instagram", {}).get("username", "mi_instagram")

    fb_nombre = leer(f"Nombre de tu Página de Facebook [{fb_nombre}]: ", fb_nombre)
    ig_nombre = leer(f"Usuario de Instagram (sin @)   [{ig_nombre}]: ", ig_nombre)

    # ── Configuración de SLA
    sep()
    print("  CONFIGURACIÓN DE SLA")
    print("  Define el tiempo máximo de respuesta por plataforma.\n")
    sla_fb = leer_float(f"  SLA Facebook  en horas [{config.get('facebook',{}).get('sla_horas',8)}]: ",
                        config.get("facebook", {}).get("sla_horas", 8))
    sla_ig = leer_float(f"  SLA Instagram en horas [{config.get('instagram',{}).get('sla_horas',12)}]: ",
                        config.get("instagram", {}).get("sla_horas", 12))
    sla_tw = leer_float(f"  SLA Twitter   en horas [{config.get('twitter',{}).get('sla_horas',4)}]: ",
                        config.get("twitter", {}).get("sla_horas", 4))

    # Actualizar y guardar config
    config.setdefault("facebook",  {})["page_id"]  = fb_nombre
    config.setdefault("instagram", {})["username"]  = ig_nombre
    config["facebook"]["sla_horas"]  = sla_fb
    config["instagram"]["sla_horas"] = sla_ig
    config.setdefault("twitter", {})["sla_horas"]   = sla_tw

    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    print(f"\n  ✓ SLA guardado:  Facebook={sla_fb}h  |  Instagram={sla_ig}h  |  Twitter={sla_tw}h\n")

    # ── Ingreso de publicaciones
    resultados = {"twitter": [], "facebook": [], "instagram": []}
    resultados["facebook"]  = ingresar_plataforma("facebook",  fb_nombre, sla_fb)
    resultados["instagram"] = ingresar_plataforma("instagram", ig_nombre, sla_ig)

    total = sum(len(v) for v in resultados.values())
    if total == 0:
        print("\n⚠️  No ingresaste ninguna publicación. Saliendo.")
        return

    # Guardar JSON
    with open(JSON_OUT, "w", encoding="utf-8") as f:
        json.dump(resultados, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n✅ Datos guardados ({total} publicaciones)")

    # Generar reporte
    print("🚀 Generando reporte HTML y Excel...")
    monitor_py = os.path.join(BASE, "monitor.py")
    result = subprocess.run([sys.executable, monitor_py, "--desde-json"], capture_output=False)

    if result.returncode == 0:
        html_path = os.path.join(BASE, "reporte_social_icetex.html")
        print(f"\n🎉 Listo. Abre el reporte con doble clic:")
        print(f"   {html_path}")
    else:
        print("\n⚠️  Error generando el reporte. Revisa los mensajes anteriores.")

if __name__ == "__main__":
    main()
