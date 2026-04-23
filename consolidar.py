#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Consolida múltiples archivos JSON enviados por el equipo
y genera el reporte unificado.

Uso:
  1. Coloca todos los archivos JSON en la carpeta 'envios/'
  2. Ejecuta:  python consolidar.py
"""

import json
import os
import sys
import subprocess
import glob

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

BASE      = os.path.dirname(os.path.abspath(__file__))
ENVIOS    = os.path.join(BASE, "envios")
JSON_OUT  = os.path.join(BASE, "datos_social_icetex.json")

def sep(n=54): print("─" * n)

def main():
    print()
    sep()
    print("  CONSOLIDADOR — Social Monitor")
    sep()

    # Crear carpeta envios si no existe
    os.makedirs(ENVIOS, exist_ok=True)

    archivos = glob.glob(os.path.join(ENVIOS, "*.json"))
    if not archivos:
        print(f"\n⚠️  No hay archivos JSON en la carpeta 'envios/'")
        print(f"   Ruta esperada: {ENVIOS}")
        print(f"\n   Pasos:")
        print(f"   1. Pide a cada persona que abra 'formulario.html' y descargue su JSON")
        print(f"   2. Coloca todos los JSON en la carpeta 'envios/'")
        print(f"   3. Vuelve a ejecutar este script")
        return

    print(f"\n📂 Archivos encontrados en 'envios/': {len(archivos)}")

    consolidado = {"twitter": [], "facebook": [], "instagram": []}
    analistas   = []
    sla_visto   = {}

    for ruta in sorted(archivos):
        nombre = os.path.basename(ruta)
        try:
            with open(ruta, encoding="utf-8") as f:
                d = json.load(f)

            meta = d.get("_meta", {})
            analista = meta.get("analista", nombre)
            analistas.append(analista)

            # Tomar SLA del primer archivo que lo tenga
            if not sla_visto and meta.get("sla"):
                sla_visto = meta["sla"]

            fb  = d.get("facebook",  [])
            ig  = d.get("instagram", [])
            tw  = d.get("twitter",   [])

            consolidado["facebook"]  += fb
            consolidado["instagram"] += ig
            consolidado["twitter"]   += tw

            print(f"  ✓ {analista:30s}  →  FB:{len(fb):2d}  IG:{len(ig):2d}  TW:{len(tw):2d} publicaciones")

        except Exception as e:
            print(f"  ✗ Error en {nombre}: {e}")

    total = sum(len(v) for v in consolidado.values())
    if total == 0:
        print("\n⚠️  Ningún archivo tenía publicaciones válidas.")
        return

    total_coms = sum(
        len(c["comentarios"])
        for plat in consolidado.values()
        for c in plat
    )

    sep()
    print(f"  Total publicaciones: {total}")
    print(f"  Total comentarios:   {total_coms}")
    print(f"  Analistas:           {', '.join(analistas)}")
    sep()

    # Actualizar config con SLA si se encontró
    if sla_visto:
        config_path = os.path.join(BASE, "config.json")
        try:
            with open(config_path, encoding="utf-8") as f:
                config = json.load(f)
            config.setdefault("facebook",  {})["sla_horas"] = sla_visto.get("facebook",  8)
            config.setdefault("instagram", {})["sla_horas"] = sla_visto.get("instagram", 12)
            config.setdefault("twitter",   {})["sla_horas"] = sla_visto.get("twitter",   4)
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            print(f"  ✓ SLA actualizado en config.json")
        except Exception as e:
            print(f"  ⚠️  No se pudo actualizar config.json: {e}")

    # Guardar JSON consolidado
    with open(JSON_OUT, "w", encoding="utf-8") as f:
        json.dump(consolidado, f, ensure_ascii=False, indent=2, default=str)
    print(f"  ✓ JSON consolidado guardado")

    # Generar reporte
    print("\n🚀 Generando reporte unificado...")
    monitor_py = os.path.join(BASE, "monitor.py")
    result = subprocess.run([sys.executable, monitor_py, "--desde-json"], capture_output=False)

    if result.returncode == 0:
        html = os.path.join(BASE, "reporte_social_icetex.html")
        print(f"\n🎉 Reporte listo. Ábrelo con doble clic:")
        print(f"   {html}")
    else:
        print("\n⚠️  Error al generar el reporte.")

if __name__ == "__main__":
    main()
