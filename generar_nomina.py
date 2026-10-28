# -*- coding: utf-8 -*-
"""
Genera la nomina de docentes a partir de los cuadros de personal (PDF).

- El usuario selecciona uno o VARIOS PDFs del cuadro de personal y el lugar
  donde guardar el Excel (ventanas graficas; no requiere escribir rutas).
- Extrae TODOS los docentes y catedraticos, SIN DUPLICADOS (una fila por
  persona, identificada por cedula, aunque aparezca en varios PDFs/turnos).
- Detalla TURNO y CATEGORIA de sueldo (rubro, ej. ZZ2).
- Clasificacion por categoria de rubro:
    L*  = docente   ZZ* = docente catedratico   G* = funcionario no docente
    (los G* se excluyen; para incluirlos usar --todos o INCLUIR_FUNCIONARIOS)
- El nombre y apellido se formatea como en la nomina:
    APELLIDO(S), NOMBRE(S)     ej.  PEREZ GOMEZ, JUAN CARLOS
  La separacion es automatica; solo los casos realmente ambiguos (nombres
  compuestos o apellidos compuestos) van en la tabla MANUAL (clave = cedula).

Modos de uso:
  - Sin argumentos: modo grafico (elegir PDFs y donde guardar el Excel).
  - "generar_nomina.py --cli": usa los PDFs de CUADRO_PERSONAL_PDF/ y guarda
    el resultado junto al script.
  - "generar_nomina.py --pdfs a.pdf b.pdf --out salida.xlsx": modo consola
    con archivos y destino indicados.
"""
import os
import re
import sys
import time
from datetime import datetime

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

BASE = os.path.dirname(os.path.abspath(__file__))

# ----------------------------------------------------------------------
# 1. Configuracion
# ----------------------------------------------------------------------
VERSION = "1.0.0"
REPO = "rodrigoangeloni/nomina_docente_de_cuadro_perdonal"   # para chequear actualizaciones
INCLUIR_FUNCIONARIOS = False   # True -> incluye tambien funcionarios (G*)
ORDEN_TURNO = ["M", "T", "N", "TN"]
CED_RE = re.compile(r"^[0-9]{4,8}$")

# Ajustes manuales de nombres ambiguos (opcional). Si existe el archivo local
# manual_nombres.py (excluido del repo por contener datos personales) se usa;
# si no, todo se resuelve con la separacion automatica.
try:
    from manual_nombres import MANUAL
except ImportError:
    MANUAL = {}

CONNECTORES = {"DE", "DEL"}

# ----------------------------------------------------------------------
# 2. Formato de nombres
# ----------------------------------------------------------------------
def separar_apellidos_nombre(bruto):
    """'MARIA ELENA GARCIA LOPEZ' -> 'GARCIA LOPEZ, MARIA ELENA' (heuristica).
    Los apellidos paterno y materno van al final; los conectores DE / DEL
    pertenecen al bloque de apellidos."""
    t = bruto.upper().split()
    n = len(t)
    if n <= 1:
        return t[0] if t else ""
    # bloque de apellidos: por defecto los ultimos 2 token (1 si solo hay 2)
    start = n - 2 if n >= 3 else n - 1
    cambio = True
    while cambio and start > 0:
        cambio = False
        if t[start] in CONNECTORES and start - 1 > 0:
            # el bloque empieza con conector -> el apellido previo tambien es apellido
            start -= 1
            cambio = True
    return " ".join(t[start:]) + ", " + " ".join(t[:start])

def formatear_nombre(bruto, cedula):
    if cedula in MANUAL:
        return MANUAL[cedula]
    return separar_apellidos_nombre(bruto)

# ----------------------------------------------------------------------
# 3. Extraccion de filas del cuadro
# ----------------------------------------------------------------------
def extract_filas(pdf_path):
    """Devuelve (nombre_bruto, cedula, categoria_rubro, turno) por cada fila
    de persona de las tablas del PDF."""
    filas = []
    with pdfplumber.open(pdf_path) as pdf:
        for p in pdf.pages:
            for t in p.extract_tables():
                for r in t:
                    cells = [str(c).replace("\n", " ").strip() if c is not None else "" for c in r]
                    if len(cells) < 12:
                        continue
                    nombre, ced, cat = cells[1], cells[2], cells[3]
                    turno = cells[10]
                    if not nombre or not ced:
                        continue
                    if nombre.upper().startswith("NOMBRES") or re.match(r"^[0-9.\s-]+$", nombre):
                        continue
                    if not CED_RE.match(ced):
                        continue
                    filas.append((nombre, ced, cat, turno))
    return filas

# ----------------------------------------------------------------------
# 4. Procesamiento principal
# ----------------------------------------------------------------------
def tipo(cats):
    for c in cats:
        if c.startswith("Z"):   # ZZ* = catedratico
            return "Catedratico"
        if c.startswith("L"):   # L* = docente
            return "Docente"
        if c.startswith("G"):   # G* = funcionario no docente
            return "Funcionario"
    return "Otro"

def procesar(pdf_paths, out_path, incluir_funcionarios=INCLUIR_FUNCIONARIOS):
    """Extrae, consolida y genera el Excel. Devuelve un resumen (dict)."""
    # dedupe por cedula normalizada (sin ceros a la izquierda)
    personas = {}
    for path in pdf_paths:
        for nombre, ced, cat, turno in extract_filas(path):
            clave = str(int(ced))
            p = personas.setdefault(clave, {"nombre": nombre, "turnos": [], "cats": []})
            if cat and cat.upper() not in p["cats"]:
                p["cats"].append(cat.upper())
            if turno and turno.upper() not in p["turnos"]:
                p["turnos"].append(turno.upper())

    # si ademas de 'Docente Director' hay categoria de rubro, quedan solo rubros
    for p in personas.values():
        rubros = [c for c in p["cats"] if re.match(r"^[LGZ]", c)]
        if rubros:
            p["cats"] = rubros
        p["turnos"].sort(key=lambda t: ORDEN_TURNO.index(t) if t in ORDEN_TURNO else len(ORDEN_TURNO))

    lista = []
    excluidos = 0
    for ced, p in personas.items():
        t = tipo(p["cats"])
        if t in ("Funcionario", "Otro") and not incluir_funcionarios:
            excluidos += 1
            continue
        lista.append({
            "nombre": formatear_nombre(p["nombre"], ced),
            "ced": ced,
            "turno": " / ".join(p["turnos"]),
            "categoria": " / ".join(p["cats"]),
        })
    lista.sort(key=lambda x: x["nombre"].upper())

    tot = {"Docente": 0, "Catedratico": 0, "Funcionario": 0, "Otro": 0}
    for x in lista:
        tot[tipo(x["categoria"].split(" / "))] += 1

    crear_excel(lista, out_path)

    return {
        "total": len(lista),
        "docentes": tot["Docente"],
        "catedraticos": tot["Catedratico"],
        "funcionarios": tot["Funcionario"],
        "excluidos": excluidos,
        "otros": [x for x in lista if tipo(x["categoria"].split(" / ")) == "Otro"],
        "destino": out_path,
    }

# ----------------------------------------------------------------------
# 5. Excel con el estilo de la nomina oficial
# ----------------------------------------------------------------------
def crear_excel(lista, out_path):
    hoy = datetime.now()
    wb = Workbook()
    ws = wb.active
    ws.title = "NOMINA DOCENTE " + str(hoy.year)

    ws.column_dimensions["A"].width = 6.14
    ws.column_dimensions["B"].width = 56.86
    ws.column_dimensions["C"].width = 22.57
    ws.column_dimensions["D"].width = 9.0
    ws.column_dimensions["E"].width = 13.5

    ws.row_dimensions[1].height = 36.0
    ws.row_dimensions[2].height = 40.5
    ws.row_dimensions[3].height = 15.0
    ws.row_dimensions[4].height = 5.25
    ws.row_dimensions[5].height = 20.25
    ws.row_dimensions[6].height = 7.5
    ws.row_dimensions[7].height = 39.75

    ws.merge_cells("A2:C2")
    ws["A2"] = 'ESCUELA BÁSICA N° 806 Y COLEGIO NACIONAL\n "INSPECTOR MAYOR EUSEBIO FLORENTÍN"'
    ws["A2"].font = Font(name="Times New Roman", size=12, bold=True)
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.merge_cells("A3:C3")
    ws["A3"] = "Hermes Irrazabal y /Carmelo Figueredo"
    ws["A3"].font = Font(name="Times New Roman", size=9, bold=True)
    ws["A3"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A5:C5")
    ws["A5"] = "NOMINA DE DOCENTES %02d/%d" % (hoy.month, hoy.year)
    ws["A5"].font = Font(name="Times New Roman", size=12, bold=True)
    ws["A5"].alignment = Alignment(horizontal="center", vertical="center")

    headers = ["N°", "APELLIDO Y NOMBRE", "CEDULA DE\nIDENTIDAD", "TURNO", "CATEGORIA\nDE SUELDO"]
    header_font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="2E75B6")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="000000")
    head_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col, val in zip(("A", "B", "C", "D", "E"), headers):
        c = ws[col + "7"]
        c.value = val
        c.font = header_font
        c.fill = header_fill
        c.alignment = header_align
        c.border = head_border

    def ced_val(ced):
        return int(ced) if ced.isdigit() else ced

    start = 8
    data_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for i, x in enumerate(lista):
        r = start + i
        ws.cell(r, 1, value=i + 1)
        ws.cell(r, 2, value=x["nombre"])
        ws.cell(r, 3, value=ced_val(x["ced"]))
        ws.cell(r, 4, value=x["turno"])
        ws.cell(r, 5, value=x["categoria"])
        ws.row_dimensions[r].height = 20.1
        for col, horiz in ((1, "center"), (2, "left"), (3, "center"), (4, "center"), (5, "center")):
            c = ws.cell(r, col)
            c.border = data_border
            c.alignment = Alignment(horizontal=horiz, vertical="center")
            c.font = Font(name="Calibri", size=11 if col == 1 else 12)

    wb.save(out_path)

# ----------------------------------------------------------------------
# 6. Chequeo de actualizaciones (solo cuando corre como EXE)
# ----------------------------------------------------------------------
# GitHub limita su API (60 consultas/hora por IP sin autenticar), por eso
# el chequeo se hace A LO SUMO una vez cada INTERVALO_CHEQUEO_HORAS y se
# guarda en cache local (tambien la version que el usuario decidio omitir).
INTERVALO_CHEQUEO_HORAS = 24

def _version_tuple(v):
    nums = re.findall(r"\d+", v)[:3]
    return tuple(int(x) for x in nums) if nums else (0,)

def _ruta_cache():
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    carpeta = os.path.join(base, "NominaDocentes")
    os.makedirs(carpeta, exist_ok=True)
    return os.path.join(carpeta, "actualizacion.json")

def _leer_cache():
    try:
        import json
        with open(_ruta_cache(), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def _guardar_cache(data):
    try:
        import json
        with open(_ruta_cache(), "w", encoding="utf-8") as f:
            json.dump(data, f)
    except Exception:
        pass   # la cache nunca debe romper el programa

def _consultar_github():
    """(tag, url_exe) del ultimo release publicado."""
    import json
    import urllib.request
    req = urllib.request.Request(
        "https://api.github.com/repos/%s/releases/latest" % REPO,
        headers={"User-Agent": "nomina-docentes"},
    )
    with urllib.request.urlopen(req, timeout=5) as r:
        data = json.load(r)
    tag = (data.get("tag_name") or "").lstrip("vV")
    asset = next((a for a in data.get("assets", []) if a.get("name", "").lower().endswith(".exe")), None)
    return tag, (asset["browser_download_url"] if asset else None)

def hay_nueva_version():
    """Devuelve (tag, url_exe) si hay version mas nueva que VERSION pendiente.
    Consulta la red solo si paso el intervalo; si no, usa la cache."""
    if not getattr(sys, "frozen", False):   # solo tiene sentido en el EXE
        return None

    cache = _leer_cache()
    ahora = time.time()
    vence = cache.get("last_check", 0) + INTERVALO_CHEQUEO_HORAS * 3600

    if ahora < vence:
        # dentro del intervalo: usar lo ya consultado, sin tocar la red
        tag, url = cache.get("last_tag"), cache.get("last_url")
        if tag and url and _version_tuple(tag) > _version_tuple(VERSION) and tag != cache.get("declined"):
            return tag, url
        return None

    # paso el intervalo: consultar la red (una sola vez cada X horas)
    try:
        tag, url = _consultar_github()
    except Exception:
        cache["last_check"] = ahora   # sin red: reintentar recien pasado el intervalo
        _guardar_cache(cache)
        return None

    cache["last_check"] = ahora
    cache["last_tag"], cache["last_url"] = tag, url
    _guardar_cache(cache)

    if tag and url and _version_tuple(tag) > _version_tuple(VERSION) and tag != cache.get("declined"):
        return tag, url
    return None

def descargar_y_reemplazar(url):
    """Descarga el EXE nuevo y reemplaza el actual (se puede renombrar un exe
    en ejecucion, no sobrescribir). Devuelve la ruta del .old."""
    import shutil
    import urllib.request
    exe = sys.executable
    nuevo = exe + ".nueva"
    req = urllib.request.Request(url, headers={"User-Agent": "nomina-docentes"})
    with urllib.request.urlopen(req, timeout=120) as r, open(nuevo, "wb") as f:
        shutil.copyfileobj(r, f)
    viejo = exe + ".old"
    if os.path.exists(viejo):
        os.remove(viejo)
    os.rename(exe, viejo)
    shutil.move(nuevo, exe)
    # limpiar la cache: el exe nuevo arranca con datos frescos
    try:
        os.remove(_ruta_cache())
    except Exception:
        pass
    return viejo

def verificar_actualizacion(root):
    """Pregunta al usuario si quiere actualizar. Si acepta, descarga, reemplaza
    y reinicia. Devuelve True si se actualizo (y hay que salir)."""
    from tkinter import messagebox
    try:
        upd = hay_nueva_version()
    except Exception:
        return False          # sin internet o sin release: seguir normal
    if not upd:
        return False
    tag, url = upd
    si = messagebox.askyesno(
        "Nueva version disponible",
        "Hay una version nueva: v%s\nTu version actual: v%s\n\n"
        "Quieres descargarla ahora?\n\n(Si eliges No, no se volvera a preguntar por esta version)" % (tag, VERSION),
        parent=root,
    )
    if not si:
        cache = _leer_cache()
        cache["declined"] = tag   # no volver a preguntar por esta version
        _guardar_cache(cache)
        return False
    try:
        descargar_y_reemplazar(url)
    except Exception as e:
        messagebox.showerror(
            "Actualizacion",
            "No se pudo completar la actualizacion:\n\n%s\n\nSe continuara con la version actual." % e,
            parent=root,
        )
        return False
    messagebox.showinfo("Actualizacion", "Actualizado a v%s. El programa se reiniciara." % tag, parent=root)
    import subprocess
    subprocess.Popen([sys.executable])
    return True

# ----------------------------------------------------------------------
# 7. Modo grafico (para usuarios no tecnicos)
# ----------------------------------------------------------------------
def main_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    if verificar_actualizacion(root):
        root.destroy()
        return

    pdfs = filedialog.askopenfilenames(
        parent=root,
        title="Selecciona el/los cuadro(s) de personal (PDF) - puedes elegir varios",
        filetypes=[("Archivos PDF", "*.pdf")],
    )
    if not pdfs:
        messagebox.showwarning("Nomina de docentes", "No se selecciono ningun PDF. Operacion cancelada.")
        return

    hoy = datetime.now()
    out = filedialog.asksaveasfilename(
        parent=root,
        title="Guardar nomina como...",
        defaultextension=".xlsx",
        initialfile="NOMINA DOCENTES %02d-%d.xlsx" % (hoy.month, hoy.year),
        filetypes=[("Libro de Excel", "*.xlsx")],
    )
    if not out:
        return

    try:
        resumen = procesar(list(pdfs), out)
    except Exception as e:
        messagebox.showerror(
            "Nomina de docentes",
            "Ocurrio un error al generar la nomina:\n\n%s" % e,
        )
        raise

    messagebox.showinfo(
        "Nomina de docentes",
        "Nomina generada con exito\n\n"
        "  Docentes: %d\n"
        "  Catedraticos: %d\n"
        "  Total: %d\n\n"
        "Guardado en:\n%s" % (
            resumen["docentes"], resumen["catedraticos"],
            resumen["total"], resumen["destino"],
        ),
    )

# ----------------------------------------------------------------------
# 8. Modo consola
# ----------------------------------------------------------------------
def find_pdf(keyword):
    candidates = []
    for root, dirs, files in os.walk(BASE):
        depth = root[len(BASE):].count(os.sep)
        if depth > 2:
            continue
        for f in files:
            if f.lower().endswith(".pdf") and keyword.lower() in f.lower():
                candidates.append(os.path.join(root, f))
    candidates.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return candidates[0] if candidates else None

def main_cli(argv):
    pdfs, out, todos = [], None, False
    if "--pdfs" in argv:
        i = argv.index("--pdfs")
        while i + 1 < len(argv) and not argv[i + 1].startswith("--"):
            pdfs.append(argv[i + 1])
            i += 1
    if "--out" in argv:
        out = argv[argv.index("--out") + 1]
    todos = "--todos" in argv

    if not pdfs:  # fallback: PDFs de la carpeta estandar
        for kw in ("ESCUELA", "COLEGIO"):
            p = find_pdf(kw)
            if p:
                pdfs.append(p)
    if not pdfs:
        raise FileNotFoundError("No se encontraron PDFs del cuadro de personal")
    if not out:
        out = os.path.join(BASE, "NOMINA DOCENTES ESCUELA Y COLEGIO 2026 - GENERADO.xlsx")

    print("PDFs:")
    for p in pdfs:
        print("  -", p)

    resumen = procesar(pdfs, out, incluir_funcionarios=todos)

    print()
    print("Generado:", resumen["destino"])
    print("Version del script:", VERSION)
    print("Total docentes:", resumen["total"])
    print("  Docentes (L*):", resumen["docentes"])
    print("  Catedraticos (ZZ*):", resumen["catedraticos"])
    print("  Excluidos funcionarios (G*):", resumen["excluidos"])
    for x in resumen["otros"]:
        print("  !! Revisar categoria:", x["nombre"], x["ced"], x["categoria"])

if __name__ == "__main__":
    argv = sys.argv[1:]
    if "--cli" in argv or "--pdfs" in argv or "--out" in argv or "--todos" in argv:
        main_cli(argv)
    else:
        main_gui()
