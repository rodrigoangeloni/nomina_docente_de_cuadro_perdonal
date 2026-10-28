# -*- coding: utf-8 -*-
"""Actualiza la constante VERSION de generar_nomina.py.
Uso: python bump_version.py 1.0.1"""
import io
import re
import sys

if len(sys.argv) < 2:
    sys.exit("Uso: python bump_version.py VERSION (ej: 1.0.1)")

version = sys.argv[1].lstrip("vV")
if not re.fullmatch(r"[0-9]+(\.[0-9]+){0,2}", version):
    sys.exit("Version invalida: %r (formato esperado: 1.0.0)" % version)

with io.open("generar_nomina.py", encoding="utf-8") as f:
    contenido = f.read()

nuevo, n = re.subn(r'VERSION = "[0-9.]+"', 'VERSION = "%s"' % version, contenido, count=1)
if n == 0:
    sys.exit("No se encontro la linea VERSION en generar_nomina.py")

with io.open("generar_nomina.py", "w", encoding="utf-8", newline="\n") as f:
    f.write(nuevo)

print("VERSION actualizada a", version)
