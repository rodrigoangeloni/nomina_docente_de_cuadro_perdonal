# -*- coding: utf-8 -*-
"""Genera version_info.txt para los metadatos de Windows del EXE.

La version y el autor se leen de generar_nomina.py para mantener una sola
fuente de verdad. El archivo generado se pasa a PyInstaller con
--version-file. Se ejecuta automaticamente desde build_exe.bat y publicar.bat.
"""
import io
import os
import re
import sys
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_APP = os.path.join(BASE, "generar_nomina.py")
ARCHIVO_SALIDA = os.path.join(BASE, "version_info.txt")
NOMBRE_EXE = "NominaDocentes"
LANG_ES_PARAGUAY = 15370   # 0x3C0A
CODEPAGE_UNICODE = 1200   # 0x04B0

PLANTILLA = """# -*- coding: utf-8 -*-
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers={filevers},
    prodvers={prodvers},
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        '3C0A04B0',
        [
          StringStruct('CompanyName', {company}),
          StringStruct('FileDescription', {description}),
          StringStruct('FileVersion', {file_version}),
          StringStruct('InternalName', {internal_name}),
          StringStruct('LegalCopyright', {copyright}),
          StringStruct('OriginalFilename', {original_filename}),
          StringStruct('ProductName', {product_name}),
          StringStruct('ProductVersion', {product_version}),
          StringStruct('Comments', {comments})
        ]
      )
    ]),
    VarFileInfo([VarStruct('Translation', [{lang}, {codepage}])])
  ]
)
"""


def _leer_constante(contenido, nombre):
    patron = re.compile(r'^%s\s*=\s*"([^"]+)"' % re.escape(nombre), re.MULTILINE)
    coincidencia = patron.search(contenido)
    if not coincidencia:
        sys.exit("No se encontro la constante %s en generar_nomina.py" % nombre)
    return coincidencia.group(1)


def main():
    with io.open(ARCHIVO_APP, encoding="utf-8", newline="") as archivo:
        contenido = archivo.read()

    version = _leer_constante(contenido, "VERSION")
    autor = _leer_constante(contenido, "AUTOR")
    producto = _leer_constante(contenido, "NOMBRE_APP")

    if not re.fullmatch(r"[0-9]+(\.[0-9]+){0,2}", version):
        sys.exit("Version invalida: %r (formato esperado: 1.0.0)" % version)

    partes = [int(parte) for parte in version.split(".")]
    partes += [0] * (4 - len(partes))
    partes = partes[:4]
    version4 = ".".join(str(parte) for parte in partes)
    anio = datetime.now().year

    valores = {
        "filevers": repr(tuple(partes)),
        "prodvers": repr(tuple(partes)),
        "company": repr(autor),
        "description": repr("Generador de nómina docente"),
        "file_version": repr(version4),
        "internal_name": repr(NOMBRE_EXE),
        "copyright": repr("© %d %s" % (anio, autor)),
        "original_filename": repr(NOMBRE_EXE + ".exe"),
        "product_name": repr(producto),
        "product_version": repr(version4),
        "comments": repr(
            "Aplicación de escritorio para generar la nómina de docentes y "
            "catedráticos en Excel a partir de los Cuadros de Personal en PDF."),
        "lang": LANG_ES_PARAGUAY,
        "codepage": CODEPAGE_UNICODE,
    }

    with io.open(ARCHIVO_SALIDA, "w", encoding="utf-8", newline="\n") as archivo:
        archivo.write(PLANTILLA.format(**valores))

    print("Metadatos generados: %s (v%s, autor: %s)" % (
        os.path.basename(ARCHIVO_SALIDA), version4, autor))


if __name__ == "__main__":
    main()
