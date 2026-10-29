# -*- coding: utf-8 -*-
"""Lee, valida o actualiza la constante VERSION de generar_nomina.py (SemVer).

Uso:
  python bump_version.py 1.1.0        actualiza VERSION
  python bump_version.py --show       muestra la version actual
  python bump_version.py --next patch muestra la proxima version (patch/minor/major)
  python bump_version.py --validate 1.1.1  valida que sea mayor que la actual
"""
import io
import re
import sys

ARCHIVO = "generar_nomina.py"


def leer_contenido():
    # newline="" conserva los saltos de linea originales (CRLF o LF)
    with io.open(ARCHIVO, encoding="utf-8", newline="") as archivo:
        return archivo.read()


def version_actual(contenido):
    coincidencia = re.search(r'VERSION = "([0-9.]+)"', contenido)
    if not coincidencia:
        sys.exit("No se encontro la linea VERSION en generar_nomina.py")
    return coincidencia.group(1)


def normalizar(version):
    version = version.lstrip("vV")
    if not re.fullmatch(r"[0-9]+(\.[0-9]+){0,2}", version):
        sys.exit("Version invalida: %r (formato esperado: 1.0.0)" % version)
    return version


def clave(version):
    partes = [int(parte) for parte in version.split(".")]
    partes += [0] * (3 - len(partes))
    return tuple(partes[:3])


def siguiente(version, tipo):
    mayor, menor, parche = clave(version)
    if tipo == "major":
        return "%d.0.0" % (mayor + 1)
    if tipo == "minor":
        return "%d.%d.0" % (mayor, menor + 1)
    if tipo == "patch":
        return "%d.%d.%d" % (mayor, menor, parche + 1)
    sys.exit("Tipo invalido: %r (usar patch, minor o major)" % tipo)


def main():
    if len(sys.argv) < 2:
        sys.exit(
            "Uso: python bump_version.py VERSION | --show | "
            "--next patch|minor|major | --validate VERSION")

    argumento = sys.argv[1]
    contenido = leer_contenido()
    actual = version_actual(contenido)

    if argumento in ("--show", "-s"):
        print(actual)
        return

    if argumento == "--next":
        if len(sys.argv) < 3:
            sys.exit("Uso: python bump_version.py --next patch|minor|major")
        print(siguiente(actual, sys.argv[2]))
        return

    if argumento == "--validate":
        if len(sys.argv) < 3:
            sys.exit("Uso: python bump_version.py --validate VERSION")
        nueva = normalizar(sys.argv[2])
        if clave(nueva) <= clave(actual):
            sys.exit(
                "La version nueva (%s) debe ser mayor que la actual (%s)."
                % (nueva, actual))
        print("Version valida: %s (actual: %s)" % (nueva, actual))
        return

    version = normalizar(argumento)
    nuevo, cantidad = re.subn(
        r'VERSION = "[0-9.]+"', 'VERSION = "%s"' % version,
        contenido, count=1)
    if cantidad == 0:
        sys.exit("No se encontro la linea VERSION en generar_nomina.py")

    with io.open(ARCHIVO, "w", encoding="utf-8", newline="") as archivo:
        archivo.write(nuevo)

    print("VERSION actualizada a", version)


if __name__ == "__main__":
    main()
