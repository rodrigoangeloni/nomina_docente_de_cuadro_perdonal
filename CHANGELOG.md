# Changelog

Todos los cambios importantes de este proyecto se documentan en este archivo.

El formato sigue [Keep a Changelog](https://keepachangelog.com/es-ES/1.1.0/) y el
versionado adopta [Semantic Versioning](https://semver.org/lang/es/)
(`MAJOR.MINOR.PATCH`).

## [1.1.0] - 2026-10-28

### Added

- Interfaz gráfica guiada de tres pasos: selección de PDFs, destino/opciones y generación.
- Lista de archivos con opciones para quitar seleccionados o limpiar la lista.
- Barra de progreso, resumen final con totales y accesos para abrir el Excel o la carpeta.
- Guía de ayuda integrada (botón «? Guía de uso» y tecla `F1`) con requisitos y preguntas frecuentes.
- Procesamiento en segundo plano para no congelar la ventana.
- Cabecera institucional del Excel y de la GUI tomada del encabezado del PDF: nombre y dirección; si hay varios PDFs, se muestran las instituciones detectadas.
- Metadatos de Windows en el EXE: descripción, producto, autor, copyright, versión e idioma.

### Changed

- El ejecutable pasa de `NOMINA DOCENTES.exe` a `NominaDocentes.exe` (PascalCase, sin espacios ni acentos).
- La versión se mantiene en una sola fuente (`VERSION`) y se publica como `FileVersion`/`ProductVersion` de cuatro componentes.
- `bump_version.py` conserva los saltos de línea originales del código.
- `publicar.bat` ya no sobrescribe tags existentes y sube únicamente el tag de la versión publicada.

### Fixed

- La interfaz ya no depende de diálogos sueltos: el flujo completo se inicia desde la ventana principal.
- Los errores de escritura (por ejemplo, el Excel abierto) se muestran con un mensaje claro.

## [1.0.0] - 2026-10-28

### Added

- Versión inicial: extracción de docentes y catedráticos desde los Cuadros de Personal en PDF.
- Unificación por cédula, consolidación de turnos y rubros, y exportación al Excel institucional.
- Modo consola y autoactualización del EXE mediante GitHub Releases.
