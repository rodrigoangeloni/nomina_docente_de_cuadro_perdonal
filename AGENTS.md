# AGENTS.md

Windows desktop tool that builds an institutional Excel payroll from "cuadro de personal" PDFs. `generar_nomina.py` contains the whole application; `README.md` is the Spanish end-user documentation.

## Commands

- Install dependencies: `python -m pip install -r requirements.txt`
- Generate from explicit PDFs:
  `python generar_nomina.py --pdfs "ESCUELA.pdf" "COLEGIO.pdf" --out "NOMINA.xlsx"`
- Auto-discover PDFs and write next to the script:
  `python generar_nomina.py --cli`
- Include non-teaching staff (`G*`): add `--todos` to either form.

**Do not run `python generar_nomina.py` bare for non-interactive verification.** With no arguments it opens the Tkinter main window and blocks in `mainloop()`; use `--cli`/`--pdfs` or instantiate `NominaApp` and drive it programmatically. `generar.bat` also ends by launching the GUI (and may try to install Python).

There is no test suite, linter, or CI. A focused check that needs no PDFs:

```bash
python -c "from generar_nomina import separar_apellidos_nombre as s; print(s('MARIA ELENA GARCIA LOPEZ'))"
```

Full verification requires real cuadro PDFs; they are gitignored and, if present locally, live in `CUADRO_PERSONAL_PDF/`.

## Architecture and gotchas

- Flow: `extract_filas` -> `procesar` -> `crear_excel`. Dedupe key is the cédula normalized with `str(int(ced))`, so leading zeros are dropped.
- The GUI is `NominaApp` in the same file: a 3-page wizard (`_pagina_archivos`, `_pagina_destino`, `_pagina_generar`) with `_abrir_guia` for help. `procesar(..., progreso=callback)` reports progress; the GUI runs it in a worker thread and polls a queue. The header updates from `extraer_instituciones(self.pdfs)`; do not hardcode a school name. Keep user-facing strings Spanish (with accents) and the three-step flow intact. Tkinter is stdlib; no extra dependency.
- PDF parsing is positional and silently skips rows that do not fit: requires at least 12 cells; name = index 1, cédula = index 2, rubro = index 3, turno = index 10; cédula must match `^[0-9]{4,8}$`. A changed PDF layout requires updating those indexes.
- Institution metadata is also positional: `extraer_institucion` reads fields 3-6 (name, code, phone, address) from the first page by word coordinates, and `extraer_instituciones` dedupes them. Both the GUI header and the Excel header use this; if the official form layout changes, adjust its x/y ranges. If detection fails, the Excel falls back to "NÓMINA DE DOCENTES".
- Rubro prefixes drive classification: `L*` = docente, `Z*` = catedrático, `G*` = non-teaching staff. `G*` (and unknown categories) are excluded unless `--todos` or `INCLUIR_FUNCIONARIOS = True`.
- `crear_excel` writes the extracted institution names/addresses into the merged header (rows 2-3); column widths and styling are hardcoded. Sheet name and "NOMINA DE DOCENTES" month/year use `datetime.now()`.
- `separar_apellidos_nombre` assumes the last two tokens are surnames, extending left across `DE`/`DEL`. Ambiguous real names go in untracked `manual_nombres.py` as `MANUAL = {"<normalized cédula>": "APELLIDOS, NOMBRES"}`; keys must have no leading zeros.
- `VERSION` in `generar_nomina.py` is the release source of truth and follows SemVer (`MAJOR.MINOR.PATCH`); `python bump_version.py X.Y.Z` rewrites it while preserving line endings.
- `generar_version_info.py` reads `VERSION`, `AUTOR`, and `NOMBRE_APP` from `generar_nomina.py` and writes the generated `version_info.txt` (four-part Windows `FileVersion`/`ProductVersion`, description, product, author, copyright, language). Both build scripts run it before PyInstaller; never hand-edit the EXE metadata.
- Self-update is intentionally inert when run from source: it only activates in the frozen EXE (`sys.frozen`), with a 24-hour cache at `%LOCALAPPDATA%\NominaDocentes\actualizacion.json`.
- Build: `build_exe.bat` requires `.venv\Scripts\python.exe` (create it with `python -m venv .venv` and install dependencies) and writes `dist\NominaDocentes.exe`. PyInstaller is not in `requirements.txt`; both build scripts install it on demand.
- The PyInstaller invocation is duplicated in `build_exe.bat` and `publicar.bat`; keep the flags and the `NOMBRE_EXE`/`NominaDocentes` name in sync, especially `--version-file "version_info.txt"`, `--collect-all pypdfium2`, and `--add-data "favicon.ico;."`.
- `publicar.bat X.Y.Z` bumps the version, regenerates metadata, commits the release files, refuses to overwrite an existing tag, pushes only the new tag, and creates a GitHub Release with `gh`. Do not run it unless the user explicitly asks to publish.
- Naming: the distributable is `NominaDocentes.exe` (PascalCase, ASCII, no spaces), matching the product name. Do not reintroduce `NOMINA DOCENTES.exe` or accents in the filename. Update the build scripts, README, and version metadata together if the name changes.
- Keep `CHANGELOG.md` current for user-visible changes (Keep a Changelog format).

## Data and conventions

- `*.pdf`, `*.xlsx`, `CUADRO_PERSONAL_PDF/`, and `manual_nombres.py` are gitignored because they contain real names and cédulas. Never commit them or use real data as test fixtures.
- Keep user-facing strings, error messages, and README updates in Spanish; identifiers and comments also follow Spanish naming.
- The Python script is cross-platform; the `.bat` workflow and self-update path target Windows.
