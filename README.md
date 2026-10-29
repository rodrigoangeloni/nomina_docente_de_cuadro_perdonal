# 📊 Nómina de Docentes — Escuela Básica y Colegio Nacional

![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?logo=python&logoColor=white)
![Versión](https://img.shields.io/github/v/release/rodrigoangeloni/nomina_docente_de_cuadro_perdonal?label=versi%C3%B3n&color=blue)
![openpyxl](https://img.shields.io/badge/openpyxl-Excel-217346?logo=microsoftexcel&logoColor=white)
![Estado](https://img.shields.io/badge/Estado-Operativo-brightgreen)

Genera automáticamente la **nómina de docentes y catedráticos** de una Escuela Básica y un Colegio Nacional a partir de los **Cuadros de Personal en PDF**, con formato, turnos y categorías de sueldo listos para usar. ⚡

---

## 📑 Tabla de contenidos

- [✨ Qué hace](#-qué-hace)
- [📁 Estructura del proyecto](#-estructura-del-proyecto)
- [⚙️ Requisitos e instalación](#️-requisitos-e-instalación)
- [🚀 Uso](#-uso)
- [🔁 Actualizaciones automáticas](#-actualizaciones-automáticas)
- [🏷️ Publicar una nueva versión](#️-publicar-una-nueva-versión)
- [🧠 Cómo funciona](#-cómo-funciona)
- [📤 Salida generada](#-salida-generada)
- [✏️ Casos especiales (nombres)](#️-casos-especiales-nombres)
- [⚠️ Notas y advertencias](#️-notas-y-advertencias)
- [🗓️ Última revisión](#️-última-revisión)

---

## ✨ Qué hace

1. 📥 **Lee** los PDFs del cuadro de personal que **tú mismo seleccionas** en una ventana gráfica (puedes elegir uno o varios a la vez, de cualquier carpeta).
2. 🔍 **Extrae** a todos los docentes y catedráticos, **sin duplicados** (una sola fila por persona aunque figure en varios PDFs o en varios turnos).
3. 🏷️ **Clasifica** cada persona según su **rubro de sueldo** del cuadro:

   | Rubro | Significado | ¿Se incluye? |
   |-------|-------------|:------------:|
   | `L*` (LCE, LCV, LAF, LAC, LBG, LCH, LDU, LA7…) | 👩‍🏫 Docente | ✅ Sí |
   | `ZZ*` (ZZ2, ZZ5, ZZ8…) | 🎓 Docente catedrático | ✅ Sí |
   | `G*` (G3N, G3Z…) | 🖥️ Funcionario no docente | ❌ No* |

   \* *Se puede incluir con el flag `--todos` en consola o `INCLUIR_FUNCIONARIOS = True` en el script.*
4. 📝 **Formatea** el nombre como en la nómina oficial: `APELLIDO(S), NOMBRE(S)` → `PEREZ GOMEZ, JUAN CARLOS`.
5. 📤 **Genera** el Excel final con el estilo institucional (encabezado azul, bordes, Times New Roman/Calibri) y el **nombre y la dirección de la institución tomados del encabezado del PDF**.

---

## 📁 Estructura del proyecto

```text
NOMINA_DOCENTES/
├── 📄 generar_nomina.py              # Script principal (lógica + interfaz gráfica)
├── 🖱️ generar.bat                    # PC nueva: instala todo + abre el programa
├── 🔨 build_exe.bat                  # Compila un EXE para PC sin Python
├── 🚢 publicar.bat                   # Versiona (tag) y publica el EXE en GitHub Releases
├── 🔢 bump_version.py                # Actualiza la constante VERSION (SemVer)
├── 🏷️ generar_version_info.py        # Genera los metadatos de Windows del EXE
├── 📝 CHANGELOG.md                   # Historial de versiones
├── 📦 requirements.txt               # Dependencias (pdfplumber, openpyxl)
├── 🎨 favicon.ico                    # Icono de la aplicación y del EXE
├── 🔏 manual_nombres.py              # (local, no se sube) nombres ambiguos
├── 📁 CUADRO_PERSONAL_PDF/           # (opcional) PDFs del cuadro de personal
└── 📊 NOMINA ...xlsx                 # resultado (excluido del repo por privacidad)
```

---

## ⚙️ Requisitos e instalación

### 🅰️ PC nueva, sin nada instalado (recomendado)

No hay que hacer nada manual: **doble clic en `generar.bat`** y él solo:

1. 🔍 Detecta si falta Python y lo **instala automáticamente** (con `winget` o el instalador oficial, sin permisos de administrador).
2. 🧪 Crea el entorno virtual **`.venv`** aislado para este proyecto.
3. 📦 Instala las dependencias (`pdfplumber`, `openpyxl`) desde `requirements.txt`.
4. 🚀 Abre el programa.

> ⚠️ La primera vez pide conexión a internet y tarda unos minutos. Las siguientes veces abre directo.

### 🅱️ Instalación manual (opcional)

Si ya tenés Python y preferís hacerlo a mano:

```bash
pip install -r requirements.txt
```

---

## 🚀 Uso

### 👩‍🏫 Para usuarios no técnicos (sin Python) — EXE

1. 🔨 En una PC con todo configurado, doble clic en **`build_exe.bat`**.
2. 📦 Espera a que compile: queda en **`dist\NominaDocentes.exe`** (un único archivo).
3. 📋 Copia ese EXE a cualquier PC (pendrive, correo, red) — **no necesita Python ni instalación**.
4. 🖱️ Doble clic en el EXE y seguí el flujo de abajo.

### 🖼️ Flujo del programa (interfaz guiada)

1. 🪟 Se abre la ventana **Nómina de Docentes**, con el nombre de la aplicación, un flujo de tres pasos y una guía de ayuda integrada (botón **? Guía de uso** o tecla `F1`).
2. 📥 **Paso 1 — Archivos:** hacé clic en **Seleccionar PDFs…** y elegí uno o varios cuadros de personal (con `Ctrl` o `Shift` podés marcar varios; por ejemplo, el de la Escuela y el del Colegio). La lista muestra los archivos elegidos y permite quitarlos o limpiarla.
3. 💾 **Paso 2 — Destino:** la aplicación propone una carpeta y un nombre (`NOMINA DOCENTES MM-AAAA.xlsx`). Podés cambiarlos con **Cambiar ubicación y nombre…**. Ahí también podés marcar **Incluir funcionarios no docentes (G*)** y **Abrir el Excel automáticamente al finalizar**.
4. ✅ **Paso 3 — Generar:** revisá el resumen y hacé clic en **Generar nómina ahora**. La barra de progreso y el estado muestran qué está haciendo. Al terminar ves el total de docentes, catedráticos y funcionarios, y podés **Abrir Excel**, **Abrir carpeta** o **Generar de nuevo**.

> 💡 La guía integrada explica cada paso, los requisitos de los PDF y las preguntas frecuentes. También podés volver atrás en cualquier momento para cambiar archivos, destino u opciones.
>
> 🏫 Al seleccionar los PDFs, la cabecera de la ventana muestra las instituciones detectadas. El Excel usa ese mismo nombre y la dirección que figuran en el encabezado del Cuadro de Personal.

### 💻 Modo consola (opcional, para automatizar)

```bash
:: Con PDFs y destino propios
python generar_nomina.py --pdfs "ESCUELA_CUADRO_PERSONAL.pdf" "COLEGIO_CUADRO_PERSONAL.pdf" --out "NOMINA 2026.xlsx"

:: Incluyendo también funcionarios (G*)
python generar_nomina.py --pdfs *.pdf --todos --out "NOMINA 2026.xlsx"

:: Sin argumentos: busca los PDFs de CUADRO_PERSONAL_PDF\ y guarda junto al script
python generar_nomina.py --cli
```

---

## 🔁 Actualizaciones automáticas

El **EXE se autoverifica** cada vez que se ejecuta:

1. 🌐 Consulta la última versión publicada en GitHub Releases (`releases/latest`).
2. 💬 Si existe una versión más nueva, pregunta: **«¿Quieres descargarla ahora?»**.
   - ✅ **Sí** → descarga el EXE nuevo, reemplaza el actual y se reinicia solo.
   - ❌ **No** → continúa normalmente con la versión actual (y **no vuelve a preguntar por esa versión**).
3. 📴 Si no hay internet o falla la consulta, **abre directamente sin molestar**.

> ⏱️ **Límite de API respetado:** GitHub permite 60 consultas/hora por IP sin autenticar. Por eso el EXE consulta la red **como máximo una vez cada 24 horas** y guarda el resultado en una caché local (`%LOCALAPPDATA%\NominaDocentes\actualizacion.json`). Abrir el programa muchas veces seguidas **no** genera consultas repetidas; si hay una actualización pendiente que no descargaste, el aviso reaparece sin volver a consultar la red.

> ℹ️ El chequeo funciona porque el repositorio es **público** (la API de GitHub no pide token).

## 🏷️ Publicar una nueva versión

Para desarrolladores (en la PC de desarrollo, con el repo clonado):

```bash
:: 1. Actualiza VERSION, compila el EXE, crea el tag vX.Y.Z
::    y publica el Release con el binario (requiere GitHub CLI "gh")
publicar.bat 1.1.0
```

Lo que hace por dentro:

| Paso | Acción |
|------|--------|
| 1️⃣ | Actualiza `VERSION = "x.y.z"` en `generar_nomina.py` (con `bump_version.py`) |
| 2️⃣ | Genera los metadatos de Windows y compila `dist\NominaDocentes.exe` con PyInstaller |
| 3️⃣ | Crea el tag `vx.y.z` y lo sube al repo (sin sobrescribir tags existentes) |
| 4️⃣ | Crea el **Release** en GitHub adjuntando el EXE |

A partir del siguiente arranque, todos los EXE viejos detectan el release nuevo y ofrecen actualizarse. 🎉

> 🏷️ **Versionado:** se usa [SemVer](https://semver.org/lang/es/) (`MAJOR.MINOR.PATCH`). La versión se define una sola vez en `VERSION` dentro de `generar_nomina.py`; `generar_version_info.py` la convierte a `1.1.0.0` y completa los metadatos de Windows del EXE (descripción, producto, autor, copyright, idioma). No edites el EXE a mano.
>
> 📝 Cada versión se documenta en [`CHANGELOG.md`](CHANGELOG.md).
>
> 💡 El script auxiliar `bump_version.py` también se puede usar a mano (`python bump_version.py 1.1.1`) para ajustar la versión sin publicar todavía.

---

## 🧠 Cómo funciona

| Paso | Detalle |
|------|---------|
| 🔎 **Selección de PDFs** | En modo gráfico elegís uno o varios PDFs de cualquier carpeta; en modo consola (`--cli`) busca recursivamente (hasta 2 niveles) archivos `.pdf` con `ESCUELA` / `COLEGIO` en el nombre y usa el más reciente de cada uno. |
| 🏫 **Institución** | El nombre y la dirección se leen de la cabecera del Cuadro de Personal (campos 3 y 6 del formulario). Si seleccionás varios PDFs, la nómina muestra las instituciones detectadas. |
| 🧾 **Lectura de tablas** | Con `pdfplumber` recorre todas las páginas y tablas del cuadro; descarta encabezados, filas vacías y vacantes. |
| 🔑 **Sin duplicados** | La **cédula** (normalizada, sin ceros a la izquierda) es la clave única: los rubros y turnos repetidos de una misma persona se consolidan en una fila. |
| 🕐 **Turnos** | Se listan todos los turnos de la persona en orden M → T → N → TN (ej. `M / T / N`). |
| 🏷️ **Categorías** | Si una persona tiene varios rubros se muestran todos (ej. `LCE / LCH`). El rótulo `Docente Director` se descarta si existe un rubro real. |
| 🔤 **Nombres** | Separación automática de `NOMBRES APELLIDOS` → `APELLIDOS, NOMBRES`, respetando los conectores `DE` / `DEL` (ej. `GARCIA DE SUAREZ, ANA LUCIA`). |
| ↕️ **Orden** | Lista ordenada alfabéticamente por apellido. |

---

## 📤 Salida generada

Hoja **`NOMINA DOCENTE 2026`** con estas columnas:

| N° | APELLIDO Y NOMBRE | CEDULA DE IDENTIDAD | TURNO | CATEGORIA DE SUELDO |
|----|-------------------|--------------------:|:-----:|:-------------------:|
| 1  | PEREZ GOMEZ, JUAN CARLOS | 1234567 | M | ZZ2 |
| 2  | LOPEZ, MARIA ELENA | 7654321 | M | LCE |
| 3  | GARCIA DE SUAREZ, ANA LUCIA | 4444444 | T | LCE |

- 📌 Cabecera institucional con el nombre y la dirección extraídos del PDF (si hay varios, se muestran ambas instituciones).
- 📌 Filas de altura fija, bordes y encabezado azul con texto blanco.

---

## ✏️ Casos especiales (nombres)

La separación automática de nombres no puede decidir en casos **realmente ambiguos** (nombres compuestos o apellidos compuestos). Para eso existe el archivo **`manual_nombres.py`** (opcional, local), que asocia cada **cédula** con su formato correcto:

```python
MANUAL = {
    "1234567": "GARCIA LOPEZ, MARIA ELENA",
    "7654321": "RUIZ DIAZ MONTIEL, JUAN PEDRO",
    ...
}
```

> 🔒 Este archivo contiene datos personales, así que está **excluido del repositorio** por `.gitignore`. Si no existe, el script funciona igual usando solo la separación automática.

### ➕ Cómo agregar un caso nuevo

1. 🔍 Verificar en el Excel generado cómo quedó formateado el nombre.
2. ✍️ Si está mal, agregar en `manual_nombres.py` la cédula y el formato correcto (crear el archivo si no existe).
3. 🔁 Volver a ejecutar el programa.

> 💡 **Tip:** si la persona no está en el diccionario y no es ambigua, el script la formatea solo — no hace falta tocar nada.

---

## ⚠️ Notas y advertencias

- 📄 **Los PDFs son la única fuente de datos.** Lo que no figure en el cuadro de personal no aparecerá en la nómina. El script no guarda memoria de ejecuciones anteriores.
- 🔄 Si se actualizan los PDFs con el mismo nombre, **se usa la versión más reciente** automáticamente.
- 🚫 Las filas con cédula `VACANTE` o sin cédula válida se descartan.
- 🧩 Si un docente nuevo tiene un nombre compuesto raro, revisar la primera ejecución y ajustar `manual_nombres.py` si hace falta.
- 📦 Si editás `manual_nombres.py` **después de compilar el EXE**, hay que recompilar con `build_exe.bat` para que el cambio llegue al EXE.
- 📄 Si volvés a generar la nómina, cerrá el Excel antes de reemplazar el archivo.
- 🏫 Si el PDF no tiene el encabezado estándar del Cuadro de Personal, la nómina se genera con una cabecera genérica `NÓMINA DE DOCENTES`.
- 🖥️ El script está pensado para Windows (por el `.bat`), pero funciona igual en Linux/macOS con `python generar_nomina.py`.
- 🔒 **Privacidad:** los PDFs y los Excel generados contienen **datos personales** (nombres y cédulas). Por eso el `.gitignore` excluye todos los `*.pdf` y `*.xlsx` — **nunca subirlos al repositorio**.

---

## 🗓️ Última revisión

**28/10/2026** — **v1.1.0**: interfaz gráfica guiada de 3 pasos con nombre de la aplicación, selección múltiple de PDFs, destino, opciones, barra de progreso, resumen final con accesos para abrir el Excel/carpeta y guía de ayuda integrada (`F1`) · cabecera institucional del Excel y de la app extraída del encabezado del PDF · EXE renombrado a `NominaDocentes.exe` con metadatos de versión/autor/copyright en Windows · `generar.bat` autoinstala Python y dependencias en `.venv` · `build_exe.bat` compila EXE para PC sin Python · autoactualización del EXE vía GitHub Releases + `publicar.bat` para versionar con tags · separación de datos personales (`manual_nombres.py`) fuera del repo.
