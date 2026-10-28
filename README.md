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
5. 📤 **Genera** el Excel final con el estilo institucional (encabezado azul, bordes, Times New Roman/Calibri).

---

## 📁 Estructura del proyecto

```text
NOMINA_DOCENTES/
├── 📄 generar_nomina.py              # Script principal (toda la lógica)
├── 🖱️ generar.bat                    # PC nueva: instala todo + abre el programa
├── 🔨 build_exe.bat                  # Compila un EXE para PC sin Python
├── 🚢 publicar.bat                   # Versiona (tag) y publica el EXE en GitHub Releases
├── 🔢 bump_version.py                # Actualiza la constante VERSION
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
2. 📦 Espera a que compile: queda en **`dist\NOMINA DOCENTES.exe`** (un único archivo).
3. 📋 Copia ese EXE a cualquier PC (pendrive, correo, red) — **no necesita Python ni instalación**.
4. 🖱️ Doble clic en el EXE y seguí el flujo de abajo.

### 🖼️ Flujo del programa (con ventanas gráficas)

1. 📥 Se abre una ventana para **seleccionar el/los cuadro(s) de personal en PDF** — puedes marcar varios con `Ctrl` o `Shift` (ej. el de la Escuela y el del Colegio juntos).
2. 💾 Luego pide **dónde guardar el Excel** (elegís carpeta y nombre; propone `NOMINA DOCENTES MM-AAAA.xlsx`).
3. ✅ Al terminar muestra un resumen con la cantidad de docentes y catedráticos, y el archivo queda listo para usar.

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
   - ❌ **No** → continúa normalmente con la versión actual.
3. 📴 Si no hay internet o falla la consulta, **abre directamente sin molestar**.

> ℹ️ El chequeo funciona si el repositorio es **público** (la API de GitHub no pide token).

## 🏷️ Publicar una nueva versión

Para desarrolladores (en la PC de desarrollo, con el repo clonado):

```bash
:: 1. Sube VERSION en el código, compila el EXE, crea el tag vX.Y.Z
::    y publica el Release con el binario (requiere GitHub CLI "gh")
publicar.bat 1.0.1
```

Lo que hace por dentro:

| Paso | Acción |
|------|--------|
| 1️⃣ | Actualiza `VERSION = "x.y.z"` en `generar_nomina.py` (con `bump_version.py`) |
| 2️⃣ | Compila `dist\NOMINA DOCENTES.exe` con PyInstaller |
| 3️⃣ | Crea el tag `vx.y.z` y lo sube al repo |
| 4️⃣ | Crea el **Release** en GitHub adjuntando el EXE |

A partir del siguiente arranque, todos los EXE viejos detectan el release nuevo y ofrecen actualizarse. 🎉

---

## 🧠 Cómo funciona

| Paso | Detalle |
|------|---------|
| 🔎 **Selección de PDFs** | En modo gráfico elegís uno o varios PDFs de cualquier carpeta; en modo consola (`--cli`) busca recursivamente (hasta 2 niveles) archivos `.pdf` con `ESCUELA` / `COLEGIO` en el nombre y usa el más reciente de cada uno. |
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

- 📌 Cabecera institucional con el nombre de la escuela/colegio y el estilo de la nómina oficial.
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
- 🖥️ El script está pensado para Windows (por el `.bat`), pero funciona igual en Linux/macOS con `python generar_nomina.py`.
- 🔒 **Privacidad:** los PDFs y los Excel generados contienen **datos personales** (nombres y cédulas). Por eso el `.gitignore` excluye todos los `*.pdf` y `*.xlsx` — **nunca subirlos al repositorio**.

---

## 🗓️ Última revisión

**28/10/2026** — Modo gráfico: selección de uno o varios PDFs y destino del Excel · `generar.bat` autoinstala Python y dependencias en `.venv` · `build_exe.bat` compila EXE para PC sin Python · autoactualización del EXE vía GitHub Releases + `publicar.bat` para versionar con tags · separación de datos personales (`manual_nombres.py`) fuera del repo.
