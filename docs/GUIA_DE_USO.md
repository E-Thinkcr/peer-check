# Guía de uso — Peer Check

Esta guía explica, paso a paso y sin necesidad de saber programar, cómo preparar los archivos de Excel, correr Peer Check y leer el resultado.

**Contenido**

1. [Qué hace](#1-qué-hace)
2. [Instalación (una sola vez)](#2-instalación-una-sola-vez)
3. [Preparar la carpeta de trabajo](#3-preparar-la-carpeta-de-trabajo)
4. [Preparar el Excel](#4-preparar-el-excel)
5. [Correr la revisión](#5-correr-la-revisión)
6. [Leer el resultado](#6-leer-el-resultado)
7. [Configurar los campos](#7-configurar-los-campos-configjson)
8. [Probar con el ejemplo](#8-probar-con-el-ejemplo)
9. [Preguntas frecuentes](#9-preguntas-frecuentes)

---

## 1. Qué hace

Un analista captura una configuración (el **Agent**) y un colega la vuelve a capturar por separado como control (el **Mirror**). Peer Check compara las dos capturas **campo por campo** y genera un Excel con:

- las diferencias marcadas en colores, en las hojas originales;
- una hoja **"Peer Check Summary"** con el detalle de cada par, un puntaje y un tablero general.

---

## 2. Instalación (una sola vez)

1. Instala **Python 3.10 o superior** desde [python.org](https://www.python.org/downloads/). Durante la instalación marca **"Add Python to PATH"**.
2. Descarga el proyecto: en GitHub, botón verde **Code → Download ZIP**, y descomprímelo. Por ejemplo, en `C:\PeerCheck\programa`.
3. Abre la carpeta del proyecto, escribe `cmd` en la barra de direcciones del Explorador y presiona Enter. En la ventana negra escribe:

```bash
pip install -r requirements.txt
```

Con eso queda listo.

---

## 3. Preparar la carpeta de trabajo

Crea una carpeta para tus revisiones, por ejemplo `C:\PeerCheck\trabajo`, con esta estructura:

```
C:\PeerCheck\trabajo\
├── config.json      ← tus reglas (ver sección 7)
├── peercheck.bat    ← copia del archivo del proyecto (doble clic para correr)
├── INPUT\           ← aquí pones los Excel a revisar
├── OUTPUT\          ← aquí aparecen los resultados
└── BACKUP\          ← copia de seguridad automática de cada original
```

Las carpetas `INPUT`, `OUTPUT` y `BACKUP` se crean solas la primera vez que corre.

---

## 4. Preparar el Excel

Se aceptan archivos **`.xlsx`** y **`.xlsm`** (con macros). Cada hoja que se quiera revisar debe tener este formato:

| Fila | Columna A | Columnas B, C, D… |
|---|---|---|
| 1 | Encabezado del ID (por ejemplo, `ID`) | **Nombres de los campos** |
| 2 a N | ID de cada registro capturado (**Agent 1, 2, 3…**) | Valores capturados |
| Siguiente | La palabra **`MIRROR`** | (vacío) |
| Después | ID de cada registro de control (**Mirror 1, 2, 3…**) | Valores de control |

**Reglas importantes:**

- Los mirrors deben ir **en el mismo orden** que los agents: el Mirror 1 se compara con el Agent 1, y así sucesivamente.
- Debe haber la **misma cantidad** de agents que de mirrors. Si no cuadran, solo se comparan los primeros y aparece una advertencia.
- Una hoja sin la palabra `MIRROR` no se revisa y aparece como advertencia `NO MIRROR FOUND`.
- Mayúsculas, espacios de más, `$`, `%` y comas **no cuentan como diferencia**.

**Ejemplo:**

| ID | NAME | NETWORK | BRANCH | STATUS |
|---|---|---|---|---|
| LOC-1001 | Pulpería La Esquina | NORTE | 0012 | ACTIVE |
| LOC-1002 | Farmacia Central | CENTRO | 0045 | ACTIVE |
| **MIRROR** | | | | |
| LOC-1001 | Pulpería La Esquina | NORTE | 12 | Active |
| LOC-1002 | Farmacia Central | CENTRO | 0045 | ACTIVE |

---

## 5. Correr la revisión

1. Copia los Excel a la carpeta **`INPUT`**.
2. **Cierra esos archivos en Excel.** Si están abiertos, no se pueden guardar los resultados.
3. Corre la revisión de una de estas dos formas:
   - **Doble clic en `peercheck.bat`.** La primera vez, abre `peercheck.bat` con el Bloc de notas y cambia la línea `set "PROGRAMA=%~dp0"` por la carpeta donde descargaste el proyecto. Por ejemplo: `set "PROGRAMA=C:\PeerCheck\programa"`.
   - **Desde la ventana de comandos:**

```bash
python -m peercheck --carpeta "C:\PeerCheck\trabajo"
```

4. La ventana muestra una línea por archivo, por ejemplo:

```
OK  setup_noviembre.xlsx -> setup_noviembre_REVIEWED.xlsx | pares: 5 | score: 88.24% | críticos: 3 | avisos: 1
```

5. Abre el resultado en la carpeta **`OUTPUT`**. El archivo se llama igual que el original, terminado en **`_REVIEWED`**.

> El archivo original **no se modifica**. Antes de procesarlo se guarda una copia con fecha y hora en `BACKUP`. Si corres la revisión dos veces, el segundo resultado también lleva fecha y hora, así que nunca se pisa el anterior.

---

## 6. Leer el resultado

### Hoja "Peer Check Summary"

**Tablero ejecutivo (arriba):**

| Indicador | Significado |
|---|---|
| **PAIRS REVIEWED** | Cuántos pares Agent ↔ Mirror se revisaron. |
| **MATCHES** | Campos que coinciden. |
| **DISCREPANCIES** | Campos con diferencias, incluidos los críticos. |
| **CRITICAL** | Diferencias en campos críticos. Se marca en rojo si hay alguna. |
| **OVERALL SCORE** | Coincidencias ÷ campos comparados. |
| **WARNINGS** | Cantidad de advertencias. El detalle aparece al final de la hoja. |

**Detalle por par:** el ID del agent, el del mirror y la hoja. Debajo va una tabla con cada campo, los dos valores y el resultado:

| Resultado | Color | Qué hacer |
|---|---|---|
| `MATCH` | 🟩 Verde | Nada, está correcto. |
| `DISCREPANCY` | 🟥 Rojo | Revisar cuál de los dos valores es el correcto. |
| `CRITICAL` | 🟥 Rojo | **Corregir antes de aprobar**: es un campo crítico. |
| `EXCLUDED` | ⬜ Gris | No se compara, por regla de `config.json`. |
| `EMPTY` | ⬜ Gris | Vacío en las dos capturas. |

**PAIR SCORE:** el puntaje de cada par, con semáforo:

- 🟩 **95 % o más:** aprobado.
- 🟨 **80 % a 94.99 %:** revisar.
- 🟥 **Menos de 80 %:** rehacer o revisar a fondo.

> Los campos marcados como **sensibles** (por ejemplo, cuentas bancarias) se comparan completos, pero en el resumen solo se ven los últimos 4 dígitos: `•••••6789`.

### Hojas originales

Las celdas revisadas quedan coloreadas con los mismos colores, para ver cada diferencia en su contexto.

### Advertencias comunes

| Advertencia | Qué significa |
|---|---|
| `NO MIRROR FOUND` | La hoja no tiene la palabra `MIRROR` en la columna A. |
| `IGNORED` | La hoja está en `hojas_ignoradas`. |
| `2 agent(s) vs 1 mirror(s)` | Falta o sobra un registro en alguna de las dos capturas. |
| `IDs distintos` | Los IDs de un par no coinciden: probablemente las filas están en otro orden. |

---

## 7. Configurar los campos (`config.json`)

Copia **`config.example.json`** a tu carpeta de trabajo, renómbralo **`config.json`** y ábrelo con el Bloc de notas:

```json
{
  "marcador_mirror": "MIRROR",
  "campos_criticos": ["BANK ACCOUNT", "ROUTING NUMBER", "PAYMENT METHOD"],
  "campos_excluidos": ["ID", "EMAIL", "PHONE", "START DATE"],
  "campos_sensibles": ["BANK ACCOUNT", "ROUTING NUMBER"],
  "hojas_ignoradas": ["INSTRUCTIONS"],
  "ignorar_ceros_izquierda": true,
  "umbral_verde": 95,
  "umbral_amarillo": 80
}
```

| Opción | Qué hace |
|---|---|
| `marcador_mirror` | Palabra que separa las dos capturas. |
| `campos_criticos` | Si tienen diferencias, se marcan como `CRITICAL`. |
| `campos_excluidos` | No se comparan: IDs, fechas, datos de contacto… |
| `campos_sensibles` | Se ocultan en el resumen y solo se ven los últimos 4 caracteres. |
| `hojas_ignoradas` | Hojas que no se revisan. |
| `ignorar_ceros_izquierda` | `true`: `0012` = `12`. `false`: los códigos deben ser idénticos. |
| `umbral_verde` / `umbral_amarillo` | Límites del semáforo. |

**Consejos:**

- Escribe los nombres de los campos **igual que en el encabezado del Excel**. No importan las mayúsculas ni los espacios de más.
- Respeta las comillas `"` y las comas entre elementos. Si te equivocas, Peer Check te dice en qué línea está el error.
- **No subas tu `config.json` a internet**: puede contener nombres internos de tu empresa.

---

## 8. Probar con el ejemplo

El proyecto trae un generador de un Excel con datos **ficticios**:

```bash
python ejemplo/generar_ejemplo.py
python -m peercheck --carpeta ejemplo --config config.example.json
```

El resultado queda en `ejemplo\OUTPUT\ejemplo_setup_REVIEWED.xlsx`.

---

## 9. Preguntas frecuentes

**"No hay archivos .xlsx o .xlsm en …\INPUT".**
Copia los archivos a la carpeta `INPUT` de la carpeta que indicaste con `--carpeta`.

**"No se pudo guardar …: ¿está abierto en Excel?"**
Cierra el archivo en Excel y vuelve a correr la revisión.

**"No se encontró config.json: se usan los valores por defecto".**
No hay `config.json` en tu carpeta. Funciona igual, pero sin campos críticos, excluidos ni sensibles.

**"Configuración inválida: …"**
Hay un error en `config.json`. El mensaje indica la línea o la opción con el problema.

**¿Se ejecutan las macros de mis archivos `.xlsm`?**
No. Las macros se conservan en el resultado, pero Peer Check nunca las ejecuta.

**Un campo sale como diferencia aunque se ve igual.**
Revisa si hay caracteres distintos, como letras con tilde (`Súper` vs. `Super`) o guiones diferentes. Los espacios, las mayúsculas y los símbolos `$ % ,` ya se ignoran.

**Un código con ceros a la izquierda sale como MATCH aunque no debería.**
Cambia `"ignorar_ceros_izquierda"` a `false` en `config.json`.
