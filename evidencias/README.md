# Robot de causaciones SICOF

Automatización en Python con Playwright para leer facturas desde Excel,
consultar si ya están causadas en SICOF y registrar las que aún no existan.

## Estructura del proyecto

```text
robot_causaciones/
├── datos/
│   └── Santiago_Giraldo_por_radicar.xlsx
├── evidencias/
├── salida/
│   ├── bot.log
│   ├── resultados.xlsx
│   └── evidencias/
├── main.py
├── .env
├── .python-version
├── pyproject.toml
└── uv.lock
```

- `main.py`: lógica de lectura del Excel, validación y automatización de SICOF.
- `datos/`: libro de Excel de entrada.
- `evidencias/`: carpeta de evidencias que ya forma parte del proyecto.
- `salida/resultados.xlsx`: resultados guardados durante el procesamiento.
- `salida/bot.log`: registro de ejecución, también visible en la consola.
- `salida/evidencias/`: capturas de pantalla de fallos guardadas por el bot.
- `.env`: configuración local y credenciales; no compartir ni publicar este archivo.

## Requisitos

- Python 3.13 o superior.
- Chromium instalado para Playwright.
- Acceso a SICOF y credenciales válidas.
- `uv` para instalar el entorno según `uv.lock` (alternativamente, `pip`).

## Instalación

Abre PowerShell en la carpeta `robot_causaciones` y ejecuta:

```powershell
uv sync
uv run playwright install chromium
```

Si no usas `uv`, instala las dependencias declaradas en `pyproject.toml` y el
navegador de Playwright:

```powershell
python -m pip install openpyxl playwright python-dotenv
python -m playwright install chromium
```

## Configuración

Completa el `.env` con tu configuración local:

```dotenv
SICOF_URL=http://127.0.0.1:8020
SICOF_USUARIO=tu_usuario
SICOF_CLAVE=tu_clave
ARCHIVO_ENTRADA=datos/Santiago_Giraldo_por_radicar.xlsx
ARCHIVO_SALIDA=salida/resultados.xlsx
```

`SICOF_URL` usa `http://127.0.0.1:8020` si no se especifica. Las rutas de
entrada y salida predeterminadas son relativas a la carpeta desde la que se
ejecuta el comando; por eso se recomienda ejecutar el bot desde
`robot_causaciones`.

## Formato del Excel

El archivo de entrada debe tener una hoja llamada `PorRadicar` y una fila de
encabezados. El bot busca estas columnas por nombre, ignorando mayúsculas,
espacios, tildes y otros caracteres:

| Columna | Contenido esperado |
| --- | --- |
| `IdLista` | Identificador de la fila; las filas sin identificador se omiten. |
| `Factura` | Número o identificador de factura, no el nombre del proveedor. |
| `NIT` | Identificación del proveedor; se eliminan caracteres que no sean dígitos. |
| `Proveedor` | Nombre del proveedor. |
| `Valor` | Valor positivo y entero, por ejemplo `1850000` o `1.850.000`. |
| `CentroCosto` | Código de centro de costo disponible en SICOF, por ejemplo `CC-100`. |
| `FechaFactura` | Fecha válida en Excel, o texto `AAAA-MM-DD`, `DD/MM/AAAA`, `DD-MM-AAAA`, `AAAA/MM/DD` o `DD/MM/AAAA HH:MM:SS`. También acepta fecha ISO UTC como `2026-10-18T05:00:00Z`. |

La fecha se convierte a `DD/MM/AAAA` antes de enviarse a SICOF. El libro debe
contener datos consistentes: la factura y el proveedor son campos distintos,
y el NIT debe estar registrado como proveedor en SICOF.

## Ejecución

Desde la carpeta `robot_causaciones`:

```powershell
uv run python main.py
```

También se puede ejecutar con el Python del entorno que tenga instaladas las
dependencias:

```powershell
python main.py
```

El navegador se abre en modo visible. El bot procesa las filas en orden,
comprueba si ya existe una causación para el NIT y la factura, y si no existe
intenta crearla hasta tres veces. Los errores de validación quedan marcados
como `RECHAZADA`; los fallos inesperados o tras agotar intentos quedan como
`FALLA_SISTEMA`.

## Resultados y diagnóstico

- Revisa `salida/resultados.xlsx` para el estado de cada factura. Incluye las
  columnas `IdLista`, `Factura`, `NIT`, `Estado`, `Causacion` y `Detalle`.
- Revisa `salida/bot.log` para el avance, los errores y los resultados de cada
  fila. El mismo registro se muestra en la terminal mientras el bot trabaja.
- Consulta `salida/evidencias/` si se guardó una captura para un fallo.

El libro de resultados se guarda después de cada factura para conservar los
avances si el proceso se interrumpe.