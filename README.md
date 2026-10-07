# VetGestión — versión académica

Gestor de stock para una farmacia veterinaria: productos, stock, ventas, aumentos de
precios e historial. Aplicación de escritorio en Python, con la interfaz en HTML, CSS y
JavaScript corriendo dentro de una ventana propia (sin barra del navegador).

> Proyecto académico. Los datos que trae son de demostración, no son de un local real.

## Solo quiero verlo funcionando

Descargá **[VetGestion.exe](https://github.com/Cmiloos/vetgestion-academico/releases/latest)**
desde la última versión publicada, doble clic y listo. No se instala nada y no necesita
Python. Arranca con datos de demostración.

## Cómo levantarlo

Necesitás **Python 3.12** y, en Windows, el runtime de Edge WebView2 (ya viene con
Windows 10 y 11 actualizados).

```bash
git clone https://github.com/Cmiloos/vetgestion-academico.git
cd vetgestion-academico
pip install pywebview openpyxl
python app/main.py
```

Arranca en modo local: crea `datos/vetgestion.xlsx` al lado del script y lo llena con
datos de prueba en la primera ejecución. No necesita internet ni configuración.

### Cómo entrar

| Perfil | Puede | Contraseña |
|---|---|---|
| Empleado | Buscar, vender y ver la tabla de productos | sin contraseña |
| Fátima (admin) | Todo: precios, altas, bajas e historial | `1234` |

## Conectarlo a tu propia nube (Google Sheets)

El sistema guarda en un Excel local **o** en una planilla de Google, compartida entre
varias computadoras. Para usar tu propia planilla:

1. Creá una planilla nueva en Google Sheets.
2. Abrí **Extensiones → Apps Script**, borrá lo que haya y pegá el contenido de
   [`app/ui/apps_script.gs`](app/ui/apps_script.gs). Guardá.
3. **Implementar → Nueva implementación → Aplicación web**, con *Ejecutar como: Yo* y
   *Quién tiene acceso: Cualquier persona* (la última opción de la lista, la que **no**
   dice "con una cuenta de Google").
4. Autorizá los permisos. Si aparece "Google no verificó esta app",
   entrá por **Configuración avanzada → Ir a...**.
5. Copiá la URL que termina en `/exec` y pegala en el programa, en
   **Administración → Base de datos**.

El programa guarda esa URL solo en `datos/config.json`, que no se versiona: tu planilla
queda en tu máquina. Si no configurás nada, todo funciona contra el Excel local.

Las fotos de los productos se suben a tu Drive y la planilla guarda solo la referencia
como `drive:<id>`.

## Generar el ejecutable

```
build.bat
```

Deja un único **`VetGestion.exe`** en la raíz, sin carpetas al lado (usa PyInstaller, que
se instala solo). El ejecutable guarda sus datos en `%LOCALAPPDATA%\VetGestion`.

Si querés que el .exe salga ya conectado a tu planilla sin que nadie tenga que pegar la
URL, copiá `app/config.example.json` a `app/config.json` con tu `sheets_url` antes de
compilar: `build.bat` lo embebe en el ejecutable. Ese archivo no se versiona.

Cerrá VetGestión antes de compilar.

## Estructura

```
app/main.py            punto de entrada: crea la ventana y expone la API a JavaScript
app/api.py             métodos que llama el front con window.pywebview.api.<método>()
app/db.py              lectura y escritura: Google Sheets o Excel local
app/ui/index.html      la interfaz
app/ui/app.js          la lógica del front
app/ui/estilos.css     los estilos
app/ui/apps_script.gs  el código que se pega en Google Apps Script
app/ui/img/            fotos de los productos de demostración
build.bat              genera VetGestion.exe
docs/                  el manual de la propuesta
```

Los datos locales (`datos/`), el ejecutable y `app/config.json` no se versionan.

## Documentación

- [`docs/manual_propuesta.docx`](docs/manual_propuesta.docx) — manual de la propuesta.
- `CLAUDE.md` — notas del proyecto; si abrís la carpeta con Claude Code, las toma solo.
