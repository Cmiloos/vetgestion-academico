# VetGestión — desktop app

Veterinary pharmacy stock manager (products, stock, sales, price increases).
UI language is Spanish (Argentina, "vos" form); keep all user-facing text in Spanish.
Data is fake/demo — this is a school/demo project, not production.

## Token-saving rules (read first)
- NEVER Read `vetgestion (1) (1).html` in full: lines ~647-658 hold base64 images
  (18-170 KB each, ~640 KB total). It is the original prototype, kept only as reference.
  Use Grep or Read with offset/limit and skip that range.
- Never Read/Grep inside `app/ui/img/`, `datos/imagenes/`, `build/`, `dist/` or `*.xlsx`.
  To inspect the Excel data, run a short Python/openpyxl snippet that prints rows.
- Prefer Grep for a function name over reading whole files; edit with small Edits.
- Don't re-read a file right after editing it; don't restate code in replies.
- Keep replies short. Answer the user in Spanish.

## Stack
- Python 3.12 + pywebview (Edge WebView2 window) — the HTML/CSS/JS UI runs inside it.
- Storage (`db.py`, same sheets/columns in both): `SheetsStore` = Google Sheets via an Apps Script
  web app (`app/ui/apps_script.gs`, HTTP POST JSON; photos go to Drive as `drive:<id>`), or
  `ExcelStore` = local xlsx via openpyxl (offline fallback). Chosen by `datos/config.json` ("sheets_url").
- Packaging: PyInstaller -> single `VetGestion.exe`.

## Layout
- `app/main.py`   — entry point; creates the window and exposes `Api` to JS.
- `app/api.py`    — methods JS calls via `window.pywebview.api.<method>()` (all return JSON-safe dicts).
- `app/db.py`     — Excel read/write; creates + seeds the workbook on first run.
- `app/ui/index.html` markup · `app/ui/app.js` logic · `app/ui/estilos.css` (prototype CSS, rarely edit).
- `app/ui/img/`   — seed product images (extracted from the prototype). `app/icono.ico` exe icon.
- `datos/vetgestion.xlsx` — the database, created next to the exe/script on first run.
  Sheets: Productos (all product + ficha fields), Movimientos (sales, price changes, stock, alta/baja).
- `datos/imagenes/` — product photos; Excel stores only the filename.
- `build.bat` — builds a single `VetGestion.exe` in the project root (no dist/). Bundles `app/config.json`.
  The exe keeps its data in `%LOCALAPPDATA%\VetGestion` (dev runs use `datos/`). Close VetGestión before building.
- Test Sheets logic without Google: a local mock of the Apps Script protocol (302 redirect + JSON).
- Speed: UI is optimistic (app.js `enSegundoPlano` queue); startup shows `datos/copia_nube.json` then refreshes.
  Script v2 (`VERSION = 2`) applies stock/price deltas server-side and returns fresh data in one call.

## Conventions
- Product code format `NNN-NNNN`; money is integer ARS, shown as `$12.345`.
- Admin user "Fátima", password `1234`; "Empleado" can only search/sell/view the table.
- If the Excel is open in MS Excel, saving fails: catch PermissionError and show a
  Spanish message asking to close the file.

## Run / build
- Dev: `python app/main.py`
- Build: `build.bat` (needs `pip install pywebview openpyxl pyinstaller`)
