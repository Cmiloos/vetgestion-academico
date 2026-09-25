"""Funciones que la interfaz llama con window.pywebview.api.<nombre>(...).

Todas devuelven diccionarios simples: {"ok": True, ...} o {"ok": False, "error": "..."}.
"""
import base64
import mimetypes
import os
import random
import re
import time
import webbrowser
from functools import wraps

import db

CLAVE_ADMIN = "1234"
EXT_IMG = {"image/jpeg": "jpg", "image/png": "png", "image/webp": "webp", "image/gif": "gif"}


def _seguro(fn):
    """Convierte cualquier error en {'ok': False, 'error': mensaje} para mostrarlo en pantalla."""
    @wraps(fn)
    def envoltura(*a, **k):
        try:
            res = fn(*a, **k)
            return {"ok": True, **(res or {})}
        except (ValueError, db.ExcelAbierto) as e:
            return {"ok": False, "error": str(e)}
        except Exception as e:  # error inesperado: se muestra igual para no colgar la UI
            return {"ok": False, "error": f"Error inesperado: {e}"}
    return envoltura


def _entero(v, nombre, minimo=0):
    try:
        n = int(round(float(str(v).replace(",", "."))))
    except (TypeError, ValueError):
        raise ValueError(f"{nombre}: tiene que ser un número.")
    if n < minimo:
        raise ValueError(f"{nombre}: tiene que ser {minimo} o más.")
    return n


def _porcentaje(v):
    try:
        return float(str(v).replace("%", "").replace(",", ".").strip())
    except ValueError:
        raise ValueError(f"Porcentaje inválido: {v}")


class Api:
    def __init__(self):
        self._cache_img = {}

    # ------------------------------------------------------------ lectura
    @_seguro
    def estado(self):
        return self._armar_estado(*db.leer_todo(None))

    @_seguro
    def estado_guardado(self):
        """Lo último que se leyó de la nube, sin esperar a Google (para abrir al instante)."""
        copia = db.leer_copia()
        if not copia:
            raise ValueError("No hay copia guardada.")
        return {**self._armar_estado(*copia), "copia": True}

    def _armar_estado(self, prods, movs):
        hoy = time.strftime("%Y-%m-%d")
        ventas = [m for m in movs if m["tipo"] == "Venta" and str(m["fecha"]).startswith(hoy)]
        return {"productos": prods, "movimientos": movs[:40], "almacen": db.info_almacen(),
                "stockBajo": db.STOCK_BAJO, "diasPronto": db.DIAS_PRONTO,
                "ventasHoy": {"importe": sum(int(m["importe"] or 0) for m in ventas),
                              "unidades": sum(int(m["cantidad"] or 0) for m in ventas)}}

    @_seguro
    def movimientos(self):
        return {"movimientos": db.leer_todo(None)[1]}

    def imagen(self, nombre):
        """Devuelve la imagen como data URL (o '' si no existe)."""
        if not nombre:
            return ""
        try:
            local = db.imagen_local(nombre)
        except ValueError:  # sin conexión o la foto ya no está en Drive
            return ""
        if not local:
            return ""
        ruta = os.path.join(db.IMG_DIR, local)
        clave = (ruta, os.path.getmtime(ruta))
        if clave not in self._cache_img:
            tipo = mimetypes.guess_type(ruta)[0] or "image/jpeg"
            if ruta.lower().endswith(".webp"):
                tipo = "image/webp"
            with open(ruta, "rb") as f:
                self._cache_img[clave] = f"data:{tipo};base64,{base64.b64encode(f.read()).decode()}"
        return self._cache_img[clave]

    def login(self, clave):
        return clave == CLAVE_ADMIN

    # ------------------------------------------------------------ ventas y stock
    @_seguro
    def vender(self, codigo, cantidad, usuario):
        return db.vender(codigo, _entero(cantidad, "Cantidad", 1), usuario)

    @_seguro
    def reponer(self, codigo, cantidad, usuario):
        return db.ajustar_stock(codigo, _entero(cantidad, "Cantidad", 1),
                                "Reposición de stock", "", usuario)

    @_seguro
    def retirar(self, codigo, cantidad, motivo, usuario):
        return db.ajustar_stock(codigo, -_entero(cantidad, "Cantidad", 1),
                                "Baja por vencimiento" if motivo == "vencido" else "Retiro de stock",
                                "Retiro de la venta", usuario)

    # ------------------------------------------------------------ precios
    @_seguro
    def aumentar(self, laboratorio, categoria, pct, usuario):
        pct = _porcentaje(pct)
        if pct == 0:
            raise ValueError("El porcentaje tiene que ser distinto de cero.")
        return {"resumen": db.aumentar([(laboratorio, categoria or "", pct)], usuario)}

    @_seguro
    def importar_csv(self, texto, aplicar, usuario):
        """CSV: laboratorio;porcentaje[;categoría]. Con aplicar=False solo valida y devuelve las reglas."""
        prods, _ = db.leer_todo(0)
        labs = {p["laboratorio"].lower(): p["laboratorio"] for p in prods}
        texto = texto.lstrip("﻿")  # los CSV "UTF-8" de Excel empiezan con esta marca invisible
        reglas, errores = [], []
        for n, linea in enumerate(texto.splitlines(), start=1):
            if not linea.strip():
                continue
            partes = [x.strip().strip('"') for x in re.split(r"[;,\t]", linea)]
            if len(partes) > 2 and re.fullmatch(r"\d+", partes[2] or "x"):
                # "PetLab,5,5" → coma decimal usada como separador
                partes = [partes[0], partes[1] + "." + partes[2]] + partes[3:]
            if len(partes) < 2:
                errores.append(f"Línea {n}: faltan columnas")
                continue
            try:
                pct = _porcentaje(partes[1])
            except ValueError:
                if n == 1:
                    continue  # encabezado
                errores.append(f"Línea {n}: porcentaje inválido ({partes[1]})")
                continue
            lab = labs.get(partes[0].lower())
            if not lab:
                errores.append(f"Línea {n}: no existe el laboratorio «{partes[0]}»")
                continue
            cat = partes[2] if len(partes) > 2 else ""
            afect = [p for p in prods if p["laboratorio"] == lab and (not cat or p["categoria"] == cat)]
            reglas.append({"laboratorio": lab, "categoria": cat, "pct": pct, "productos": len(afect)})
        if not reglas:
            raise ValueError("El archivo no tiene filas válidas. " + "; ".join(errores[:3]))
        if not aplicar:
            return {"reglas": reglas, "errores": errores}
        res = db.aumentar([(r["laboratorio"], r["categoria"], r["pct"]) for r in reglas], usuario)
        return {"resumen": res, "errores": errores}

    # ------------------------------------------------------------ administración
    @_seguro
    def guardar_producto(self, datos, foto, codigo_original, usuario):
        """foto: None = sin cambios, '' = quitar, 'data:...' = nueva imagen."""
        d = {k: str(datos.get(k, "") or "").strip() for k, _, _ in db.COLS_PROD}
        if not d["nombre"]:
            raise ValueError("Completá el nombre del producto.")
        if not d["laboratorio"] or not d["categoria"]:
            raise ValueError("Elegí o escribí el laboratorio y la categoría.")
        d["precio"] = _entero(datos.get("precio"), "Precio de venta", 1)
        d["stock"] = _entero(datos.get("stock") or 0, "Stock", 0)
        if d["vencimiento"]:
            m = re.fullmatch(r"(\d{1,2})\s*/\s*(\d{4})", d["vencimiento"])
            if not m or not 1 <= int(m.group(1)) <= 12:
                raise ValueError("El vencimiento tiene que tener el formato MM/AAAA (por ejemplo 05/2027).")
            d["vencimiento"] = f"{int(m.group(1)):02d}/{m.group(2)}"
        if d["codigo"] and not re.fullmatch(r"[\w-]{3,20}", d["codigo"]):
            raise ValueError("El código solo puede tener letras, números y guiones.")
        if not d["codigo"]:
            d["codigo"] = self._nuevo_codigo(d["laboratorio"])
        d["droga"] = d["droga"] or "—"

        if codigo_original and foto is None:
            d.pop("imagen")  # se conserva la que tenía
        elif foto:
            d["imagen"] = self._guardar_foto(d["codigo"], foto)
        else:
            d["imagen"] = ""
        db.guardar_producto(d, usuario, codigo_original or None)
        return {"codigo": d["codigo"], "nombre": d["nombre"], "precio": d["precio"]}

    @_seguro
    def baja(self, codigo, usuario):
        return {"nombre": db.baja(codigo, usuario)}

    # ------------------------------------------------------------ base de datos
    @_seguro
    def abrir_base(self):
        info = db.info_almacen()
        if info["modo"] == "sheets":
            if not info["link"]:
                raise ValueError("No se conoce el link de la planilla. Volvé a conectarla.")
            webbrowser.open(info["link"])
        else:
            os.startfile(db.XLSX)

    @_seguro
    def conectar_sheets(self, url):
        return db.conectar_sheets(url or "")

    @_seguro
    def desconectar_sheets(self):
        return db.desconectar_sheets()

    def codigo_script(self):
        with open(os.path.join(db._recursos_dir(), "ui", "apps_script.gs"), encoding="utf-8") as f:
            return f.read()

    # ------------------------------------------------------------ auxiliares
    def _nuevo_codigo(self, laboratorio):
        prods, _ = db.leer_todo(0)
        existentes = {p["codigo"] for p in prods}
        prefijos = {p["codigo"][:3] for p in prods if p["laboratorio"] == laboratorio and p["codigo"][:3].isdigit()}
        if prefijos:
            prefijo = sorted(prefijos)[0]
        else:
            nums = [int(p["codigo"][:3]) for p in prods if p["codigo"][:3].isdigit()]
            prefijo = f"{(max(nums) + 1) if nums else 1:03d}"
        while True:
            cod = f"{prefijo}-{random.randint(1000, 9999)}"
            if cod not in existentes:
                return cod

    def _guardar_foto(self, codigo, data_url):
        m = re.match(r"data:(image/[\w+.-]+);base64,(.+)", data_url, re.S)
        if not m:
            raise ValueError("La imagen no tiene un formato válido.")
        ext = EXT_IMG.get(m.group(1))
        if not ext:
            raise ValueError("Formato de imagen no soportado (usá JPG, PNG o WEBP).")
        nombre = f"{codigo}-{int(time.time())}.{ext}"
        datos = base64.b64decode(m.group(2))
        almacen = db.store()
        if almacen.modo == "sheets":  # a Drive, para que se vea desde cualquier computadora
            id_drive = almacen.subir_imagen(nombre, m.group(1), datos)
            nombre, local = "drive:" + id_drive, f"drive_{id_drive}.{ext}"
        else:
            local = nombre
        with open(os.path.join(db.IMG_DIR, local), "wb") as f:
            f.write(datos)
        return nombre
