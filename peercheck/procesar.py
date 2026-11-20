"""Procesa los archivos de la carpeta INPUT y deja el resultado en OUTPUT."""
from __future__ import annotations

import logging
import shutil
from datetime import datetime
from pathlib import Path
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from .motor import buscar_pares, comparar_par
from .reporte import HOJA_RESUMEN, colorear_origen, crear_resumen

log = logging.getLogger(__name__)

EXTENSIONES = {".xlsx", ".xlsm"}
MAX_MB = 50  # evita colgarse con archivos enormes o dañados


class ErrorArchivo(Exception):
    pass


def archivos_entrada(carpeta: Path) -> list[Path]:
    return sorted(
        p for p in carpeta.iterdir()
        if p.is_file() and p.suffix.lower() in EXTENSIONES
        and not p.name.startswith("~$")  # archivos temporales de Excel abierto
    )


def ruta_salida(entrada: Path, carpeta_salida: Path) -> Path:
    base = carpeta_salida / f"{entrada.stem}_REVIEWED{entrada.suffix.lower()}"
    if not base.exists():
        return base
    sello = datetime.now().strftime("%Y%m%d_%H%M%S")
    return carpeta_salida / f"{entrada.stem}_REVIEWED_{sello}{entrada.suffix.lower()}"


def respaldar(entrada: Path, carpeta_respaldo: Path) -> Path:
    """Copia el original con fecha y hora; antes, un segundo respaldo pisaba al primero."""
    sello = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = carpeta_respaldo / f"{entrada.stem}_{sello}{entrada.suffix}"
    shutil.copy2(entrada, destino)
    return destino


def procesar_archivo(entrada: Path, carpeta_salida: Path, carpeta_respaldo: Path, cfg: dict) -> dict:
    if entrada.stat().st_size > MAX_MB * 1024 * 1024:
        raise ErrorArchivo(f"pesa más de {MAX_MB} MB")

    respaldar(entrada, carpeta_respaldo)
    es_xlsm = entrada.suffix.lower() == ".xlsm"
    try:
        # keep_vba conserva las macros del .xlsm, pero openpyxl nunca las ejecuta.
        wb = load_workbook(entrada, keep_vba=es_xlsm)
    except (InvalidFileException, BadZipFile, KeyError, OSError) as e:
        raise ErrorArchivo(f"no se pudo abrir ({type(e).__name__})") from e

    avisos: list[str] = []
    resultados = []
    for ws in wb.worksheets:
        if ws.title == HOJA_RESUMEN:
            continue
        if " ".join(ws.title.split()).upper() in cfg["hojas_ignoradas"]:
            avisos.append(f"{ws.title}: IGNORED")
            continue
        pares = buscar_pares(ws, cfg["marcador_mirror"], avisos)
        if not pares:
            avisos.append(f"{ws.title}: NO MIRROR FOUND")
            continue
        for par in pares:
            res = comparar_par(ws, par, cfg)
            colorear_origen(ws, res)
            resultados.append(res)

    crear_resumen(wb, resultados, avisos, cfg)
    salida = ruta_salida(entrada, carpeta_salida)
    try:
        wb.save(salida)
    except PermissionError:
        raise ErrorArchivo(f"no se pudo guardar {salida.name}: ¿está abierto en Excel?") from None
    finally:
        wb.close()

    coincidencias = sum(r.coincidencias for r in resultados)
    diferencias = sum(r.diferencias for r in resultados)
    total = coincidencias + diferencias
    return {
        "salida": salida,
        "pares": len(resultados),
        "criticas": sum(r.criticas for r in resultados),
        "diferencias": diferencias,
        "puntaje": round(coincidencias / total * 100, 2) if total else 0.0,
        "avisos": len(avisos),
    }


def procesar_carpeta(base: Path, cfg: dict) -> int:
    entrada, salida, respaldo = base / "INPUT", base / "OUTPUT", base / "BACKUP"
    for c in (entrada, salida, respaldo):
        c.mkdir(parents=True, exist_ok=True)

    archivos = archivos_entrada(entrada)
    if not archivos:
        log.info("No hay archivos .xlsx o .xlsm en %s", entrada)
        return 0

    errores = 0
    for archivo in archivos:
        try:
            r = procesar_archivo(archivo, salida, respaldo, cfg)
            log.info("OK  %s -> %s | pares: %d | score: %.2f%% | críticos: %d | avisos: %d",
                     archivo.name, r["salida"].name, r["pares"], r["puntaje"], r["criticas"], r["avisos"])
        except ErrorArchivo as e:
            errores += 1
            log.error("ERROR %s: %s", archivo.name, e)
    log.info("Terminado: %d archivo(s), %d con error", len(archivos), errores)
    return errores
