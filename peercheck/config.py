"""Configuración: qué campos son críticos, cuáles se excluyen y cuáles se enmascaran.

Las listas reales de campos de tu trabajo van en un `config.json` local (no se
sube a GitHub). `config.example.json` trae nombres genéricos de ejemplo.
"""
from __future__ import annotations

import json
import logging
from pathlib import Path

log = logging.getLogger(__name__)

DEFAULTS: dict = {
    # Texto de la columna A que separa los registros originales de los de control.
    "marcador_mirror": "MIRROR",
    # Diferencias en estos campos se marcan como CRITICAL.
    "campos_criticos": [],
    # Estos campos no se comparan (IDs, fechas, datos de contacto...).
    "campos_excluidos": [],
    # Estos campos se comparan completos, pero en el resumen solo se muestran los últimos 4 caracteres.
    "campos_sensibles": [],
    # Hojas que no se revisan.
    "hojas_ignoradas": [],
    # "0012" y "12" se consideran iguales (Excel suele borrar los ceros a la izquierda).
    "ignorar_ceros_izquierda": True,
    # Umbrales del puntaje por par.
    "umbral_verde": 95,
    "umbral_amarillo": 80,
}

_LISTAS = ("campos_criticos", "campos_excluidos", "campos_sensibles", "hojas_ignoradas")


class ErrorConfig(ValueError):
    pass


def _texto_campo(valor: object) -> str:
    return " ".join(str(valor).split()).upper()


def cargar_config(ruta: Path | None) -> dict:
    """Lee el JSON y lo combina con los valores por defecto. Valida tipos para que un
    archivo mal editado dé un mensaje claro en lugar de un error raro a mitad del proceso."""
    cfg = {k: (list(v) if isinstance(v, list) else v) for k, v in DEFAULTS.items()}
    if ruta is None:
        return _preparar(cfg)
    try:
        datos = json.loads(Path(ruta).read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise ErrorConfig(f"No existe el archivo de configuración: {ruta}") from None
    except json.JSONDecodeError as e:
        raise ErrorConfig(f"{ruta} no es un JSON válido (línea {e.lineno}): {e.msg}") from None
    if not isinstance(datos, dict):
        raise ErrorConfig(f"{ruta} debe contener un objeto JSON {{...}}")

    desconocidas = set(datos) - set(DEFAULTS)
    if desconocidas:
        log.warning("Claves desconocidas en la configuración (se ignoran): %s", sorted(desconocidas))

    for clave in _LISTAS:
        if clave in datos:
            valor = datos[clave]
            if not isinstance(valor, list) or not all(isinstance(x, str) for x in valor):
                raise ErrorConfig(f"'{clave}' debe ser una lista de textos")
            cfg[clave] = valor
    if "marcador_mirror" in datos:
        if not isinstance(datos["marcador_mirror"], str) or not datos["marcador_mirror"].strip():
            raise ErrorConfig("'marcador_mirror' debe ser un texto no vacío")
        cfg["marcador_mirror"] = datos["marcador_mirror"]
    if "ignorar_ceros_izquierda" in datos:
        if not isinstance(datos["ignorar_ceros_izquierda"], bool):
            raise ErrorConfig("'ignorar_ceros_izquierda' debe ser true o false")
        cfg["ignorar_ceros_izquierda"] = datos["ignorar_ceros_izquierda"]
    for clave in ("umbral_verde", "umbral_amarillo"):
        if clave in datos:
            v = datos[clave]
            if isinstance(v, bool) or not isinstance(v, (int, float)) or not 0 <= v <= 100:
                raise ErrorConfig(f"'{clave}' debe ser un número entre 0 y 100")
            cfg[clave] = v
    if cfg["umbral_amarillo"] > cfg["umbral_verde"]:
        raise ErrorConfig("'umbral_amarillo' no puede ser mayor que 'umbral_verde'")
    return _preparar(cfg)


def _preparar(cfg: dict) -> dict:
    """Normaliza los nombres de campos (mayúsculas, espacios) para compararlos igual que los encabezados."""
    for clave in _LISTAS:
        cfg[clave] = {_texto_campo(x) for x in cfg[clave] if str(x).strip()}
    cfg["marcador_mirror"] = _texto_campo(cfg["marcador_mirror"])
    return cfg
