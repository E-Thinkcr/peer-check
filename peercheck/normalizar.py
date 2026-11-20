"""Limpieza de valores antes de comparar.

Quita espacios invisibles, símbolos de formato ($ % ,), unifica sí/no y números
(1, 1.0 y "1" son iguales), para que solo cuenten las diferencias reales.
"""
from __future__ import annotations

import math
import re
from datetime import date, datetime, time

_INVISIBLES = str.maketrans({c: " " for c in "\xa0​‌‍﻿\t\n\r"})
_VACIOS = {"", "NULL", "NONE", "N/A", "NA", "-"}
_SI = {"TRUE", "YES", "Y", "SI", "SÍ"}
_NO = {"FALSE", "NO", "N"}
_NUMERO = re.compile(r"[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?")


def normalizar(valor: object, ignorar_ceros_izquierda: bool = True) -> str:
    if valor is None:
        return ""
    if isinstance(valor, bool):
        return "YES" if valor else "NO"
    if isinstance(valor, datetime):
        return valor.strftime("%Y-%m-%d") if valor.time() == time(0) else valor.isoformat(" ")
    if isinstance(valor, date):
        return valor.isoformat()

    texto = str(valor).translate(_INVISIBLES)
    for simbolo in ("$", "%", ","):
        texto = texto.replace(simbolo, "")
    texto = re.sub(r"\s+", " ", texto).strip().upper()

    if texto in _VACIOS:
        return ""
    if texto in _SI:
        return "YES"
    if texto in _NO:
        return "NO"

    # Solo se convierten a número los textos que realmente son números
    # (antes, "INF" o "NAN" se convertían en número por accidente).
    if _NUMERO.fullmatch(texto):
        if not ignorar_ceros_izquierda and re.fullmatch(r"0\d+", texto):
            return texto  # códigos como "0012" se respetan tal cual
        numero = float(texto)
        if math.isfinite(numero):
            if numero.is_integer():
                return str(int(numero))
            return repr(numero)
    return texto


def enmascarar(valor: str, visibles: int = 4) -> str:
    """'123456789' -> '•••••6789'. Para que el resumen no exponga cuentas completas."""
    if not valor:
        return valor
    if len(valor) <= visibles:
        return "•" * len(valor)
    return "•" * (len(valor) - visibles) + valor[-visibles:]
