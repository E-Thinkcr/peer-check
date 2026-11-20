"""Escribe la hoja "Peer Check Summary" y colorea las celdas revisadas en las hojas originales."""
from __future__ import annotations

from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from .motor import CRITICAL, DISCREPANCY, EMPTY, EXCLUDED, MATCH, ResultadoPar
from .normalizar import enmascarar

HOJA_RESUMEN = "Peer Check Summary"

VERDE = PatternFill("solid", fgColor="92D050")
ROJO = PatternFill("solid", fgColor="FFC7CE")
AMARILLO = PatternFill("solid", fgColor="FFF2CC")
GRIS = PatternFill("solid", fgColor="D9D9D9")
AZUL = PatternFill("solid", fgColor="1F4E78")
AZUL_CLARO = PatternFill("solid", fgColor="D9EAF7")
DELGADO = Border(*(Side(style="thin"),) * 4)
BLANCO_NEGRITA = Font(bold=True, color="FFFFFF")

COLOR = {MATCH: VERDE, DISCREPANCY: ROJO, CRITICAL: ROJO, EXCLUDED: GRIS, EMPTY: GRIS}
_PELIGROSO = ("=", "+", "-", "@", "\t", "\r")


def escribir(ws: Worksheet, fila: int, col: int, valor, relleno=None, borde=False, fuente=None):
    """Escribe un valor como TEXTO. Si una celda trae algo como '=HYPERLINK(...)', Excel
    lo ejecutaría como fórmula; aquí se guarda siempre como texto plano."""
    celda = ws.cell(fila, col)
    celda.value = valor
    if isinstance(valor, str) and valor.startswith(_PELIGROSO):
        celda.data_type = "s"
        celda.quotePrefix = True
    if relleno is not None:
        celda.fill = relleno
    if borde:
        celda.border = DELGADO
    if fuente is not None:
        celda.font = fuente
    return celda


def colorear_origen(ws: Worksheet, res: ResultadoPar) -> None:
    for c in res.campos:
        for fila in (res.par.agent_fila, res.par.mirror_fila):
            celda = ws.cell(fila, c.columna)
            if not isinstance(celda, MergedCell):  # las celdas combinadas no aceptan formato
                celda.fill = COLOR[c.resultado]


def crear_resumen(wb: Workbook, resultados: list[ResultadoPar], avisos: list[str], cfg: dict) -> Worksheet:
    if HOJA_RESUMEN in wb.sheetnames:
        del wb[HOJA_RESUMEN]
    ws = wb.create_sheet(HOJA_RESUMEN, 0)

    ws.merge_cells("A1:F1")
    ws["A1"] = "PEER CHECK SUMMARY"
    ws["A1"].fill = AZUL
    ws["A1"].font = Font(bold=True, size=16, color="FFFFFF")
    ws["A1"].alignment = Alignment(horizontal="center")

    coincidencias = sum(r.coincidencias for r in resultados)
    diferencias = sum(r.diferencias for r in resultados)
    criticas = sum(r.criticas for r in resultados)
    total = coincidencias + diferencias
    general = round(coincidencias / total * 100, 2) if total else 0.0

    tablero = [("PAIRS REVIEWED", len(resultados)), ("MATCHES", coincidencias),
               ("DISCREPANCIES", diferencias), ("CRITICAL", criticas),
               ("OVERALL SCORE", general / 100), ("WARNINGS", len(avisos))]
    escribir(ws, 3, 1, "EXECUTIVE DASHBOARD", AZUL, fuente=BLANCO_NEGRITA)
    for i, (nombre, valor) in enumerate(tablero, start=4):
        escribir(ws, i, 1, nombre, AZUL_CLARO, borde=True)
        celda = escribir(ws, i, 2, valor, borde=True)
        if nombre == "OVERALL SCORE":
            celda.number_format = "0.00%"
            celda.fill = _relleno_puntaje(general, cfg)
        if nombre == "CRITICAL" and valor:
            celda.fill = ROJO

    fila = 12
    for n, res in enumerate(resultados, start=1):
        escribir(ws, fila, 1, f"PAIR {n}", AZUL, fuente=BLANCO_NEGRITA)
        for etiqueta, valor in (("AGENT", res.par.agent_id), ("MIRROR", res.par.mirror_id), ("SHEET", res.hoja)):
            fila += 1
            escribir(ws, fila, 1, etiqueta)
            escribir(ws, fila, 2, valor)
        fila += 2
        for col, texto in enumerate(("FIELD", "AGENT VALUE", "MIRROR VALUE", "RESULT"), start=1):
            escribir(ws, fila, col, texto, AZUL_CLARO, borde=True)
        fila += 1
        for c in res.campos:
            sensible = c.campo in cfg["campos_sensibles"]
            agent = enmascarar(c.agent) if sensible else c.agent
            mirror = enmascarar(c.mirror) if sensible else c.mirror
            for col, valor in enumerate((c.campo, agent, mirror, c.resultado), start=1):
                escribir(ws, fila, col, valor, borde=True)
            ws.cell(fila, 4).fill = COLOR[c.resultado]
            fila += 1
        escribir(ws, fila, 1, "PAIR SCORE")
        celda = escribir(ws, fila, 2, res.puntaje / 100, _relleno_puntaje(res.puntaje, cfg))
        celda.number_format = "0.00%"
        fila += 3

    escribir(ws, fila, 1, "WARNINGS", AZUL, fuente=BLANCO_NEGRITA)
    for aviso in avisos or ["(none)"]:
        fila += 1
        escribir(ws, fila, 1, aviso, AMARILLO if avisos else None)

    _ajustar_anchos(ws)
    ws.freeze_panes = "A2"
    return ws


def _relleno_puntaje(puntaje: float, cfg: dict) -> PatternFill:
    if puntaje >= cfg["umbral_verde"]:
        return VERDE
    if puntaje >= cfg["umbral_amarillo"]:
        return AMARILLO
    return ROJO


def _ajustar_anchos(ws: Worksheet) -> None:
    anchos: dict[int, int] = {}
    for fila in ws.iter_rows(min_row=2):
        for celda in fila:
            if celda.value is not None and not isinstance(celda, MergedCell):
                anchos[celda.column] = max(anchos.get(celda.column, 0), len(str(celda.value)))
    for col, ancho in anchos.items():
        ws.column_dimensions[get_column_letter(col)].width = min(ancho + 3, 60)
