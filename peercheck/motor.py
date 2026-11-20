"""Motor de comparación: encuentra los pares Agent ↔ Mirror de una hoja y compara campo por campo.

Formato esperado de cada hoja:

    Fila 1            encabezados (columna A = ID)
    Filas 2..k-1      registros capturados (Agent 1..N)
    Fila k            el texto "MIRROR" en la columna A
    Filas k+1..       registros de control (Mirror 1..N), en el mismo orden

El Agent N se compara con el Mirror N.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from openpyxl.worksheet.worksheet import Worksheet

from .normalizar import normalizar

MATCH, DISCREPANCY, CRITICAL, EXCLUDED, EMPTY = "MATCH", "DISCREPANCY", "CRITICAL", "EXCLUDED", "EMPTY"


@dataclass
class Par:
    agent_fila: int
    mirror_fila: int
    agent_id: str
    mirror_id: str


@dataclass
class ResultadoCampo:
    columna: int
    campo: str
    agent: str
    mirror: str
    resultado: str


@dataclass
class ResultadoPar:
    hoja: str
    par: Par
    campos: list[ResultadoCampo] = field(default_factory=list)

    def contar(self, *tipos: str) -> int:
        return sum(1 for c in self.campos if c.resultado in tipos)

    @property
    def coincidencias(self) -> int:
        return self.contar(MATCH)

    @property
    def diferencias(self) -> int:
        """Incluye las críticas (igual que la versión original)."""
        return self.contar(DISCREPANCY, CRITICAL)

    @property
    def criticas(self) -> int:
        return self.contar(CRITICAL)

    @property
    def puntaje(self) -> float:
        total = self.coincidencias + self.diferencias
        return round(self.coincidencias / total * 100, 2) if total else 0.0


def _texto_id(valor: object) -> str:
    return "" if valor is None else str(valor).strip()


def buscar_pares(ws: Worksheet, marcador: str, avisos: list[str]) -> list[Par]:
    mirror = None
    for fila in range(1, ws.max_row + 1):
        if normalizar(ws.cell(fila, 1).value) == marcador:
            mirror = fila
            break
    if mirror is None:
        return []

    agents = [f for f in range(2, mirror) if normalizar(ws.cell(f, 1).value)]
    mirrors = [f for f in range(mirror + 1, ws.max_row + 1) if normalizar(ws.cell(f, 1).value)]

    if len(agents) != len(mirrors):
        # Antes se ignoraba en silencio: ahora queda como advertencia en el reporte.
        avisos.append(f"{ws.title}: {len(agents)} agent(s) vs {len(mirrors)} mirror(s); "
                      f"solo se comparan los primeros {min(len(agents), len(mirrors))}")

    pares = []
    for a, m in zip(agents, mirrors):
        par = Par(a, m, _texto_id(ws.cell(a, 1).value), _texto_id(ws.cell(m, 1).value))
        if normalizar(par.agent_id) != normalizar(par.mirror_id):
            avisos.append(f"{ws.title}: el par agent {par.agent_id} / mirror {par.mirror_id} "
                          "tiene IDs distintos; revisa el orden de las filas")
        pares.append(par)
    return pares


def comparar_par(ws: Worksheet, par: Par, cfg: dict) -> ResultadoPar:
    res = ResultadoPar(ws.title, par)
    ceros = cfg["ignorar_ceros_izquierda"]
    for col in range(2, ws.max_column + 1):
        campo = normalizar(ws.cell(1, col).value)
        if not campo:
            continue
        agent = normalizar(ws.cell(par.agent_fila, col).value, ceros)
        mirror = normalizar(ws.cell(par.mirror_fila, col).value, ceros)

        if agent == "" and mirror == "":
            resultado = EMPTY
        elif campo in cfg["campos_excluidos"]:
            resultado = EXCLUDED
        elif agent == mirror:
            resultado = MATCH
        elif campo in cfg["campos_criticos"]:
            resultado = CRITICAL
        else:
            resultado = DISCREPANCY
        res.campos.append(ResultadoCampo(col, campo, agent, mirror, resultado))
    return res
