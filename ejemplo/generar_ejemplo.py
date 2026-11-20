"""Crea un Excel de ejemplo con datos 100 % ficticios para probar Peer Check:
    python ejemplo/generar_ejemplo.py
Genera ejemplo/INPUT/ejemplo_setup.xlsx. Después corre:
    python -m peercheck --carpeta ejemplo --config config.example.json
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

CARPETA = Path(__file__).resolve().parent / "INPUT"

UBICACIONES = ["ID", "NAME", "ACCOUNT TYPE", "NETWORK", "SUBNETWORK", "BRANCH", "PAYMENT METHOD",
               "STATUS", "EMAIL", "PHONE", "START DATE", "OPEN SUNDAY"]
AGENTS = [
    ["LOC-1001", "Pulpería La Esquina", "RETAIL", "NORTE", "N-01", "0012", "CASH", "ACTIVE",
     "esquina@example.com", "2222-0001", "2026-11-02", "YES"],
    ["LOC-1002", "Farmacia Central", "RETAIL", "CENTRO", "C-04", "0045", "CASH", "ACTIVE",
     "central@example.com", "2222-0002", "2026-11-03", "NO"],
    ["LOC-1003", "Súper El Roble", "CHAIN", "OESTE", "O-02", "0101", "BANK DEPOSIT", "PENDING",
     "roble@example.com", "2222-0003", "2026-11-04", "Y"],
]
MIRRORS = [
    ["LOC-1001", "Pulpería La Esquina ", "RETAIL", "NORTE", "N-01", "12", "CASH", "ACTIVE",
     "esquina@example.com", "2222-0001", "2026-11-05", "TRUE"],          # todo coincide
    ["LOC-1002", "Farmacia Central", "RETAIL", "CENTRO", "C-05", "0045", "CASH", "Active",
     "otra@example.com", "2222-0002", "2026-11-03", "N"],                # SUBNETWORK crítico
    ["LOC-1003", "Super El Roble", "CHAIN", "OESTE", "O-02", "0101", "CASH", "PENDING",
     "roble@example.com", "2222-0003", "2026-11-04", "YES"],             # nombre + método de pago
]

BANCO = ["ID", "BANK CODE", "BANK ACCOUNT", "ROUTING NUMBER", "SETTLEMENT TYPE", "COMMISSION PLAN"]
BANCO_A = [["LOC-1001", "BK-07", "100200300400", "021000089", "DAILY", "PLAN A"],
           ["LOC-1003", "BK-11", "555000111222", "011000015", "WEEKLY", "PLAN B"]]
BANCO_M = [["LOC-1001", "BK-07", "100200300400", "021000089", "DAILY", "PLAN A"],
           ["LOC-1003", "BK-11", "555000111223", "011000015", "WEEKLY", "Plan B"]]  # cuenta crítica


def hoja(wb, titulo, encabezados, agents, mirrors):
    ws = wb.create_sheet(titulo)
    ws.append(encabezados)
    for c in ws[1]:
        c.font = Font(bold=True)
    for fila in agents:
        ws.append(fila)
    ws.append(["MIRROR"])
    for fila in mirrors:
        ws.append(fila)


def main() -> Path:
    wb = Workbook()
    wb.remove(wb.active)
    hoja(wb, "Location Setup", UBICACIONES, AGENTS, MIRRORS)
    hoja(wb, "Banking", BANCO, BANCO_A, BANCO_M)
    ws = wb.create_sheet("Instructions")
    ws.append(["Hoja de instrucciones: Peer Check la ignora (ver hojas_ignoradas)."])
    CARPETA.mkdir(parents=True, exist_ok=True)
    destino = CARPETA / "ejemplo_setup.xlsx"
    wb.save(destino)
    print("Creado:", destino)
    return destino


if __name__ == "__main__":
    main()
