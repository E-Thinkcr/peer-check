import json
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from peercheck.__main__ import main  # noqa: E402
from peercheck.config import ErrorConfig, cargar_config  # noqa: E402
from peercheck.motor import CRITICAL, DISCREPANCY, EMPTY, EXCLUDED, MATCH, buscar_pares, comparar_par  # noqa: E402
from peercheck.normalizar import enmascarar, normalizar  # noqa: E402
from peercheck.procesar import archivos_entrada, procesar_archivo  # noqa: E402
from peercheck.reporte import HOJA_RESUMEN  # noqa: E402

CFG = {
    "campos_criticos": ["CODE", "ACCOUNT"],
    "campos_excluidos": ["EMAIL"],
    "campos_sensibles": ["ACCOUNT"],
    "hojas_ignoradas": ["NOTES"],
}


@pytest.fixture
def cfg(tmp_path):
    ruta = tmp_path / "config.json"
    ruta.write_text(json.dumps(CFG), encoding="utf-8")
    return cargar_config(ruta)


def libro(tmp_path, filas, nombre="prueba.xlsx", titulo="Setup"):
    wb = Workbook()
    ws = wb.active
    ws.title = titulo
    for f in filas:
        ws.append(f)
    ruta = tmp_path / nombre
    wb.save(ruta)
    return ruta


BASICO = [
    ["ID", "NAME", "CODE", "ACCOUNT", "EMAIL", "NOTES"],
    ["A1", "Tienda", "C1", "123456789", "a@x.com", "=HYPERLINK(\"http://x\",\"click\")"],
    ["A2", "Otra", "C2", "555", "b@x.com", None],
    ["MIRROR"],
    ["A1", "TIENDA ", "C1", "123456780", "z@x.com", ""],
    ["A2", "Otra", "C9", "555", "b@x.com", None],
]


# ----------------------------------------------------------------- normalizar
@pytest.mark.parametrize("a,b", [
    (" Tienda\xa0Uno ", "TIENDA UNO"), ("$1,250.00", "1250"), ("7.0", "7"), (7, "7"),
    ("Y", "YES"), ("true", "YES"), (False, "NO"), ("N/A", ""), (None, ""), ("0012", "12"),
])
def test_normalizar(a, b):
    assert normalizar(a) == b


def test_ceros_izquierda_configurables():
    assert normalizar("0012", ignorar_ceros_izquierda=False) == "0012"


def test_textos_no_numericos_no_se_convierten():
    assert normalizar("inf") == "INF"
    assert normalizar("nan") == "NAN"


def test_enmascarar():
    assert enmascarar("123456789") == "•••••6789"
    assert enmascarar("12") == "••"
    assert enmascarar("") == ""


# ----------------------------------------------------------------- motor
def test_pares_y_resultados(tmp_path, cfg):
    ws = load_workbook(libro(tmp_path, BASICO)).active
    avisos = []
    pares = buscar_pares(ws, cfg["marcador_mirror"], avisos)
    assert [(p.agent_id, p.mirror_id) for p in pares] == [("A1", "A1"), ("A2", "A2")]
    r1 = {c.campo: c.resultado for c in comparar_par(ws, pares[0], cfg).campos}
    assert r1 == {"NAME": MATCH, "CODE": MATCH, "ACCOUNT": CRITICAL, "EMAIL": EXCLUDED, "NOTES": DISCREPANCY}
    r2 = comparar_par(ws, pares[1], cfg)
    assert {c.campo: c.resultado for c in r2.campos}["CODE"] == CRITICAL
    assert {c.campo: c.resultado for c in r2.campos}["NOTES"] == EMPTY
    assert r2.puntaje == 66.67
    assert avisos == []


def test_avisa_si_no_cuadran_agents_y_mirrors(tmp_path, cfg):
    filas = BASICO[:3] + [["MIRROR"], BASICO[4]]
    ws = load_workbook(libro(tmp_path, filas)).active
    avisos = []
    assert len(buscar_pares(ws, "MIRROR", avisos)) == 1
    assert "2 agent(s) vs 1 mirror(s)" in avisos[0]


def test_avisa_si_ids_del_par_no_coinciden(tmp_path, cfg):
    filas = BASICO[:4] + [BASICO[5], BASICO[4]]  # mirrors en orden invertido
    ws = load_workbook(libro(tmp_path, filas)).active
    avisos = []
    buscar_pares(ws, "MIRROR", avisos)
    assert any("IDs distintos" in a for a in avisos)


# ----------------------------------------------------------------- archivo completo
def resumen_valores(ruta):
    ws = load_workbook(ruta)[HOJA_RESUMEN]
    return ws, [c.value for fila in ws.iter_rows() for c in fila if c.value is not None]


def test_procesar_archivo(tmp_path, cfg):
    entrada = libro(tmp_path, BASICO)
    salida, respaldo = tmp_path / "OUT", tmp_path / "BAK"
    salida.mkdir(), respaldo.mkdir()
    r = procesar_archivo(entrada, salida, respaldo, cfg)
    assert r["pares"] == 2 and r["criticas"] == 2
    assert len(list(respaldo.iterdir())) == 1
    ws, valores = resumen_valores(r["salida"])
    assert "PEER CHECK SUMMARY" in valores and "PAIR 2" in valores


def test_no_escribe_formulas_en_el_resumen(tmp_path, cfg):
    """Seguridad: un valor como '=HYPERLINK(...)' debe quedar como texto, no como fórmula."""
    entrada = libro(tmp_path, BASICO)
    (tmp_path / "O").mkdir(), (tmp_path / "B").mkdir()
    r = procesar_archivo(entrada, tmp_path / "O", tmp_path / "B", cfg)
    ws, _ = resumen_valores(r["salida"])
    peligrosas = [c for fila in ws.iter_rows() for c in fila
                  if isinstance(c.value, str) and c.value.startswith("=")]
    assert peligrosas, "la prueba debe incluir un valor que empiece con '='"
    assert all(c.data_type == "s" for c in peligrosas)


def test_enmascara_campos_sensibles(tmp_path, cfg):
    entrada = libro(tmp_path, BASICO)
    (tmp_path / "O").mkdir(), (tmp_path / "B").mkdir()
    r = procesar_archivo(entrada, tmp_path / "O", tmp_path / "B", cfg)
    _, valores = resumen_valores(r["salida"])
    assert "123456789" not in valores and "•••••6789" in valores


def test_hojas_ignoradas_y_sin_mirror(tmp_path, cfg):
    wb = Workbook()
    wb.active.title = "Notes"
    wb.create_sheet("Vacia").append(["ID", "X"])
    ruta = tmp_path / "x.xlsx"
    wb.save(ruta)
    (tmp_path / "O").mkdir(), (tmp_path / "B").mkdir()
    r = procesar_archivo(ruta, tmp_path / "O", tmp_path / "B", cfg)
    assert r["pares"] == 0 and r["avisos"] == 2


def test_celdas_combinadas_no_rompen(tmp_path, cfg):
    ruta = libro(tmp_path, BASICO)
    wb = load_workbook(ruta)
    wb.active.merge_cells("B2:C2")
    wb.save(ruta)
    (tmp_path / "O").mkdir(), (tmp_path / "B").mkdir()
    procesar_archivo(ruta, tmp_path / "O", tmp_path / "B", cfg)


def test_ignora_temporales_de_excel(tmp_path):
    (tmp_path / "~$abierto.xlsx").write_bytes(b"x")
    (tmp_path / "bueno.XLSX").write_bytes(b"x")
    (tmp_path / "nota.txt").write_text("x")
    assert [p.name for p in archivos_entrada(tmp_path)] == ["bueno.XLSX"]


def test_archivo_danado_da_error_claro(tmp_path, cfg):
    (tmp_path / "INPUT").mkdir()
    (tmp_path / "INPUT" / "roto.xlsx").write_bytes(b"no soy un excel")
    (tmp_path / "config.json").write_text(json.dumps(CFG), encoding="utf-8")
    assert main(["--carpeta", str(tmp_path)]) == 1


def test_no_pisa_resultados_previos(tmp_path, cfg):
    entrada = libro(tmp_path, BASICO)
    (tmp_path / "O").mkdir(), (tmp_path / "B").mkdir()
    a = procesar_archivo(entrada, tmp_path / "O", tmp_path / "B", cfg)["salida"]
    b = procesar_archivo(entrada, tmp_path / "O", tmp_path / "B", cfg)["salida"]
    assert a != b and a.exists() and b.exists()


# ----------------------------------------------------------------- config
@pytest.mark.parametrize("contenido", [
    "no es json", "[]", '{"campos_criticos": "CODE"}', '{"umbral_verde": 150}',
    '{"umbral_verde": 50, "umbral_amarillo": 90}', '{"marcador_mirror": ""}',
])
def test_config_invalida(tmp_path, contenido):
    ruta = tmp_path / "c.json"
    ruta.write_text(contenido, encoding="utf-8")
    with pytest.raises(ErrorConfig):
        cargar_config(ruta)


def test_config_de_ejemplo_es_valida():
    cfg = cargar_config(Path(__file__).resolve().parents[1] / "config.example.json")
    assert "BANK ACCOUNT" in cfg["campos_criticos"]


def test_ejemplo_completo(tmp_path):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "ejemplo"))
    import generar_ejemplo

    generar_ejemplo.CARPETA = tmp_path / "INPUT"
    generar_ejemplo.main()
    cfg_ruta = Path(__file__).resolve().parents[1] / "config.example.json"
    assert main(["--carpeta", str(tmp_path), "--config", str(cfg_ruta)]) == 0
    salida = next((tmp_path / "OUTPUT").iterdir())
    _, valores = resumen_valores(salida)
    assert "CRITICAL" in valores and "Instructions: IGNORED" in valores
