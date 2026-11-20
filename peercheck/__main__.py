"""Uso:
    python -m peercheck                      # usa la carpeta actual
    python -m peercheck --carpeta "C:\\PeerCheck" --config config.json
"""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from . import APP_NOMBRE, __version__
from .config import ErrorConfig, cargar_config
from .procesar import procesar_carpeta


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="peercheck", description=f"{APP_NOMBRE} {__version__}")
    p.add_argument("--carpeta", type=Path, default=Path.cwd(),
                   help="Carpeta que contiene INPUT, OUTPUT y BACKUP (por defecto, la actual)")
    p.add_argument("--config", type=Path, default=None,
                   help="Archivo JSON con los campos críticos/excluidos (por defecto, <carpeta>/config.json)")
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = p.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(message)s")
    log = logging.getLogger("peercheck")

    ruta_cfg = args.config or (args.carpeta / "config.json")
    if args.config is None and not ruta_cfg.exists():
        log.warning("No se encontró %s: se usan los valores por defecto (sin campos críticos ni "
                    "excluidos). Copia config.example.json como config.json y ajústalo.", ruta_cfg)
        ruta_cfg = None
    try:
        cfg = cargar_config(ruta_cfg)
    except ErrorConfig as e:
        log.error("Configuración inválida: %s", e)
        return 2

    errores = procesar_carpeta(args.carpeta, cfg)
    return 1 if errores else 0


if __name__ == "__main__":
    sys.exit(main())
