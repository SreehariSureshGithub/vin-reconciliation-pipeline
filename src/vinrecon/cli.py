"""Command line: `vinrecon run` and `vinrecon sample`."""

from __future__ import annotations

import argparse
import logging
import sys

from . import __version__
from .config import load_config
from .core import run
from .report import write_report
from .sample import generate


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="vinrecon", description="Reconcile vehicle stock against QuickBooks and bank data.")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)
    p_run = sub.add_parser("run", help="run the reconciliation and write the Excel report")
    p_run.add_argument("-c", "--config", default="config.yaml", help="YAML config (default: config.yaml)")
    p_run.add_argument("-o", "--output", help="output .xlsx (overrides the config)")
    p_sample = sub.add_parser("sample", help="write a synthetic demo dataset and config")
    p_sample.add_argument("directory", nargs="?", default="sample")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    log = logging.getLogger("vinrecon")

    if args.command == "sample":
        log.info("Sample data written. Next: vinrecon run -c %s", generate(args.directory))
        return 0
    try:
        cfg = load_config(args.config)
        result = run(cfg)
    except (FileNotFoundError, ValueError) as exc:
        log.error("error: %s", exc)
        return 1
    out = write_report(result, args.output or cfg["root"] / cfg["output"])
    recon = result["reconciliation"]
    log.info("%d deal cycles | %d discrepancies | %d need review -> %s",
             len(recon), len(result["discrepancies"]), recon["needs_review"].sum(), out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
