"""Command line entry point: ``python -m oi_tracker KC``."""

from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path

import pandas as pd

from .barchart import BarchartClient
from .data import append_history, snapshot
from .report import build_report


def parse_args(argv=None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="oi_tracker",
        description="Snapshot futures-options open interest from Barchart and "
                    "build an Excel report tracking it over time.",
    )
    p.add_argument("root", help="Futures root symbol, e.g. KC (coffee), CL (crude), ZC (corn)")
    p.add_argument("-c", "--contracts", type=int, default=4,
                   help="number of nearest futures contracts to track (default: 4)")
    p.add_argument("-s", "--strikes", type=int, default=10,
                   help="strikes shown on each side of the money (default: 10)")
    p.add_argument("-o", "--out", type=Path, default=Path("data"),
                   help="output folder for the CSV history and Excel report (default: data)")
    p.add_argument("--show-browser", action="store_true",
                   help="run Chrome visibly instead of headless (useful for debugging)")
    return p.parse_args(argv)


def main(argv=None) -> None:
    args = parse_args(argv)
    root = args.root.upper()
    today = date.today()

    snapshots = []
    with BarchartClient(headless=not args.show_browser) as bc:
        bc.open(root)
        contracts = bc.futures_contracts(root)[: args.contracts]
        if not contracts:
            raise SystemExit(f"No futures contracts found for root {root!r}")
        for c in contracts:
            chain = bc.option_chain(c.symbol)
            snap = snapshot(chain, c.symbol, c.last_price, today)
            print(f"{c.symbol}: price {c.last_price}, {len(snap)} option strikes, "
                  f"total OI {snap['open_interest'].sum():,}")
            snapshots.append(snap)

    history_path = args.out / f"{root}_history.csv"
    report_path = args.out / f"{root}_open_interest.xlsx"
    history = append_history(history_path, pd.concat(snapshots, ignore_index=True))
    build_report(history, report_path, args.strikes)
    print(f"History: {history_path} ({history['date'].nunique()} day(s))")
    print(f"Report:  {report_path}")
