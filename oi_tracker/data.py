"""Turn raw option chains into tidy snapshots and keep a CSV history."""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd

COLUMNS = [
    "date", "contract", "underlying_price", "option_type", "strike", "open_interest",
]


def snapshot(
    chain: list[dict], contract: str, underlying_price: float, as_of: date,
) -> pd.DataFrame:
    """Convert one raw option chain into rows of the history table."""
    df = pd.DataFrame(chain)
    if df.empty:
        return pd.DataFrame(columns=COLUMNS)
    df = df.rename(columns={
        "strikePrice": "strike", "optionType": "option_type",
        "openInterest": "open_interest",
    })
    df = df[df["option_type"].isin(["Call", "Put"])]
    df["open_interest"] = pd.to_numeric(df["open_interest"]).fillna(0).astype(int)
    df["strike"] = pd.to_numeric(df["strike"])
    df["date"] = pd.Timestamp(as_of)
    df["contract"] = contract
    df["underlying_price"] = underlying_price
    return df[COLUMNS].reset_index(drop=True)


def strikes_around(strikes, price: float, n: int) -> list[float]:
    """The ``n`` strikes just below ``price`` and the ``n`` just above it."""
    unique = sorted(set(strikes))
    below = [s for s in unique if s <= price][-n:]
    above = [s for s in unique if s > price][:n]
    return below + above


MONTH_CODES = "FGHJKMNQUVXZ"  # Jan..Dec futures month letters


def contract_sort_key(symbol: str) -> tuple[int, int]:
    """Sort key putting e.g. KCZ26 before KCH27 (by expiry, not alphabetically)."""
    month, year = symbol[-3], int(symbol[-2:])
    return year, MONTH_CODES.index(month)


def load_history(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=COLUMNS)
    return pd.read_csv(path, parse_dates=["date"])


def append_history(path: Path, new: pd.DataFrame) -> pd.DataFrame:
    """Add a snapshot to the history, replacing any earlier run on the same day."""
    history = load_history(path)
    frames = [f for f in (history, new) if not f.empty]
    combined = pd.concat(frames, ignore_index=True) if frames else new
    combined = combined.drop_duplicates(
        subset=["date", "contract", "option_type", "strike"], keep="last",
    ).sort_values(["date", "contract", "option_type", "strike"])
    path.parent.mkdir(parents=True, exist_ok=True)
    combined.to_csv(path, index=False, date_format="%Y-%m-%d")
    return combined.reset_index(drop=True)
