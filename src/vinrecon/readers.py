"""Readers for the stock sheet, QuickBooks report exports and bank statements."""

from __future__ import annotations

import glob
import re
import warnings
from pathlib import Path

import pandas as pd

VIN_RE = re.compile(r"\b[A-HJ-NPR-Z0-9]{17}\b")
warnings.filterwarnings("ignore", message="Parsing dates in", category=UserWarning)  # mixed d/m/Y exports


# ---------------------------------------------------------------- value cleaning
def clean(v) -> str:
    """Whitespace-normalised text; '' for missing values."""
    return "" if v is None or (not isinstance(v, str) and pd.isna(v)) else " ".join(str(v).split())


def norm_ref(v) -> str:
    """Match key for invoice/bill numbers: upper-case, no whitespace, no trailing punctuation."""
    return re.sub(r"\s+", "", clean(v)).upper().rstrip(".,;")


def norm_vin(v) -> str:
    return re.sub(r"[^A-Z0-9]", "", clean(v).upper())


def norm_name(v) -> str:
    return clean(v).casefold()


def to_num(v) -> float | None:
    if isinstance(v, (int, float)) and not pd.isna(v):
        return float(v)
    s = clean(v).replace(",", "")
    neg = s.startswith("(") and s.endswith(")")
    try:
        n = float(s.strip("()"))
    except ValueError:
        return None
    return -n if neg else n


def to_date(v):
    if isinstance(v, pd.Timestamp) or hasattr(v, "year"):
        return pd.Timestamp(v)
    s = clean(v)
    return pd.to_datetime(s, dayfirst=True, errors="coerce") if s else pd.NaT


def files(cfg: dict, key: str) -> list[Path]:
    """Files for an input key; patterns may be relative to the config file or absolute, one or a list."""
    patterns = cfg["inputs"].get(key) or []
    found = set()
    for p in [patterns] if isinstance(patterns, str) else patterns:
        found.update(Path(f) for f in glob.glob(str(cfg["root"] / p)))  # absolute patterns override root
    return sorted(found)


def _find_header(raw: pd.DataFrame, required: str) -> int:
    for i, row in raw.head(30).iterrows():
        if required.casefold() in {clean(v).casefold() for v in row}:
            return i
    raise ValueError(f"header row containing '{required}' not found")


def _read_table(path: Path, sheet, required: str) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object)
    h = _find_header(raw, required)
    df = raw.iloc[h + 1:].copy()
    df.columns = [clean(c) or f"col{i}" for i, c in enumerate(raw.iloc[h])]
    return df.dropna(how="all")


# ---------------------------------------------------------------- stock sheet
def read_stock(cfg: dict) -> pd.DataFrame:
    """One row per deal cycle with internal column names (see config 'stock_columns')."""
    cols = cfg["stock_columns"]
    frames = []
    for path in files(cfg, "stock"):
        df = _read_table(path, cfg["stock_sheet"], cols["vin"])
        df = df.rename(columns={v: k for k, v in cols.items()}).reindex(columns=list(cols))
        frames.append(df.assign(source=[f"{path.name} row {i + 1}" for i in df.index]))
    if not frames:
        raise FileNotFoundError(f"no stock file matches {cfg['inputs']['stock']}")
    s = pd.concat(frames, ignore_index=True)
    s["vin_raw"] = s["vin"].map(clean)
    s = s[s["vin_raw"] != ""].copy()
    s["vin"] = s["vin_raw"].map(norm_vin)
    for c in ("purchase_value", "sale_value", "margin"):
        s[c] = s[c].map(to_num)
    for c in ("purchase_date", "sale_date"):
        s[c] = s[c].map(to_date)
    for c in ("supplier", "buyer", "deal_type", "purchase_ref", "sale_ref"):
        s[c] = s[c].map(clean)
    s["cycle"] = pd.to_numeric(s["cycle"], errors="coerce")
    missing = s["cycle"].isna()
    if missing.any():  # derive cycles from purchase order where the sheet has none
        order = s.sort_values(["vin", "purchase_date"], na_position="last").groupby("vin").cumcount() + 1
        s.loc[missing, "cycle"] = order[missing]
    return s.astype({"cycle": int}).sort_values(["vin", "cycle"]).reset_index(drop=True)


# ---------------------------------------------------------------- QuickBooks reports
_FIELD = {  # QuickBooks column name -> internal field
    "transaction date": "date", "date": "date", "transaction type": "type", "number": "ref",
    "name": "party", "supplier display name": "party", "supplier full name": "party",
    "customer display name": "party", "customer full name": "party",
    "distribution account": "account", "description": "description",
    "product/service full name": "product", "amount": "amount", "open balance": "open_balance",
}


def read_qb_report(path: Path, kind: str) -> pd.DataFrame:
    """Parse any QuickBooks 'detail' export into one row per transaction line.

    Layout handled: title rows ('As of <date>' optional), a header row, group heading
    rows (an account or a party name in column A), detail rows, 'Total for ...' rows.
    The innermost heading is kept as `group`; it is the party in party-grouped reports.
    """
    raw = pd.read_excel(path, header=None, dtype=object)
    title = " ".join(clean(v) for v in raw.iloc[:4, 0])
    as_of = re.search(r"As of (.+)$", title)
    h = next(i for i, r in raw.head(30).iterrows() if not clean(r.iloc[0]) and r.iloc[1:].notna().sum() >= 3)
    names = [_FIELD.get(clean(c).casefold(), clean(c).casefold()) for c in raw.iloc[h]]
    stack, rows = [], []
    for r in raw.iloc[h + 1:].itertuples(index=False):
        first = clean(r[0])
        if first:
            if first.upper() == "TOTAL":
                break
            if first.startswith("Total for"):
                stack = stack[:-1]
            elif all(pd.isna(v) or clean(v) == "" for v in r[1:]):
                stack.append(first)
            continue
        if any(clean(v) for v in r[1:]):
            rows.append({**{n: v for n, v in zip(names, r) if n}, "group": stack[-1] if stack else ""})
    df = pd.DataFrame(rows).reindex(columns=["date", "type", "ref", "party", "account", "description",
                                             "product", "amount", "open_balance", "group"])
    df["party"] = [clean(p) or (g if kind != "general_ledger" else "") for p, g in zip(df["party"], df["group"])]
    df["account"] = [clean(a) or (g if kind == "general_ledger" else "") for a, g in zip(df["account"], df["group"])]
    df["text"] = (df["description"].map(clean) + " " + df["product"].map(clean)).str.upper()
    return df.assign(
        kind=kind, file=path.name,
        as_of=to_date(as_of.group(1)) if as_of else pd.NaT,
        date=df["date"].map(to_date), type=df["type"].map(clean),
        ref=df["ref"].map(norm_ref), ref_raw=df["ref"].map(clean),
        amount=df["amount"].map(to_num), open_balance=df["open_balance"].map(to_num),
    ).drop(columns=["description", "product", "group"])


def read_qb(cfg: dict) -> pd.DataFrame:
    kinds = ("general_ledger", "purchases_detail", "sales_detail", "supplier_balance", "customer_balance")
    frames = [read_qb_report(p, k) for k in kinds for p in files(cfg, k)]
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(
        columns=["date", "type", "ref", "party", "account", "amount", "open_balance", "text",
                 "kind", "file", "as_of", "ref_raw"])


# ---------------------------------------------------------------- bank statement
def read_bank(cfg: dict) -> pd.DataFrame:
    cols = cfg["bank_columns"]
    frames = []
    for path in files(cfg, "bank"):
        df = _read_table(path, cfg["bank_sheet"], cols["date"]).rename(columns={v: k for k, v in cols.items()})
        frames.append(df.reindex(columns=list(cols)).assign(file=path.name))
    if not frames:
        return pd.DataFrame(columns=[*cols, "file", "text"])
    b = pd.concat(frames, ignore_index=True)
    b["date"] = b["date"].map(to_date)
    b["income"], b["expense"] = b["income"].map(to_num), b["expense"].map(to_num)
    b["text"] = (b["description"].map(clean) + " " + b["narration"].map(clean)).str.upper()
    return b
