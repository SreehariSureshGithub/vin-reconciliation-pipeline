"""Pipeline configuration: defaults that can be overridden from a YAML file."""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path

import yaml

DEFAULTS: dict = {
    # Input files (glob patterns, relative to the config file's folder).
    "inputs": {
        "stock": "data/stock*.xlsx",
        "general_ledger": "data/*General Ledger*.xlsx",
        "purchases_detail": "data/*Purchases by Supplier Detail*.xlsx",
        "sales_detail": "data/*Sales by Customer Detail*.xlsx",
        "supplier_balance": "data/*Supplier Balance Detail*.xlsx",
        "customer_balance": "data/*Customer Balance Detail*.xlsx",
        "bank": "data/bank*.xlsx",
    },
    "output": "output/reconciliation.xlsx",
    # Stock-sheet column names -> internal field names.
    "stock_columns": {
        "vin": "VIN",
        "cycle": "Deal Cycle #",
        "deal_type": "Deal Type(s)",
        "purchase_date": "Date of Purchase",
        "purchase_ref": "Purchase Invoice No",
        "supplier": "Purchased From (Seller/Consignor)",
        "purchase_value": "Purchase Cost",
        "sale_date": "Sold Date",
        "sale_ref": "Sales Invoice No",
        "buyer": "Sold To (Customer)",
        "sale_value": "Sales Price",
        "margin": "Margin",
    },
    "stock_sheet": 0,
    # Bank statement columns; the header row is located automatically.
    "bank_columns": {"date": "DATE", "description": "Description", "narration": "NARATION",
                     "income": "INCOME", "expense": "EXPENSE"},
    "bank_sheet": 0,
    # Words in deal_type that mark a deal as consignment / cancelled.
    "consignment_keywords": ["CONSIGN"],
    "cancelled_keywords": ["CANCEL"],
    # Bank lines containing any of these are internal transfers, not collections.
    "internal_transfer_keywords": ["OWN ACCOUNT", "INTERNAL TRANSFER"],
    # QuickBooks transaction types that represent a vehicle purchase / sale.
    "purchase_types": ["Bill"],
    "sale_types": ["Invoice"],
    # General Ledger accounts that carry the full bill / invoice amount.
    "ap_account": "Accounts Payable",
    "ar_account": "Accounts Receivable",
    "amount_tolerance": 1.0,          # currency units; smaller differences count as equal
    "consignment_vat_rate": 0.05,     # consignment invoice = margin * (1 + rate)
    "fallback_supplier": "Cash Supplier",  # used only after all evidence is exhausted
    "fallback_buyer": "Cash Customer",
}


def _merge(base: dict, override: dict) -> dict:
    out = deepcopy(base)
    for key, value in (override or {}).items():
        out[key] = _merge(out[key], value) if isinstance(value, dict) and isinstance(out.get(key), dict) else value
    return out


def load_config(path: str | Path | None = None) -> dict:
    """Return DEFAULTS merged with the YAML file at `path`; paths resolve relative to it."""
    root = Path(path).resolve().parent if path else Path.cwd()
    cfg = _merge(DEFAULTS, yaml.safe_load(Path(path).read_text(encoding="utf-8")) if path else {})
    cfg["root"] = root
    return cfg
