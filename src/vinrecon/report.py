"""Excel report writer."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl.utils import get_column_letter

RECON_COLUMNS = {
    "vin_raw": "VIN", "cycle": "Deal Cycle", "status": "Status", "deal_type": "Deal Type",
    "purchase_date": "Purchase Date", "sale_date": "Sale Date",
    "final_supplier": "Supplier", "supplier": "Stock Supplier", "qb_supplier": "QB Supplier",
    "final_buyer": "Buyer", "buyer": "Stock Buyer", "qb_buyer": "QB Buyer",
    "purchase_ref": "Purchase Ref (Stock)", "qb_purchase_ref": "Purchase Ref (QB)", "purchase_link": "Purchase Link",
    "qb_purchase_date": "QB Purchase Date", "purchase_value": "Stock Purchase Value",
    "qb_purchase_value": "QB Purchase Value", "purchase_diff": "Purchase Difference",
    "sale_ref": "Sale Ref (Stock)", "qb_sale_ref": "Sale Ref (QB)", "sale_link": "Sale Link",
    "qb_sale_date": "QB Sale Date", "sale_value": "Stock Sale Value", "qb_sale_value": "QB Sale Value",
    "sale_diff": "Sale Difference", "ap_open": "QB Open Payable", "ar_open": "QB Open Receivable",
    "margin": "Margin", "consignment_invoice": "Consignment Invoice", "bank_collected": "Bank Collections",
    "collected": "Collected (used)", "settlement": "Amount to Settle with Owner",
    "settlement_vs_ap": "Settlement vs QB Payable", "needs_review": "Needs Review",
    "review_reason": "Review Reason", "source": "Source Row",
}
SETTLEMENT_COLUMNS = ["VIN", "Deal Cycle", "Status", "Supplier", "Buyer", "Purchase Date", "Sale Date",
                      "QB Purchase Value", "QB Sale Value", "Amount to Settle with Owner", "Needs Review"]


def summary(recon: pd.DataFrame, disc: pd.DataFrame) -> pd.DataFrame:
    rows = [("Deal cycles", len(recon)), ("Unique VINs", recon["vin"].nunique()),
            ("Cycles needing review", int(recon["needs_review"].sum())), ("Discrepancies", len(disc))]
    rows += [(f"Status: {k}", v) for k, v in recon["status"].value_counts().items()]
    rows += [(f"Issue: {k}", v) for k, v in disc["issue"].value_counts().items()]
    rows += [(f"Purchase link: {k or 'unlinked'}", v) for k, v in recon["purchase_link"].value_counts().items()]
    rows += [(f"Sale link: {k or 'unlinked'}", v) for k, v in recon["sale_link"].value_counts().items()]
    return pd.DataFrame(rows, columns=["Metric", "Value"])


def write_report(result: dict[str, pd.DataFrame], path: str | Path) -> Path:
    recon, disc = result["reconciliation"], result["discrepancies"]
    table = recon[list(RECON_COLUMNS)].rename(columns=RECON_COLUMNS)
    sheets = {
        "Summary": summary(recon, disc),
        "Reconciliation": table,
        "Settlement": table[SETTLEMENT_COLUMNS],
        "Discrepancies": disc.rename(columns={"vin_raw": "VIN", "cycle": "Deal Cycle", "source": "Source Row",
                                              "issue": "Issue", "stock_value": "Stock Value", "qb_value": "QB Value",
                                              "difference": "Difference"}),
        "Review": table.loc[table["Needs Review"], ["VIN", "Deal Cycle", "Status", "Review Reason", "Source Row"]],
    }
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl", datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as xw:
        for name, df in sheets.items():
            df.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            ws.freeze_panes, ws.auto_filter.ref = "B2", ws.dimensions
            for i, col in enumerate(df.columns, 1):
                width = max([len(str(col))] + [len(str(v)) for v in df[col].head(500)])
                ws.column_dimensions[get_column_letter(i)].width = min(width + 2, 60)
    return path
