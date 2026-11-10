"""Generate a small synthetic dataset (stock sheet, QuickBooks exports, bank statement).

Every edge case the engine handles appears at least once; tests rely on these scenarios.
All names, VINs and amounts are fictional.
"""

from __future__ import annotations

import random
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

ALPHABET = "ABCDEFGHJKLMNPRSTUVWXYZ0123456789"
SAMPLE_CONFIG = """\
# Sample configuration: every key not listed here falls back to vinrecon.config.DEFAULTS.
inputs:
  stock: data/stock.xlsx
  general_ledger: data/General Ledger*.xlsx
  purchases_detail: data/Purchases by Supplier Detail.xlsx
  sales_detail: data/Sales by Customer Detail.xlsx
  supplier_balance: data/Supplier Balance Detail*.xlsx
  customer_balance: data/Customer Balance Detail*.xlsx
  bank: data/bank.xlsx
output: output/reconciliation.xlsx
internal_transfer_keywords: [OWN ACCOUNT, INTERNAL TRANSFER]
"""


def vin(rng: random.Random) -> str:
    return "".join(rng.choice(ALPHABET) for _ in range(17))


def scenarios(seed: int = 7) -> list[dict]:
    """Deal cycles. qb_* keys describe what QuickBooks holds; None = absent from QuickBooks."""
    rng = random.Random(seed)
    v = {k: vin(rng) for k in "ABCDEFGHIJKLMNO"}
    d = lambda m, day: pd.Timestamp(2025, m, day)  # noqa: E731
    return [
        # key, vin, cycle, deal type, purchase(date, ref, supplier, value), sale(date, ref, buyer, value), margin, extras
        dict(case="matched_paid", vin=v["A"], deal="Traditional Stock", pd=d(1, 5), pref="PUR-001", sup="Alpha Motors LLC",
             pv=50000, sd=d(2, 1), sref="SAL-001", buy="John Carter", sv=58000, ap=0, ar=0),
        dict(case="purchase_mismatch", vin=v["B"], deal="Traditional Stock", pd=d(1, 8), pref="PUR-002", sup="Beta Cars",
             pv=40000, qb_pv=41500, sd=d(2, 3), sref="SAL-002", buy="Mary Hill", sv=46000, ap=0, ar=0),
        dict(case="blank_stock_supplier", vin=v["C"], deal="Traditional Stock", pd=d(1, 9), pref="PUR-003", sup="",
             qb_sup="Gamma Auto Trading", pv=30000, sd=d(2, 9), sref="SAL-003", buy="", qb_buy="Omar Said", sv=35500, ap=0, ar=0),
        dict(case="messy_reference", vin=v["D"], deal="Traditional Stock", pd=d(1, 11), pref="PUR-004", sup="Delta Motors",
             pv=22000, sd=d(2, 11), sref="SAL- 004.\n", qb_sref="SAL-004", buy="Lena Ford", sv=26000, ap=0, ar=0),
        dict(case="sale_ref_missing_in_stock", vin=v["E"], deal="Traditional Stock", pd=d(1, 12), pref="PUR-005",
             sup="Alpha Motors LLC", pv=61000, sd=d(3, 2), sref="", qb_sref="SAL-005", buy="Ravi Nair", sv=67000, ap=0, ar=0),
        dict(case="sale_not_in_qb", vin=v["F"], deal="Traditional Stock", pd=d(1, 15), pref="PUR-006", sup="Beta Cars",
             pv=18000, sd=d(3, 5), sref="SAL-006", buy="Sara Kim", sv=21000, ap=0, qb_sale=False),
        dict(case="purchase_only", vin=v["G"], deal="Traditional Stock", pd=d(4, 1), pref="PUR-007", sup="Delta Motors",
             pv=27000, ap=27000),
        dict(case="consignment_bank_paid", vin=v["H"], deal="Consignment", pd=d(2, 1), pref="CON-001", sup="Peter Owens",
             sd=d(3, 10), sref="SAL-008", buy="Nina Patel", sv=90000, margin=4000, ar=0, ap=85800, bank=90000),
        dict(case="consignment_qb_ar", vin=v["I"], deal="Consignment", pd=d(2, 2), pref="CON-002", sup="Ali Hassan",
             sd=d(3, 12), sref="SAL-009", buy="Tom Reed", sv=50000, margin=2000, ar=10000),
        dict(case="consignment_no_evidence", vin=v["J"], deal="Consignment", pd=d(2, 3), pref="CON-003", sup="Kate Moss",
             sd=d(3, 14), sref="SAL-010", buy="Yusuf Ali", sv=42000, margin=1500, qb_sale=False),
        dict(case="consignment_not_sold", vin=v["K"], deal="Consignment", pd=d(5, 1), pref="CON-004", sup="Leo Grant"),
        dict(case="cancelled", vin=v["L"], deal="Cancelled / Consignment", pd=d(5, 2), pref="CON-005", sup="Ivy Chen"),
        dict(case="repeat_cycle_1", vin=v["M"], cycle=1, deal="Traditional Stock", pd=d(1, 20), pref="PUR-013",
             sup="Alpha Motors LLC", pv=33000, sd=d(2, 20), sref="SAL-013", buy="Ben Stone", sv=36000, ap=0, ar=0),
        dict(case="repeat_cycle_2", vin=v["M"], cycle=2, deal="Traditional Stock", pd=d(4, 2), pref="PUR-014",
             sup="Ben Stone", pv=30000, sd=d(5, 5), sref="SAL-014", buy="Chloe West", sv=34000, ap=0, ar=3000),
        dict(case="invalid_vin", vin="TRADE-IN-0001", deal="Traditional Stock", pd=d(6, 1), pref="", sup="Walk-in"),
        dict(case="ambiguous_sale", vin=v["N"], deal="Traditional Stock", pd=d(1, 25), pref="PUR-016", sup="Beta Cars",
             pv=25000, sd=d(3, 25), sref="", buy="Anna Bell", sv=29000, ap=0,
             extra_invoices=[("SAL-016A", "Anna Bell", 29000), ("SAL-016B", "Anna Bell", 29000)]),
        dict(case="conflicting_party", vin=v["O"], deal="Traditional Stock", pd=d(1, 26), pref="PUR-017", sup="Zeta Motors",
             pv=15000, qb_sup_alt="Zeta Motor Group", sd=d(3, 26), sref="SAL-017", buy="Dan Fox", sv=17500, ap=0, ar=0),
    ]


def _qb_sheet(path: Path, title: str, period: str, header: list[str], groups: list[tuple[str, list[list]]]):
    """Write a workbook in the QuickBooks 'detail report' layout."""
    wb = Workbook()
    ws = wb.active
    for line in ("Sample Dealer LLC", title, period, None):
        ws.append([line])
    ws.append([None, *header])
    for heading, rows in groups:
        ws.append([heading])
        for row in rows:
            ws.append([None, *row])
        ws.append([f"Total for {heading}"])
    ws.append(["TOTAL"])
    wb.save(path)


def _date(t: pd.Timestamp) -> str:
    return t.strftime("%d/%m/%Y")


def generate(out_dir: str | Path, seed: int = 7) -> Path:
    """Write the sample dataset and a config.yaml into `out_dir`; return the config path."""
    out, deals = Path(out_dir), scenarios(seed)
    data = out / "data"
    data.mkdir(parents=True, exist_ok=True)

    stock = pd.DataFrame([{
        "VIN": x["vin"], "Deal Cycle #": x.get("cycle", 1), "Deal Type(s)": x["deal"],
        "Date of Purchase": x.get("pd"), "Purchase Invoice No": x.get("pref", ""),
        "Purchased From (Seller/Consignor)": x.get("sup", ""), "Purchase Cost": x.get("pv"),
        "Sold Date": x.get("sd"), "Sales Invoice No": x.get("sref", ""), "Sold To (Customer)": x.get("buy", ""),
        "Sales Price": x.get("sv"), "Margin": x.get("margin"),
    } for x in deals])
    stock.to_excel(data / "stock.xlsx", index=False)

    bills = [(x["pd"], x["pref"], x.get("qb_sup") or x["sup"], x.get("qb_pv", x["pv"]), x)
             for x in deals if x.get("pv") and x.get("pref")]
    invoices = [(x["sd"], x.get("qb_sref", x["sref"]), x.get("qb_buy") or x["buy"], x["sv"], x)
                for x in deals if x.get("sv") and x.get("qb_sale", True) and x.get("qb_sref", x["sref"])]
    invoices += [(x["sd"], ref, party, amt, x) for x in deals for ref, party, amt in x.get("extra_invoices", [])]
    desc = lambda x, what: f"{what} Vehicle VIN {x['vin']}"  # noqa: E731

    gl_head = ["Distribution account", "Transaction date", "Transaction type", "Number", "Name", "Description",
               "Split", "Amount", "Balance"]
    ap_rows = [["Accounts Payable (A/P)", _date(t), "Bill", r, p, desc(x, "Purchase of"), "Inventory Asset", a, None]
               for t, r, p, a, x in bills]
    ar_rows = [["Accounts Receivable (A/R)", _date(t), "Invoice", r, p, desc(x, "Sale of"), "Sales", a, None]
               for t, r, p, a, x in invoices]
    inv_rows = [["Inventory Asset", _date(t), "Bill", r, p, desc(x, "Purchase of"), "Accounts Payable (A/P)", a, None]
                for t, r, p, a, x in bills]
    _qb_sheet(data / "General Ledger 2025.xlsx", "General Ledger", "January-December, 2025", gl_head,
              [("Accounts Payable (A/P)", ap_rows), ("Accounts Receivable (A/R)", ar_rows), ("Inventory Asset", inv_rows)])

    def by_party(rows, kind):
        groups: dict[str, list] = {}
        for t, r, p, a, x in rows:
            groups.setdefault(p, []).append([_date(t), kind, r, desc(x, ""), desc(x, ""), 1, a, a, None])
        return sorted(groups.items())

    # A conflicting party: one bill also filed under a second supplier name in the purchases report.
    alt = [(t, r, x["qb_sup_alt"], a, x) for t, r, p, a, x in bills if x.get("qb_sup_alt")]
    for name, rows, kind, rate in (("Purchases by Supplier Detail", bills + alt, "Bill", "Rate"),
                                   ("Sales by Customer Detail", invoices, "Invoice", "Sales price")):
        _qb_sheet(data / f"{name}.xlsx", name, "January-December, 2025",
                  ["Transaction date", "Transaction type", "Number", "Product/Service full name", "Description",
                   "Quantity", rate, "Amount", "Balance"], by_party(rows, kind))

    # Balance snapshots: every document is open at mid-year; year-end lists only what is still unpaid
    # (scenario keys 'ap' / 'ar'). Consignment payouts owed to the owner exist only as open bills here.
    payouts = [(x["sd"], x["pref"], x["sup"], x["ap"], x) for x in deals if x["deal"] == "Consignment" and x.get("ap")]
    head = ["Date", "Transaction type", "Number", "Due date", "Amount", "Open balance", "Balance"]
    for label, rows, key, kind in (("Supplier", bills + payouts, "ap", "Bill"), ("Customer", invoices, "ar", "Invoice")):
        for as_of, tag in (("Jun 30, 2025", "2025-06"), ("Dec 31, 2025", "2025-12")):
            year_end = tag.endswith("12")
            groups: dict[str, list] = {}
            for t, r, p, a, x in rows:
                open_bal = (x.get(key) or 0) if year_end else a
                if open_bal:
                    groups.setdefault(p, []).append([_date(t), kind, r, _date(t), a, open_bal, None])
            _qb_sheet(data / f"{label} Balance Detail {tag}.xlsx", f"{label} Balance Detail Report",
                      f"As of {as_of}", head, sorted(groups.items()))

    wb = Workbook()
    ws = wb.active
    for line in ("Bank transactions", "All accounts", None):
        ws.append([line])
    ws.append(["DATE", "BANK", "Description", "REF#", "EXPENSE", "INCOME", "NARATION"])
    for x in deals:
        if x.get("bank"):
            ws.append([x["sd"], "Main Bank", f"Received for VIN {x['vin']}", "TT-1", None, x["bank"], x["buy"]])
    ws.append([pd.Timestamp(2025, 3, 11), "Main Bank", "OWN ACCOUNT TRANSFER", "TT-2", None, 100000, "internal"])
    wb.save(data / "bank.xlsx")

    (out / "config.yaml").write_text(SAMPLE_CONFIG, encoding="utf-8")
    return out / "config.yaml"
