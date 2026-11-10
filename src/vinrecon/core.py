"""Reconciliation engine: stock deal cycles vs QuickBooks transactions vs bank money flow."""

from __future__ import annotations

import re
from collections import defaultdict

import pandas as pd

from . import readers
from .readers import VIN_RE, norm_name, norm_ref

VALID_VIN = re.compile(r"[A-HJ-NPR-Z0-9]{17}")


class TxnIndex:
    """Purchase- or sale-side QuickBooks transactions keyed by normalised reference."""

    def __init__(self, qb: pd.DataFrame, types: list[str], account: str, detail_kind: str):
        t = qb[qb["type"].isin(types) & (qb["ref"] != "")]
        gl = t[(t["kind"] == "general_ledger") & t["account"].str.contains(account, case=False, regex=False)]
        detail = t[t["kind"] == detail_kind]
        # Priority: ledger control-account line (document total) > summed detail-report lines >
        # document amount in balance reports (covers documents that only appear there).
        other = t[~t["kind"].isin(["general_ledger", detail_kind])]
        self.amount = (gl.groupby("ref")["amount"].sum().abs()
                       .combine_first(detail.groupby("ref")["amount"].sum())
                       .combine_first(other.groupby("ref")["amount"].max()))
        self.date = t.groupby("ref")["date"].min()
        self.raw = t.groupby("ref")["ref_raw"].first()
        self.parties: dict[str, list[str]] = {}
        for ref, names in t[t["party"] != ""].groupby("ref")["party"]:
            self.parties[ref] = list({norm_name(n): n for n in names}.values())
        self.by_vin: dict[str, set[str]] = defaultdict(set)
        for ref, text in zip(t["ref"], t["text"]):
            for vin in VIN_RE.findall(text):
                self.by_vin[vin].add(ref)

    def link(self, vin: str, ref: str, value, used: set[str], tol: float, side: str):
        """Return (ref, method, review_reason) for one deal cycle."""
        if ref and ref in self.amount.index:
            return ref, "reference", None
        cands = sorted(self.by_vin.get(vin, set()) - used)
        by_amount = [r for r in cands if value is not None and abs(self.amount.get(r, float("nan")) - value) <= tol]
        if len(cands) == 1:
            return cands[0], "vin", None
        if len(by_amount) == 1:
            return by_amount[0], "vin+amount", None
        if cands:
            return None, "", f"{len(cands)} {side} transactions mention this VIN: {', '.join(self.raw[c] for c in cands)}"
        return None, "", None


def open_balances(qb: pd.DataFrame, kind: str):
    """ref -> open balance in the latest snapshot; 0.0 if the ref only appears in older snapshots."""
    b = qb[(qb["kind"] == kind) & (qb["ref"] != "")]
    if b.empty:
        return lambda ref: None
    latest = b[b["as_of"] == b["as_of"].max()] if b["as_of"].notna().any() else b
    # Sum lines within each file, then one value per ref (balance detail and ageing reports overlap).
    current = latest.groupby(["file", "ref"])["open_balance"].sum().groupby(level="ref").first().to_dict()
    seen = set(b["ref"])
    return lambda ref: current.get(ref, 0.0 if ref in seen else None)


def bank_collections(bank: pd.DataFrame, stock: pd.DataFrame, internal: list[str]) -> pd.Series:
    """Customer money received per stock row: VIN-tagged income lines, assigned to the latest cycle bought on/before."""
    total = pd.Series(0.0, index=stock.index)
    if bank.empty:
        return total.where(False)
    income = bank[(bank["income"].fillna(0) > 0) & ~bank["text"].map(lambda t: any(k in t for k in internal))]
    cycles = stock.groupby("vin")
    for text, date, amount in zip(income["text"], income["date"], income["income"]):
        for vin in set(VIN_RE.findall(text)) & set(cycles.groups):
            rows = cycles.get_group(vin)
            earlier = rows[rows["purchase_date"].isna() | (rows["purchase_date"] <= date)] if pd.notna(date) else rows
            total[(earlier if len(earlier) else rows).index[-1]] += amount
    return total


def status(r) -> str:
    if r.is_cancelled:
        return "CANCELLED"
    if not r.is_sold:
        return "CONSIGNMENT_NOT_SOLD" if r.is_consignment else "PURCHASE_ONLY"
    if r.is_consignment:
        if pd.notna(r.settlement):
            return "CONSIGNMENT_SETTLEMENT_CALCULATED"
        return "CONSIGNMENT_QB_AR_AVAILABLE" if pd.notna(r.ar_open) else "CONSIGNMENT_REVIEW"
    ap, ar = pd.notna(r.ap_open), pd.notna(r.ar_open)
    return {(True, True): "NORMAL_AR_AP_AVAILABLE", (True, False): "NORMAL_AP_ONLY",
            (False, True): "NORMAL_AR_ONLY"}.get((ap, ar), "NORMAL_NO_QB_BALANCE")


def reconcile(stock: pd.DataFrame, qb: pd.DataFrame, bank: pd.DataFrame, cfg: dict) -> dict[str, pd.DataFrame]:
    tol, s = cfg["amount_tolerance"], stock.copy()
    has = lambda words: s["deal_type"].str.upper().map(lambda t: any(w in t for w in words))  # noqa: E731
    s["is_cancelled"], s["is_consignment"] = has(cfg["cancelled_keywords"]), has(cfg["consignment_keywords"])
    s["is_sold"] = s["sale_value"].notna() | (s["sale_ref"] != "")  # a date alone is not a sale
    s["valid_vin"] = s["vin"].map(lambda v: bool(VALID_VIN.fullmatch(v)))

    sides = {"purchase": TxnIndex(qb, cfg["purchase_types"], cfg["ap_account"], "purchases_detail"),
             "sale": TxnIndex(qb, cfg["sale_types"], cfg["ar_account"], "sales_detail")}
    reviews = defaultdict(list)
    for side, idx in sides.items():
        party_col = "supplier" if side == "purchase" else "buyer"
        out, used = [], defaultdict(set)
        for i, r in s.iterrows():
            ref, method, reason = idx.link(r.vin, norm_ref(r[f"{side}_ref"]), r[f"{side}_value"], used[r.vin], tol, side)
            if ref:
                used[r.vin].add(ref)
            parties = idx.parties.get(ref, []) if ref else []
            if len(parties) > 1:
                reason = f"{side} reference {idx.raw[ref]} has several QuickBooks parties: {' / '.join(parties)}"
            if reason:
                reviews[i].append(reason)
            out.append((idx.raw.get(ref, "") if ref else "", method, idx.amount.get(ref) if ref else None,
                        idx.date.get(ref) if ref else pd.NaT, parties[0] if parties else ""))
        cols = [f"qb_{side}_ref", f"{side}_link", f"qb_{side}_value", f"qb_{side}_date", f"qb_{party_col}"]
        s[cols] = pd.DataFrame(out, index=s.index, columns=cols)
        s[f"{side}_diff"] = s[f"{side}_value"] - s[f"qb_{side}_value"]
        # Source fields stay independent; the fallback label is applied only when no source names a party.
        s[f"final_{party_col}"] = [q or st or cfg[f"fallback_{party_col}"] for q, st in zip(s[f"qb_{party_col}"], s[party_col])]

    ap, ar = open_balances(qb, "supplier_balance"), open_balances(qb, "customer_balance")
    s["ap_open"] = [ap(norm_ref(r)) if r else None for r in s["qb_purchase_ref"]]
    s["ar_open"] = [ar(norm_ref(r)) if r else None for r in s["qb_sale_ref"]]
    s["bank_collected"] = bank_collections(bank, s, cfg["internal_transfer_keywords"])

    # Consignment: the owner is owed what was collected minus the dealer's commission (margin + VAT).
    live_cons = s["is_consignment"] & ~s["is_cancelled"] & s["is_sold"]
    s["consignment_invoice"] = (s["margin"] * (1 + cfg["consignment_vat_rate"])).where(live_cons)
    paid_per_qb = (s["qb_sale_value"].fillna(s["sale_value"]) - s["ar_open"]).where(s["ar_open"].notna())
    collected = s["bank_collected"].where(s["bank_collected"] > 0).combine_first(paid_per_qb)
    s["collected"] = collected.where(live_cons)
    s["settlement"] = s["collected"] - s["consignment_invoice"]
    s["settlement_vs_ap"] = s["settlement"] - s["ap_open"]
    s["status"] = [status(r) for r in s.itertuples()]

    for i in s.index[~s["valid_vin"]]:
        reviews[i].append("identifier is not a valid 17-character VIN")
    for i in s.index[s["status"] == "CONSIGNMENT_REVIEW"]:
        reviews[i].append("sold consignment without collection evidence (no bank receipt or QuickBooks A/R)")
    s["review_reason"] = [" | ".join(reviews.get(i, [])) for i in s.index]
    s["needs_review"] = s["review_reason"] != ""
    return {"reconciliation": s, "discrepancies": discrepancies(s, tol)}


def discrepancies(s: pd.DataFrame, tol: float) -> pd.DataFrame:
    normal = ~s["is_consignment"] & ~s["is_cancelled"]
    rules = [  # (code, mask, stock value col, qb value col)
        ("INVALID_VIN", ~s["valid_vin"], None, None),
        ("SALE_DATE_WITHOUT_PRICE", s["sale_date"].notna() & ~s["is_sold"] & ~s["is_cancelled"], None, None),
        ("PURCHASE_NOT_IN_QB", normal & (s["purchase_ref"] != "") & (s["qb_purchase_ref"] == ""), "purchase_value", None),
        ("SALE_NOT_IN_QB", ~s["is_cancelled"] & s["is_sold"] & (s["sale_ref"] != "") & (s["qb_sale_ref"] == ""), "sale_value", None),
        ("PURCHASE_VALUE_MISMATCH", normal & (s["purchase_diff"].abs() > tol), "purchase_value", "qb_purchase_value"),
        ("SALE_VALUE_MISMATCH", ~s["is_cancelled"] & (s["sale_diff"].abs() > tol), "sale_value", "qb_sale_value"),
        ("OPEN_PAYABLE", s["ap_open"].abs() > tol, None, "ap_open"),
        ("OPEN_RECEIVABLE", s["ar_open"].abs() > tol, None, "ar_open"),
        ("SETTLEMENT_VS_AP_MISMATCH", s["settlement_vs_ap"].abs() > tol, "settlement", "ap_open"),
    ]
    frames = []
    for code, mask, a, b in rules:
        hit = s[mask.fillna(False).astype(bool)]
        frames.append(pd.DataFrame({"vin_raw": hit["vin_raw"], "cycle": hit["cycle"], "source": hit["source"],
                                    "issue": code, "stock_value": hit[a] if a else float("nan"),
                                    "qb_value": hit[b] if b else float("nan")}))
    d = pd.concat(frames, ignore_index=True).astype({"stock_value": float, "qb_value": float})
    d["difference"] = d["stock_value"] - d["qb_value"]
    return d.sort_values(["vin_raw", "cycle", "issue"], ignore_index=True)


def run(cfg: dict) -> dict[str, pd.DataFrame]:
    return reconcile(readers.read_stock(cfg), readers.read_qb(cfg), readers.read_bank(cfg), cfg)
