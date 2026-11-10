"""End-to-end business-rule tests on the synthetic sample dataset."""

import pandas as pd
import pytest

from vinrecon import load_config, run, write_report
from vinrecon.cli import main
from vinrecon.sample import generate, scenarios


@pytest.fixture(scope="module")
def result(tmp_path_factory):
    cfg = load_config(generate(tmp_path_factory.mktemp("sample")))
    out = run(cfg)
    rows = {f"stock.xlsx row {i + 2}": x["case"] for i, x in enumerate(scenarios())}
    out["reconciliation"]["case"] = out["reconciliation"]["source"].map(rows)
    return out


def case(result, name) -> pd.Series:
    r = result["reconciliation"]
    return r[r["case"] == name].iloc[0]


def issues(result, name) -> set[str]:
    r, d = result["reconciliation"], result["discrepancies"]
    return set(d.loc[d["source"] == r.loc[r["case"] == name, "source"].iloc[0], "issue"])


def test_matched_deal_is_clean(result):
    r = case(result, "matched_paid")
    assert (r.status, r.purchase_link, r.sale_link) == ("NORMAL_AR_AP_AVAILABLE", "reference", "reference")
    assert r.purchase_diff == 0 and r.sale_diff == 0 and r.ap_open == 0 and r.ar_open == 0
    assert not r.needs_review and issues(result, "matched_paid") == set()


def test_value_mismatch_reported(result):
    assert case(result, "purchase_mismatch").purchase_diff == -1500
    assert "PURCHASE_VALUE_MISMATCH" in issues(result, "purchase_mismatch")


def test_blank_stock_party_filled_from_quickbooks_not_cash_fallback(result):
    r = case(result, "blank_stock_supplier")
    assert (r.supplier, r.qb_supplier, r.final_supplier) == ("", "Gamma Auto Trading", "Gamma Auto Trading")
    assert r.final_buyer == "Omar Said"


def test_cash_fallback_only_when_no_source_names_a_party(result):
    assert case(result, "purchase_only").final_buyer == "Cash Customer"


def test_messy_reference_still_matches(result):
    r = case(result, "messy_reference")
    assert (r.qb_sale_ref, r.sale_link) == ("SAL-004", "reference")


def test_missing_stock_reference_linked_through_vin(result):
    r = case(result, "sale_ref_missing_in_stock")
    assert (r.qb_sale_ref, r.sale_link) == ("SAL-005", "vin")


def test_sale_absent_from_quickbooks(result):
    assert "SALE_NOT_IN_QB" in issues(result, "sale_not_in_qb")
    assert case(result, "sale_not_in_qb").status == "NORMAL_AP_ONLY"


def test_open_payable_from_latest_snapshot(result):
    r = case(result, "purchase_only")
    assert r.status == "PURCHASE_ONLY" and r.ap_open == 27000
    assert "OPEN_PAYABLE" in issues(result, "purchase_only")


def test_consignment_settlement_from_bank_collection(result):
    r = case(result, "consignment_bank_paid")
    assert r.consignment_invoice == pytest.approx(4200)  # margin 4000 + 5% VAT
    assert r.settlement == pytest.approx(85800) and r.settlement_vs_ap == pytest.approx(0)
    assert r.status == "CONSIGNMENT_SETTLEMENT_CALCULATED"


def test_consignment_settlement_from_quickbooks_receivable(result):
    r = case(result, "consignment_qb_ar")  # 50,000 sale - 10,000 still owed - 2,100 commission
    assert r.collected == pytest.approx(40000) and r.settlement == pytest.approx(37900)


def test_consignment_without_evidence_goes_to_review(result):
    r = case(result, "consignment_no_evidence")
    assert r.status == "CONSIGNMENT_REVIEW" and r.needs_review and pd.isna(r.settlement)


@pytest.mark.parametrize("name,status", [("consignment_not_sold", "CONSIGNMENT_NOT_SOLD"), ("cancelled", "CANCELLED")])
def test_no_settlement_for_unsold_or_cancelled(result, name, status):
    r = case(result, name)
    assert r.status == status and pd.isna(r.settlement)


def test_repeat_sale_cycles_are_linked_separately(result):
    c1, c2 = case(result, "repeat_cycle_1"), case(result, "repeat_cycle_2")
    assert c1.vin == c2.vin and (c1.cycle, c2.cycle) == (1, 2)
    assert (c1.qb_purchase_ref, c1.qb_sale_ref) == ("PUR-013", "SAL-013")
    assert (c2.qb_purchase_ref, c2.qb_sale_ref, c2.ar_open) == ("PUR-014", "SAL-014", 3000)


def test_invalid_vin_flagged(result):
    r = case(result, "invalid_vin")
    assert not r.valid_vin and r.needs_review and "INVALID_VIN" in issues(result, "invalid_vin")


def test_ambiguous_vin_candidates_go_to_review(result):
    r = case(result, "ambiguous_sale")
    assert r.qb_sale_ref == "" and "SAL-016A" in r.review_reason and "SAL-016B" in r.review_reason


def test_conflicting_quickbooks_parties_go_to_review(result):
    r = case(result, "conflicting_party")
    assert r.needs_review and "Zeta Motor Group" in r.review_reason


def test_internal_transfers_are_not_collections(result):
    assert result["reconciliation"]["bank_collected"].sum() == 90000


def test_report_and_cli(tmp_path, result):
    path = write_report(result, tmp_path / "report.xlsx")
    sheets = pd.read_excel(path, sheet_name=None)
    assert list(sheets) == ["Summary", "Reconciliation", "Settlement", "Discrepancies", "Review"]
    assert len(sheets["Reconciliation"]) == len(scenarios()) and len(sheets["Review"]) == 4
    config = generate(tmp_path / "cli")
    assert main(["run", "-c", str(config)]) == 0 and (tmp_path / "cli" / "output" / "reconciliation.xlsx").exists()
    assert main(["run", "-c", str(tmp_path / "missing.yaml")]) == 1
