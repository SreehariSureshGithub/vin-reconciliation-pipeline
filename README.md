# vinrecon: vehicle stock reconciliation

**Reconcile a car dealer's stock sheet against QuickBooks exports and bank statements: per VIN, per deal cycle, with every discrepancy explained.**

Dealers usually track vehicles in a spreadsheet (what was bought, from whom, for how much, and who it was sold to), while the accounts live in QuickBooks. The two drift apart. Invoice numbers get typed with stray spaces, suppliers are left blank, a car is bought back and resold under the same VIN, and consignment owners need to know what they're owed. `vinrecon` reads the raw exports as they come out of QuickBooks and produces one Excel workbook that shows, for every deal:

- which QuickBooks bill and invoice it belongs to, and **how** that link was made
- purchase and sale value differences between the stock sheet and the books
- open payables and receivables from the latest balance snapshot
- the amount to settle with the owner, for consignment deals
- a **review queue** for cases where the evidence is ambiguous, instead of a silent guess

It was built from, and validated against, a real dealership's dataset of about 4,000 deal cycles and 4 years of ledgers. No real data is included here; a synthetic dataset ships with the code.

---

## Key features

- **Works on raw QuickBooks exports.** One parser handles every "detail" report layout: General Ledger, Supplier/Customer Balance Detail, A/P and A/R Ageing, Purchases by Supplier and Sales by Customer. It reads title rows, group headings (account or party), detail lines and "Total for …" rows.
- **Deal-cycle aware.** A VIN bought and sold several times is reconciled cycle by cycle. A bill or invoice is never reused across cycles, and values from different cycles are never mixed.
- **Explainable linking.** Every link is labelled `reference` (exact document number after normalising spaces, case and punctuation), `vin` (the VIN appears in exactly one QuickBooks document), or `vin+amount` (several documents mention the VIN, and exactly one matches the amount).
- **Independent sources.** Stock-sheet and QuickBooks party names are kept side by side and never "corrected" into each other. A fallback label like `Cash Customer` is applied only after every source has been checked.
- **Consignment settlement.** Collected amount − (dealer margin + VAT). Collections come from VIN-tagged bank receipts, falling back to QuickBooks (invoice − open receivable). Internal transfers are excluded.
- **Configurable, not hardcoded.** Column names, file patterns, keywords, tolerances, VAT rate and GL accounts all live in one YAML file.
- **Small and tested.** About 450 lines of library code (it replaces a 44-script, ~44,000-line original) and 19 end-to-end tests covering each business rule. CI runs on Python 3.10–3.13.

## Tech stack

Python 3.10+ · pandas · openpyxl · PyYAML · pytest · ruff · GitHub Actions

---

## Quick start

```bash
git clone https://github.com/SreehariSureshGithub/vin-reconciliation-pipeline.git
cd vin-reconciliation-pipeline
python -m venv .venv
.venv\Scripts\activate          # Windows
# source .venv/bin/activate     # macOS / Linux
pip install -e .

vinrecon sample demo                  # write a synthetic dataset + config into ./demo
vinrecon run -c demo/config.yaml      # -> demo/output/reconciliation.xlsx
```

```
17 deal cycles | 8 discrepancies | 4 need review -> demo/output/reconciliation.xlsx
```

### Use it on your own data
1. Export these reports from QuickBooks to Excel (keep the default layout): General Ledger, Supplier Balance Detail, Customer Balance Detail, Purchases by Supplier Detail, Sales by Customer Detail. A/P and A/R Ageing Detail are optional extras.
2. Put them, your stock sheet and (optionally) a bank statement in a `data/` folder.
3. `cp config.example.yaml config.yaml`, then edit `stock_columns` to match your sheet's headers, and set the file patterns, keywords and tolerances.
4. Run `vinrecon run` (it reads `config.yaml` by default) or `vinrecon run -c path/to/config.yaml -o report.xlsx`.

From Python:
```python
from vinrecon import load_config, run, write_report

result = run(load_config("config.yaml"))      # {"reconciliation": DataFrame, "discrepancies": DataFrame}
write_report(result, "output/reconciliation.xlsx")
```

---

## The output workbook

| Sheet | Contents |
|---|---|
| **Summary** | Counts by status, by issue type and by link method |
| **Reconciliation** | One row per deal cycle: stock and QB values side by side, refs, link method, differences, open balances, settlement, review flag, source row |
| **Settlement** | The short version for sharing: VIN, parties, purchase and sale values, amount to settle with the owner |
| **Discrepancies** | One row per issue, with stock value, QB value and difference |
| **Review** | Cycles a person must look at, each with the reason and candidate evidence |

### Status codes

| Status | Meaning |
|---|---|
| `NORMAL_AR_AP_AVAILABLE` / `_AP_ONLY` / `_AR_ONLY` / `_NO_QB_BALANCE` | Dealer-owned stock, sold; which QuickBooks balance evidence was found |
| `PURCHASE_ONLY` | Dealer-owned and still in stock |
| `CONSIGNMENT_SETTLEMENT_CALCULATED` | Sold on behalf of an owner; settlement computed |
| `CONSIGNMENT_QB_AR_AVAILABLE` | Sold consignment with a QB receivable but no margin to compute commission |
| `CONSIGNMENT_REVIEW` | Sold consignment with no collection evidence (goes to review) |
| `CONSIGNMENT_NOT_SOLD` | Consignment still in stock |
| `CANCELLED` | Deal cancelled; no settlement |

### Issue codes

`INVALID_VIN` · `SALE_DATE_WITHOUT_PRICE` · `PURCHASE_NOT_IN_QB` · `SALE_NOT_IN_QB` · `PURCHASE_VALUE_MISMATCH` · `SALE_VALUE_MISMATCH` · `OPEN_PAYABLE` · `OPEN_RECEIVABLE` · `SETTLEMENT_VS_AP_MISMATCH`

---

## How it works

```
stock sheet ─┐
             ├─ readers.py ─ normalise VINs, refs, names, dates, amounts; keep source row
QuickBooks ──┤               one generic parser for every QB detail-report layout
bank ────────┘
                 │
                 ▼
            core.py
  1. classify each deal cycle   normal / consignment / cancelled, sold or not (sale value or invoice present)
  2. link to QuickBooks         per side (purchase=Bill, sale=Invoice): reference → VIN → VIN+amount
                                document value: GL control-account line > summed detail lines > balance report
  3. parties                    QB party from the linked document; >1 distinct party → review
  4. open balances              latest "As of" snapshot; a document seen only in older snapshots = 0 (cleared)
  5. bank collections           VIN-tagged income lines, minus internal transfers, assigned to the right cycle
  6. settlement + status        consignment: collected − margin × (1 + VAT)
  7. discrepancies + review     rule table → one row per issue; ambiguity → review queue
                 │
                 ▼
            report.py ─ Excel workbook (Summary, Reconciliation, Settlement, Discrepancies, Review)
```

### Design rules
These came out of auditing the original 27-phase version of this pipeline:
1. **Never fill a party before matching.** Early "Cash Supplier" defaults hid real, evidence-backed names in over 700 rows of the original results. The fallback is now the very last step.
2. **Link per transaction, not per VIN.** Collapsing a VIN to its "first purchase / last sale" mixed dates and values across cycles.
3. **Stock and QuickBooks names are independent facts.** A spelling difference isn't an error; it's two sources.
4. **Ambiguity is output, not something to hide.** If two invoices fit equally well, the cycle goes to review with both candidates listed.

### Validation
On the real dataset, compared with the hand-verified final workbook (single-cycle VINs):

| Field | Agreement |
|---|---|
| QuickBooks purchase / sale reference | 99.5% / 99.2% |
| QuickBooks supplier / buyer | 100% / 99.2% |
| QuickBooks purchase value | 100% |
| Open payable / receivable | 99.5% / 98.5% |
| Final status | 91.4% (consignment, cancelled and purchase-only: 100%) |

The remaining status differences are normal-stock deals where the hand-verified file located a balance through free-text General Ledger descriptions, which this version deliberately doesn't do.

---

## Project structure

```
vin-reconciliation-pipeline/
├── src/vinrecon/
│   ├── config.py      # defaults + YAML loading
│   ├── readers.py     # stock sheet, generic QuickBooks report parser, bank statement
│   ├── core.py        # linking, balances, collections, settlement, status, discrepancies
│   ├── report.py      # Excel workbook writer
│   ├── sample.py      # synthetic dataset covering every edge case
│   └── cli.py         # `vinrecon run` / `vinrecon sample`
├── tests/test_pipeline.py   # one test per business rule
├── config.example.yaml
├── pyproject.toml
└── .github/workflows/ci.yml
```

## Development

```bash
pip install -e ".[dev]"
pytest -q
ruff check src tests
```

## Limitations and next steps
- Party names are compared after case and whitespace normalisation only. Fuzzy matching ("Motors LLC" vs "Motor") would reduce false review cases.
- Bank receipts are matched by full 17-character VIN. Partial VINs and invoice numbers in bank narrations are not used yet.
- Only Excel inputs are supported. CSV exports and a QuickBooks Online API reader would be natural additions.
- Single currency; amounts are compared with an absolute tolerance.

## License
[MIT](LICENSE)
