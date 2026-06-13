"""Unit tests for the semantic layer and the insight engine (pure core, no Excel)."""
from __future__ import annotations

import datetime as _dt

from core.insights import InsightKind, Severity, detect_insights
from core.models import ColumnProfile, ColumnType, TableProfile, WorkbookProfile
from core.semantic import MetricKind, ReportType, analyze


def _col(name, ctype, **kw):
    return ColumnProfile(name=name, index=kw.pop("index", 0), ctype=ctype, **kw)


def _hospital_profile():
    """Six months of departmental revenue with a sharp last-month drop and a
    dominant department, so several detectors should fire deterministically."""
    rows = []
    # Months 2026-01 .. 2026-06. Cardiology dominates; June collapses.
    monthly = {
        "2026-01": 100, "2026-02": 110, "2026-03": 120,
        "2026-04": 130, "2026-05": 140, "2026-06": 60,   # big drop
    }
    for period, base in monthly.items():
        y, m = int(period[:4]), int(period[5:])
        d = _dt.datetime(y, m, 15)
        # Cardiology ~70% of revenue, others split the rest.
        rows.append({"Date": d, "Department": "Cardiology", "Revenue": base * 700_000.0})
        rows.append({"Date": d, "Department": "Radiology", "Revenue": base * 150_000.0})
        rows.append({"Date": d, "Department": "Pharmacy", "Revenue": base * 100_000.0})
        rows.append({"Date": d, "Department": "Lab", "Revenue": base * 50_000.0})

    date_c = _col("Date", ColumnType.DATE, index=0)
    dept_c = _col("Department", ColumnType.CATEGORICAL, index=1, distinct=4, count=24)
    rev_c = _col("Revenue", ColumnType.CURRENCY, index=2, number_format='#,##0" LBP"',
                 count=24, total=sum(r["Revenue"] for r in rows),
                 minimum=50 * 50_000.0, maximum=140 * 700_000.0,
                 mean=sum(r["Revenue"] for r in rows) / 24)
    table = TableProfile(
        sheet_name="GL", header_row=1, first_data_row=2, last_data_row=25,
        first_col=1, last_col=3, columns=[date_c, dept_c, rev_c], rows=rows)
    return WorkbookProfile(path="x.xlsx", sheet_names=["GL"], tables=[table])


def test_metric_classification():
    rev = _col("Net Revenue", ColumnType.CURRENCY)
    cost = _col("Salary Expense", ColumnType.CURRENCY)
    bal = _col("Outstanding Receivable", ColumnType.CURRENCY)
    vol = _col("Admissions", ColumnType.NUMERIC, minimum=0, maximum=500, mean=120)
    pct = _col("Occupancy %", ColumnType.PERCENT)
    from core.semantic import classify_metric
    assert classify_metric(rev) == MetricKind.REVENUE
    assert classify_metric(cost) == MetricKind.COST
    assert classify_metric(bal) == MetricKind.BALANCE
    assert classify_metric(vol) == MetricKind.VOLUME
    assert classify_metric(pct) == MetricKind.RATIO


def test_semantic_model_picks_primary_money():
    profile = _hospital_profile()
    sem = analyze(profile)
    assert sem.primary_money is not None
    assert sem.primary_money.name == "Revenue"
    # Revenue header isn't explicitly "revenue"-worded here, but currency → AMOUNT
    # is still money and becomes the primary measure.
    assert sem.primary_money.kind.is_money


def test_variance_detects_the_drop():
    profile = _hospital_profile()
    sem = analyze(profile)
    insights = detect_insights(profile, sem)
    var = [i for i in insights if i.kind == InsightKind.VARIANCE]
    assert var, "expected a variance insight for the June collapse"
    top = var[0]
    assert top.severity == Severity.HIGH
    assert top.good is False or top.good is None
    assert top.period == "2026-06"
    # The driver of the drop should be the dominant department.
    assert top.evidence.get("driver") == "Cardiology"


def test_concentration_flags_dominant_department():
    profile = _hospital_profile()
    sem = analyze(profile)
    insights = detect_insights(profile, sem)
    conc = [i for i in insights if i.kind == InsightKind.CONCENTRATION]
    assert conc, "expected a concentration insight"
    assert conc[0].evidence["leader"] == "Cardiology"
    assert conc[0].evidence["leader_share"] > 0.5


def test_insights_are_ranked_and_capped():
    profile = _hospital_profile()
    sem = analyze(profile)
    insights = detect_insights(profile, sem, max_insights=5)
    assert 0 < len(insights) <= 5
    # Sorted by (severity, score) descending.
    keys = [i.sort_key() for i in insights]
    assert keys == sorted(keys, reverse=True)


def _revenue_profile():
    """A revenue book: ACTTNUMB revenue account codes + NEGATIVE LBP amounts."""
    codes = ["719300100001", "713000100201", "713000100202", "713000100203"]
    rows = []
    for m in range(1, 7):
        for j, code in enumerate(codes):
            d = _dt.datetime(2026, m, 15)
            amt = (100 + m * 10 + j) * 100_000.0
            rows.append({"TRDATE": d, "ACTTNUMB": code,
                         "ORG_AMOUNT": -amt, "USD": amt / 90000})
    date_c = _col("TRDATE", ColumnType.DATE, index=0)
    acct_c = _col("ACTTNUMB", ColumnType.CATEGORICAL, index=1, distinct=4, count=24)
    org_c = _col("ORG_AMOUNT", ColumnType.CURRENCY, index=2, number_format='#,##0" LBP"',
                 count=24, total=sum(r["ORG_AMOUNT"] for r in rows),
                 mean=sum(r["ORG_AMOUNT"] for r in rows) / 24,
                 minimum=min(r["ORG_AMOUNT"] for r in rows),
                 maximum=max(r["ORG_AMOUNT"] for r in rows))
    usd_c = _col("USD", ColumnType.CURRENCY, index=3, number_format='"$"#,##0.00',
                 count=24, total=sum(r["USD"] for r in rows))
    table = TableProfile(sheet_name="GL", header_row=1, first_data_row=2,
                         last_data_row=25, first_col=1, last_col=4,
                         columns=[date_c, acct_c, org_c, usd_c], rows=rows)
    return WorkbookProfile(path="x.xlsx", sheet_names=["GL"], tables=[table])


def test_purpose_detected_as_revenue_from_account_codes():
    from core.library import get_library
    profile = _revenue_profile()
    sem = analyze(profile, get_library())
    # The library knows 713…/719… are revenue accounts → purpose is Revenue.
    assert sem.purpose == "Revenue"
    assert sem.purpose_kind == MetricKind.REVENUE
    assert sem.account_column == "ACTTNUMB"


def test_revenue_sign_flip_makes_amounts_positive():
    from core.pipeline import _apply_revenue_sign
    profile = _revenue_profile()
    helpers = _apply_revenue_sign(profile.primary)
    org = profile.primary.column("ORG_AMOUNT")
    usd = profile.primary.column("USD")
    # LBP revenue column flipped positive, with a hidden positive helper named.
    assert org.total > 0 and org.positive_helper == "ORG_AMOUNT (+)"
    assert ("ORG_AMOUNT", "ORG_AMOUNT (+)") in helpers
    assert all(r["ORG_AMOUNT"] > 0 for r in profile.primary.rows)
    # The already-positive USD/$ column is left untouched (no helper).
    assert usd.positive_helper is None


def test_empty_profile_is_safe():
    empty = WorkbookProfile(path="x.xlsx")
    sem = analyze(empty)
    assert sem.report_type == ReportType.GENERIC
    assert detect_insights(empty, sem) == []
