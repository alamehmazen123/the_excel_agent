"""Insight detectors — explainable statistics over a profiled workbook.

Each detector reads the already-loaded :class:`TableProfile` (no Excel, no
network) and emits :class:`Insight` objects. The methods are deliberately simple
and auditable — period-over-period variance, Pareto concentration, MAD/z-score
anomalies, a least-squares trend with a one-step forecast, and receivables
ageing — because a hospital CFO must be able to trust and reproduce every claim.

Public entry point: :func:`detect_insights`.
"""
from __future__ import annotations

import datetime as _dt
import math
from collections import defaultdict
from typing import Any, Optional

from ..aggregate import _period_key, group_sum, time_series
from ..context import event_for_period
from ..formatting import fmt_measure, fmt_number
from ..models import ColumnProfile, ColumnType, TableProfile, WorkbookProfile
from ..semantic import MeasureSemantic, MetricKind, ReportType, SemanticModel
from .models import Insight, InsightKind, Severity

# Tuning thresholds (fractions / sigmas). Conservative so findings stay credible.
_VAR_WATCH = 0.08          # >=8% period change is worth a look
_VAR_HIGH = 0.20           # >=20% is a headline
_CONC_WATCH = 0.35         # leader holds >=35% of the measure
_CONC_HIGH = 0.50          # >=50% = serious single-point dependence
_ANOM_WATCH = 2.5          # robust z-score (MAD-based)
_ANOM_HIGH = 3.5
_AGING_WATCH = 0.15        # >=15% of balance is 90+ days
_AGING_HIGH = 0.30
_MIN_PERIODS = 3           # need at least this many months for trend/anomaly


# --------------------------------------------------------------------------- #
# Small statistics helpers (kept local; no numpy dependency)                  #
# --------------------------------------------------------------------------- #
def _mean(xs: list[float]) -> float:
    return sum(xs) / len(xs) if xs else 0.0


def _median(xs: list[float]) -> float:
    if not xs:
        return 0.0
    s = sorted(xs)
    n = len(s)
    mid = n // 2
    return s[mid] if n % 2 else (s[mid - 1] + s[mid]) / 2


def _mad(xs: list[float], med: float) -> float:
    """Median absolute deviation — a robust, outlier-resistant spread."""
    if not xs:
        return 0.0
    return _median([abs(x - med) for x in xs])


def _linfit(ys: list[float]) -> tuple[float, float]:
    """Least-squares slope + intercept of ys against x = 0..n-1."""
    n = len(ys)
    if n < 2:
        return 0.0, ys[0] if ys else 0.0
    xs = list(range(n))
    mx, my = _mean(xs), _mean(ys)
    denom = sum((x - mx) ** 2 for x in xs)
    if denom == 0:
        return 0.0, my
    slope = sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / denom
    return slope, my - slope * mx


def _pct(numer: float, denom: float) -> Optional[float]:
    return numer / abs(denom) if denom else None


def _clamp01(x: float) -> float:
    return max(0.0, min(1.0, x))


# --------------------------------------------------------------------------- #
# Detectors                                                                   #
# --------------------------------------------------------------------------- #
def _variance(table: TableProfile, sem: SemanticModel,
              date_col: ColumnProfile) -> list[Insight]:
    """Period-over-period change for each money measure, with the driver."""
    out: list[Insight] = []
    for ms in sem.of_kind(MetricKind.REVENUE, MetricKind.COST,
                          MetricKind.AMOUNT, MetricKind.BALANCE):
        series = time_series(table, date_col, ms.column)
        if len(series) < 2:
            continue
        (pp, prev), (lp, last) = series[-2], series[-1]
        change = last - prev
        frac = _pct(change, prev)
        if frac is None or abs(frac) < _VAR_WATCH:
            continue
        sev = (Severity.HIGH if abs(frac) >= _VAR_HIGH
               else Severity.WATCH)
        up = change >= 0
        hib = ms.kind.higher_is_better
        good = None if hib is None else (up == hib)
        arrow = "rose" if up else "fell"
        driver = _top_driver(table, ms.column, date_col, lp, pp)
        detail = (f"{ms.meaning} {arrow} {abs(frac) * 100:.0f}% "
                  f"({fmt_measure(ms.column, abs(change))}) in {lp} vs {pp}"
                  f"{f', led by {driver}' if driver else ''}.")
        score = _clamp01(abs(frac))
        # If this swing involves a KNOWN closure month, it's expected — explain it
        # and stop flagging it as an alarming finding.
        ev = event_for_period(lp) or event_for_period(pp)
        if ev is not None:
            detail += f" This reflects the {ev.short}, not an operational problem."
            sev = Severity.WATCH if sev == Severity.HIGH else sev
            good = None
            score *= 0.4
        out.append(Insight(
            kind=InsightKind.VARIANCE, severity=sev,
            title=f"{ms.meaning} {arrow} {abs(frac) * 100:.0f}% month-over-month",
            detail=detail, score=score, good=good,
            measure=ms.name, period=lp,
            evidence={"prev": prev, "last": last, "prev_period": pp,
                      "driver": driver},
        ))
    return out


def _top_driver(table: TableProfile, measure: ColumnProfile,
                date_col: ColumnProfile, last_period: str,
                prev_period: str) -> Optional[str]:
    """Which dimension item moved the measure most between the two periods."""
    best_label: Optional[str] = None
    best_delta = 0.0
    for dim in _dimensions(table)[:4]:
        sums: dict[str, list[float]] = defaultdict(lambda: [0.0, 0.0])
        for row in table.rows:
            p = _period_key(row.get(date_col.name))
            key = row.get(dim.name)
            val = row.get(measure.name)
            if key in (None, "") or not isinstance(val, (int, float)) or isinstance(val, bool):
                continue
            if p == last_period:
                sums[str(key)][0] += float(val)
            elif p == prev_period:
                sums[str(key)][1] += float(val)
        for label, (lv, pv) in sums.items():
            d = lv - pv
            if abs(d) > abs(best_delta):
                best_delta, best_label = d, label
    return best_label


def _concentration(table: TableProfile, sem: SemanticModel) -> list[Insight]:
    """Pareto: how concentrated the primary money measure is per dimension."""
    pm = sem.primary_money
    if pm is None:
        return []
    out: list[Insight] = []
    for dim in _dimensions(table)[:4]:
        ranked = group_sum(table, dim, pm.column, top_n=10_000)
        ranked = [(k, v) for k, v in ranked if v > 0]
        ranked.sort(key=lambda kv: kv[1], reverse=True)
        total = sum(v for _, v in ranked)
        if total <= 0 or len(ranked) < 3:
            continue
        leader_share = ranked[0][1] / total
        # How many items reach 80% cumulative.
        cum, n80 = 0.0, 0
        for _, v in ranked:
            cum += v
            n80 += 1
            if cum >= 0.8 * total:
                break
        if leader_share < _CONC_WATCH and n80 > max(3, len(ranked) * 0.3):
            continue
        sev = (Severity.HIGH if leader_share >= _CONC_HIGH
               else Severity.WATCH if leader_share >= _CONC_WATCH
               else Severity.INFO)
        dim_label = _label(dim.name)
        out.append(Insight(
            kind=InsightKind.CONCENTRATION, severity=sev,
            title=f"{n80} of {len(ranked)} {dim_label} drive 80% of {pm.meaning}",
            detail=(f"{ranked[0][0]} alone is {leader_share * 100:.0f}% of "
                    f"{pm.meaning} ({fmt_measure(pm.column, ranked[0][1])}); "
                    f"the top {n80} of {len(ranked)} {dim_label} make up 80%."),
            score=_clamp01(leader_share), good=False if sev == Severity.HIGH else None,
            measure=pm.name, dimension=dim_label,
            evidence={"leader": ranked[0][0], "leader_share": leader_share,
                      "n80": n80, "items": ranked[:8]},
        ))
    return out


def _anomaly(table: TableProfile, sem: SemanticModel,
             date_col: ColumnProfile) -> list[Insight]:
    """Flag months where the primary measure is far from its robust band."""
    pm = sem.primary_money
    if pm is None:
        return []
    series = time_series(table, date_col, pm.column)
    if len(series) < _MIN_PERIODS + 1:
        return []
    vals = [v for _, v in series]
    med = _median(vals)
    mad = _mad(vals, med)
    if mad == 0:
        return []
    out: list[Insight] = []
    for period, v in series:
        z = 0.6745 * (v - med) / mad        # robust z-score
        if abs(z) < _ANOM_WATCH:
            continue
        sev = Severity.HIGH if abs(z) >= _ANOM_HIGH else Severity.WATCH
        direction = "above" if z > 0 else "below"
        detail = (f"{pm.meaning} in {period} ({fmt_measure(pm.column, v)}) sits "
                  f"{abs(z):.1f}σ {direction} its typical monthly level "
                  f"({fmt_measure(pm.column, med)}).")
        score = _clamp01(abs(z) / 5.0)
        ev = event_for_period(period)
        if ev is not None:
            detail += f" This is explained by the {ev.short} — expected, not anomalous."
            sev = Severity.INFO
            score *= 0.3
        out.append(Insight(
            kind=InsightKind.ANOMALY, severity=sev,
            title=f"{pm.meaning} in {period} is unusual",
            detail=detail, score=score, good=None,
            measure=pm.name, period=period,
            evidence={"z": z, "median": med, "value": v},
        ))
    # Keep only the most extreme anomaly to avoid noise.
    out.sort(key=lambda i: i.score, reverse=True)
    return out[:1]


def _trend(table: TableProfile, sem: SemanticModel,
           date_col: ColumnProfile) -> list[Insight]:
    """Least-squares trend + one-step forecast for the primary measure."""
    pm = sem.primary_money
    if pm is None:
        return []
    series = time_series(table, date_col, pm.column)
    if len(series) < _MIN_PERIODS:
        return []
    vals = [v for _, v in series]
    slope, intercept = _linfit(vals)
    avg = _mean(vals)
    if avg == 0:
        return []
    monthly_pct = slope / abs(avg)
    if abs(monthly_pct) < 0.03:             # < 3%/month = effectively flat
        return []
    # Phase 6.4: Configurable forecast horizon (default 1 period ahead).
    forecast_horizon = getattr(_trend, "forecast_horizon", 1)
    forecast = slope * (len(vals) + forecast_horizon - 1) + intercept
    up = slope > 0
    hib = pm.kind.higher_is_better
    good = None if hib is None else (up == hib)
    horizon_label = "month" if forecast_horizon == 1 else f"{forecast_horizon} months"
    return [Insight(
        kind=InsightKind.TREND, severity=Severity.WATCH,
        title=f"{pm.meaning} trending {'up' if up else 'down'} "
              f"~{abs(monthly_pct) * 100:.0f}%/month",
        detail=(f"Over {len(vals)} months {pm.meaning} is trending "
                f"{'upward' if up else 'downward'} about "
                f"{abs(monthly_pct) * 100:.0f}% per month; next {horizon_label} projects to "
                f"~{fmt_measure(pm.column, max(0.0, forecast))}."),
        score=_clamp01(abs(monthly_pct)), good=good,
        measure=pm.name,
        evidence={"slope": slope, "forecast": forecast, "series": series,
                  "forecast_horizon": forecast_horizon},
    )]


# Allow setting the forecast horizon from outside (e.g. pipeline or UI).
_trend.forecast_horizon = 1  # default: 1 period ahead


def _aging(table: TableProfile, sem: SemanticModel,
           date_col: ColumnProfile) -> list[Insight]:
    """Receivables ageing: bucket a balance by age from the latest date."""
    bal = sem.balance
    if bal is None and sem.report_type != ReportType.RECEIVABLES:
        return []
    measure = bal.column if bal else sem.primary_money.column if sem.primary_money else None
    if measure is None:
        return []
    latest = _latest_date(table, date_col)
    if latest is None:
        return []
    buckets = {"0–30": 0.0, "31–60": 0.0, "61–90": 0.0, "90+": 0.0}
    total = 0.0
    for row in table.rows:
        d = row.get(date_col.name)
        v = row.get(measure.name)
        if not isinstance(v, (int, float)) or isinstance(v, bool):
            continue
        age = _age_days(d, latest)
        if age is None or v <= 0:
            continue
        total += v
        if age <= 30:
            buckets["0–30"] += v
        elif age <= 60:
            buckets["31–60"] += v
        elif age <= 90:
            buckets["61–90"] += v
        else:
            buckets["90+"] += v
    if total <= 0:
        return []
    over90 = buckets["90+"] / total
    if over90 < _AGING_WATCH:
        sev = Severity.INFO
    elif over90 < _AGING_HIGH:
        sev = Severity.WATCH
    else:
        sev = Severity.HIGH
    if sev == Severity.INFO:
        return []
    return [Insight(
        kind=InsightKind.AGING, severity=sev,
        title=f"{over90 * 100:.0f}% of {bal.meaning if bal else measure.name} is 90+ days old",
        detail=(f"Of {fmt_measure(measure, total)} outstanding, "
                f"{fmt_measure(measure, buckets['90+'])} ({over90 * 100:.0f}%) is "
                f"more than 90 days old — at risk of becoming uncollectable."),
        score=_clamp01(over90), good=False,
        measure=measure.name,
        evidence={"buckets": buckets, "total": total},
    )]


def _losses(table: TableProfile, sem: SemanticModel) -> list[Insight]:
    """Negative records on revenue/amount measures (refunds, write-offs, losses)."""
    out: list[Insight] = []
    for ms in sem.of_kind(MetricKind.REVENUE, MetricKind.AMOUNT):
        cnt, tot = 0, 0.0
        for row in table.rows:
            v = row.get(ms.name)
            if isinstance(v, (int, float)) and not isinstance(v, bool) and v < 0:
                cnt += 1
                tot += float(v)
        if cnt == 0:
            continue
        share = cnt / max(1, table.row_count)
        sev = Severity.WATCH if (share >= 0.05 or abs(tot) > 0) else Severity.INFO
        out.append(Insight(
            kind=InsightKind.LOSS, severity=sev,
            title=f"{cnt} negative {ms.meaning} records",
            detail=(f"{cnt} records carry a negative {ms.meaning} totalling "
                    f"{fmt_measure(ms.column, tot)} — review refunds, reversals or "
                    f"loss-making items."),
            score=_clamp01(share + 0.2), good=False,
            measure=ms.name, evidence={"count": cnt, "total": tot},
        ))
    return out


# -- Phase 3.1: Duplicate Record Detection ---------------------------------- #
def _duplicates(table: TableProfile, sem: SemanticModel) -> list[Insight]:
    """Scan key columns for duplicate values and report findings."""
    out: list[Insight] = []
    # Identify candidate key columns: identifiers, decoded helpers, text columns.
    key_cols = []
    if table.identifier_column is not None:
        key_cols.append(table.identifier_column)
    # Also look for columns with "ID", "NO", "CODE", "PATIENT", "INVOICE" in name.
    id_kw = ("id", "no", "number", "code", "invoice", "receipt", "claim",
             "encounter", "admission", "patient", "account", "voucher")
    for c in table.columns:
        if c.is_decoded_helper or c.ctype in (ColumnType.IDENTIFIER, ColumnType.DATE):
            continue
        name_lower = c.name.lower()
        if any(kw in name_lower for kw in id_kw):
            key_cols.append(c)
    if not key_cols:
        return out

    for kc in key_cols[:3]:  # limit to first 3 key columns
        seen: dict[str, list[int]] = defaultdict(list)
        for i, row in enumerate(table.rows):
            v = row.get(kc.name)
            if v in (None, ""):
                continue
            seen[str(v)].append(i)
        dups = {k: idxs for k, idxs in seen.items() if len(idxs) > 1}
        if not dups:
            continue
        total_dup_records = sum(len(idxs) for idxs in dups.values())
        unique_dup_keys = len(dups)
        # Estimate potential $ impact if a money measure exists.
        pm = sem.primary_money
        dup_value = 0.0
        if pm is not None:
            for indices in dups.values():
                for idx in indices[1:]:  # skip first occurrence
                    v = table.rows[idx].get(pm.name)
                    if isinstance(v, (int, float)):
                        dup_value += abs(float(v))
        title = (f"{unique_dup_keys} duplicate {kc.name} values "
                 f"({total_dup_records} records)")
        detail = (f"Found {unique_dup_keys} duplicate {kc.name} values "
                  f"affecting {total_dup_records} records"
                  + (f", totalling ~{fmt_measure(pm.column, dup_value)} in {pm.meaning}"
                     if pm and dup_value > 0 else "")
                  + ".")
        out.append(Insight(
            kind=InsightKind.DUPLICATE,
            severity=Severity.WATCH if total_dup_records > 10 else Severity.INFO,
            title=title, detail=detail,
            score=_clamp01(min(1.0, total_dup_records / max(1, table.row_count))),
            good=False,
            measure=kc.name,
            evidence={"count": total_dup_records, "unique_keys": unique_dup_keys,
                      "samples": list(dups.keys())[:5]},
        ))
    return out


# -- Phase 3.2: Date Continuity & Gap Analysis (with known-events integration) #
def _date_gaps(table: TableProfile, sem: SemanticModel,
               date_col: ColumnProfile) -> list[Insight]:
    """Detect gaps in the date series. Gaps overlapping known events (e.g. war
    crisis closures) are explained; unexplained gaps are flagged."""
    series = time_series(table, date_col, date_col)
    if len(series) < 2:
        return []
    # Extract months covered.
    all_periods = sorted(set(p for p, _ in series))
    if len(all_periods) < 2:
        return []
    # Find gaps.
    gaps: list[tuple[str, str]] = []  # (gap_start, gap_end)
    for i in range(len(all_periods) - 1):
        cur = all_periods[i]
        nxt = all_periods[i + 1]
        cur_y, cur_m = int(cur[:4]), int(cur[5:7])
        nxt_y, nxt_m = int(nxt[:4]), int(nxt[5:7])
        # Count months between.
        months_between = (nxt_y - cur_y) * 12 + (nxt_m - cur_m)
        if months_between > 1:  # gap of at least one missing month
            # Fill in gap label.
            gap_mid_y = cur_y + (cur_m + months_between // 2 - 1) // 12
            gap_mid_m = (cur_m + months_between // 2 - 1) % 12 + 1
            gap_label = f"{gap_mid_y:04d}-{gap_mid_m:02d}"
            ev = event_for_period(gap_label) or event_for_period(cur) or event_for_period(nxt)
            gaps.append((cur, nxt))

    if not gaps:
        return []

    total_gaps = len(gaps)
    explained = sum(1 for s, e in gaps
                    if event_for_period(s) or event_for_period(e))
    unexplained = total_gaps - explained
    gap_detail_parts = []
    if explained > 0:
        gap_detail_parts.append(f"{explained} gap(s) match known events "
                                "(war crisis closures — expected).")
    if unexplained > 0:
        gap_detail_parts.append(f"{unexplained} gap(s) are unexplained "
                                "and may indicate missing data.")

    sev = Severity.INFO if unexplained == 0 else Severity.WATCH
    score = _clamp01(unexplained / max(1, total_gaps))
    detail_parts = [f"{total_gaps} gap(s) in the date sequence."]
    detail_parts.extend(gap_detail_parts)
    if gaps:
        detail_parts.append(f"E.g. between {gaps[0][0]} and {gaps[0][1]}.")

    return [Insight(
        kind=InsightKind.DATE_GAP,
        severity=sev,
        title=f"{total_gaps} date gap(s) in the timeline",
        detail=" ".join(detail_parts),
        score=score,
        good=None if explained > 0 else False,
        evidence={"total_gaps": total_gaps, "explained": explained,
                  "unexplained": unexplained, "gaps": gaps[:5]},
    )]


# -- Phase 3.3: Cross-Column Validation ------------------------------------- #
def _validation(table: TableProfile, sem: SemanticModel) -> list[Insight]:
    """Detect logical inconsistencies across columns."""
    out: list[Insight] = []
    issues: list[str] = []
    # Find date columns for date-ordering checks.
    date_cols = table.date_columns
    if len(date_cols) >= 2:
        # Check admission/discharge ordering.
        adms = [c for c in date_cols if any(k in c.name.lower()
                for k in ("admission", "entry", "admit", "arrival", "start"))]
        disch = [c for c in date_cols if any(k in c.name.lower()
                 for k in ("discharge", "exit", "depart", "end", "leave"))]
        if adms and disch:
            inv_count = 0
            for row in table.rows:
                a = row.get(adms[0].name)
                d = row.get(disch[0].name)
                if a and d and isinstance(a, _dt.date) and isinstance(d, _dt.date):
                    if d < a:
                        inv_count += 1
            if inv_count:
                issues.append(f"{inv_count} record(s) have {disch[0].name} "
                              f"before {adms[0].name}.")

    # Check for negative quantities.
    for c in table.measures:
        if any(k in c.name.lower() for k in ("qty", "quantity", "count", "number")):
            neg = 0
            for row in table.rows:
                v = row.get(c.name)
                if isinstance(v, (int, float)) and v < 0:
                    neg += 1
            if neg:
                issues.append(f"{neg} negative {c.name} value(s).")

    # Check for zero-price items.
    for c in table.value_measures:
        if any(k in c.name.lower() for k in ("price", "rate", "fee", "charge", "cost")):
            zero = 0
            for row in table.rows:
                v = row.get(c.name)
                if isinstance(v, (int, float)) and v == 0:
                    zero += 1
            if zero:
                issues.append(f"{zero} zero-value {c.name} record(s).")

    if not issues:
        return out
    return [Insight(
        kind=InsightKind.VALIDATION,
        severity=Severity.WATCH if len(issues) >= 2 else Severity.INFO,
        title=f"{len(issues)} data validation issue(s)",
        detail="; ".join(issues),
        score=_clamp01(len(issues) * 0.15),
        good=False,
        evidence={"issues": issues},
    )]


# -- Phase 3.4: Per-Dimension Outliers -------------------------------------- #
def _per_dim_outliers(table: TableProfile, sem: SemanticModel,
                      date_col: ColumnProfile) -> list[Insight]:
    """Detect outliers within each dimension (e.g. one department's revenue is
    5x its normal level)."""
    out: list[Insight] = []
    pm = sem.primary_money
    if pm is None or not table.date_columns:
        return out

    # Find readable dimensions.
    dims = [c for c in table.columns
            if (c.ctype == ColumnType.CATEGORICAL or c.is_decoded_helper)
            and not c.is_decoded_helper
            and c.distinct >= 3 and c.distinct <= 100]
    if not dims:
        return out

    for dim in dims[:2]:
        # For each dimension value, compute per-period average and detect spikes.
        dim_vals: dict[str, list[float]] = defaultdict(list)
        for row in table.rows:
            period = _period_key(row.get(date_col.name))
            val = row.get(pm.name)
            dim_key = row.get(dim.name)
            if period and isinstance(val, (int, float)) and dim_key not in (None, ""):
                dim_vals[str(dim_key)].append(float(val))

        for dval, vals in dim_vals.items():
            if len(vals) < 4:
                continue
            med = _median(vals)
            mad = _mad(vals, med)
            if mad == 0:
                continue
            # Check latest period value.
            latest_val = vals[-1]
            z = 0.6745 * (latest_val - med) / mad
            if abs(z) >= _ANOM_WATCH:
                direction = "above" if z > 0 else "below"
                from ..decode import friendly_name
                out.append(Insight(
                    kind=InsightKind.PER_DIM_OUTLIER,
                    severity=Severity.WATCH,
                    title=f"'{dval}' in {dim.name} is unusual this period",
                    detail=(f"'{dval}' ({friendly_name(dim.name)}) has "
                            f"{pm.meaning} of {fmt_measure(pm.column, latest_val)}, "
                            f"{abs(z):.1f}σ {direction} its typical "
                            f"{fmt_measure(pm.column, med)}."),
                    score=_clamp01(abs(z) / 5.0), good=None,
                    measure=pm.name, dimension=dim.name,
                    evidence={"z": z, "median": med, "value": latest_val,
                              "dim_value": dval},
                ))
                if len(out) >= 2:  # max 2 per-dim outliers
                    break
    return out


# -- Phase 4.3: Quarter-over-Quarter (QoQ) Variance ------------------------- #
def _qoq_variance(table: TableProfile, sem: SemanticModel,
                  date_col: ColumnProfile) -> list[Insight]:
    """Quarter-over-quarter change for each money measure, with drivers."""
    out: list[Insight] = []
    for ms in sem.of_kind(MetricKind.REVENUE, MetricKind.COST,
                          MetricKind.AMOUNT, MetricKind.BALANCE):
        # Use quarter-level aggregation.
        q_series = _q_series(table, date_col, ms.column)
        if len(q_series) < 2:
            continue
        (pp, prev), (lp, last) = q_series[-2], q_series[-1]
        change = last - prev
        frac = _pct(change, prev)
        if frac is None or abs(frac) < _VAR_WATCH:
            continue
        sev = (Severity.HIGH if abs(frac) >= _VAR_HIGH else Severity.WATCH)
        up = change >= 0
        hib = ms.kind.higher_is_better
        good = None if hib is None else (up == hib)
        arrow = "rose" if up else "fell"
        detail = (f"{ms.meaning} {arrow} {abs(frac) * 100:.0f}% "
                  f"({fmt_measure(ms.column, abs(change))}) in {lp} vs {pp}.")
        score = _clamp01(abs(frac))
        out.append(Insight(
            kind=InsightKind.QOQ_VARIANCE, severity=sev,
            title=f"{ms.meaning} {arrow} {abs(frac) * 100:.0f}% QoQ",
            detail=detail, score=score, good=good,
            measure=ms.name, period=lp,
            evidence={"prev": prev, "last": last, "prev_period": pp},
        ))
    return out


def _q_series(table, date_col, measure) -> list[tuple[str, float]]:
    """Sum measure by quarter, sorted chronologically."""
    from ..aggregate import _quarter_key  # noqa: PLC0415
    totals: dict[str, float] = defaultdict(float)
    for row in table.rows:
        period = _quarter_key(row.get(date_col.name))
        val = None
        v = row.get(measure.name)
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            val = float(v)
        if period is None or val is None:
            continue
        totals[period] += val
    return sorted(totals.items(), key=lambda kv: kv[0])


# --------------------------------------------------------------------------- #
# Shared helpers                                                              #
# --------------------------------------------------------------------------- #
def _dimensions(table: TableProfile) -> list[ColumnProfile]:
    """Readable grouping columns: decoded helpers first, then plain categoricals.
    The raw CODE column (the one whose decoded helper exists) is excluded so we
    never group/label by cryptic codes when a decoded name is available."""
    helpers = [c for c in table.columns if c.is_decoded_helper]
    seen = {c.name for c in helpers}
    plain = [c for c in table.pivot_dimensions
             if not c.is_decoded_helper and not c.decoded_helper
             and c.name not in seen]
    return helpers + plain


def _label(name: str) -> str:
    """Reader-friendly header label (decodes helper/abbreviation via the library)."""
    try:
        from ..decode import friendly_name  # noqa: PLC0415 (avoid import cycle)
        return friendly_name(name)
    except Exception:
        return name


def _as_date(v: Any) -> Optional[_dt.date]:
    if isinstance(v, _dt.datetime):
        return v.date()
    if isinstance(v, _dt.date):
        return v
    return None


def _latest_date(table: TableProfile, date_col: ColumnProfile) -> Optional[_dt.date]:
    latest: Optional[_dt.date] = None
    for row in table.rows:
        d = _as_date(row.get(date_col.name))
        if d and (latest is None or d > latest):
            latest = d
    return latest


def _age_days(v: Any, latest: _dt.date) -> Optional[int]:
    d = _as_date(v)
    return (latest - d).days if d else None


# --------------------------------------------------------------------------- #
# Public entry point                                                          #
# --------------------------------------------------------------------------- #
def detect_insights(profile: WorkbookProfile, semantic: SemanticModel,
                    max_insights: int = 18) -> list[Insight]:
    """Run every detector and return findings ranked most-important first.
    Now includes Phase 3 (data quality) and Phase 4 (QoQ) detectors."""
    table = profile.primary
    if table is None or table.row_count == 0 or not semantic.measures:
        return []

    found: list[Insight] = []
    date_col = table.date_columns[0] if table.date_columns else None
    try:
        if date_col is not None:
            found += _variance(table, semantic, date_col)
            found += _anomaly(table, semantic, date_col)
            found += _trend(table, semantic, date_col)
            found += _aging(table, semantic, date_col)
        found += _concentration(table, semantic)
        found += _losses(table, semantic)
    except Exception:                      # never let one detector sink the run
        pass

    # Phase 4: QoQ variance detector.
    try:
        if date_col is not None:
            found += _qoq_variance(table, semantic, date_col)
    except Exception:
        pass

    # Phase 3: Data Quality detectors (isolated so one failure can't sink the run).
    try:
        found += _duplicates(table, semantic)
    except Exception:
        pass
    try:
        if date_col is not None:
            found += _date_gaps(table, semantic, date_col)
    except Exception:
        pass
    try:
        found += _validation(table, semantic)
    except Exception:
        pass
    try:
        if date_col is not None:
            found += _per_dim_outliers(table, semantic, date_col)
    except Exception:
        pass

    # De-duplicate by (kind, measure, period); rank by severity then score.
    seen: set[tuple] = set()
    unique: list[Insight] = []
    for ins in found:
        key = (ins.kind, ins.measure, ins.period, ins.dimension)
        if key in seen:
            continue
        seen.add(key)
        unique.append(ins)
    unique.sort(key=lambda i: i.sort_key(), reverse=True)
    return unique[:max_insights]
