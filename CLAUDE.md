# CLAUDE.md — Excel Intelligence Agent

Operating guide for any AI agent (or developer) working on this project. Read this
fully before changing code, building, or publishing. It captures the architecture,
the build/packaging/update pipeline, and the non-obvious gotchas that took many
iterations to get right.

---

## 1. What this is

A **standalone Windows desktop app** (PySide6) for non-technical business users at
**Sahel General Hospital**. The user picks an Excel workbook, clicks a button, and the
app appends analysis sheets — **Dashboard, Pivot Analysis, KPI Analysis, Executive
Summary, Smart Tables** — to the *same* workbook, leaving the original data sheet untouched.

- Distributed as a per-user **installer**: `ExcelIntelligenceAgent-Setup.exe`.
- Requires **Microsoft Excel installed** for the full feature set (real PivotTables are
  built by automating Excel via COM). Without Excel it falls back to static tables.
- The user must never see Python, a terminal, PowerShell, or a config file.

Current version is in `config.py` → `APP_VERSION` (currently `1.14.0`).

### v1.14.0 — every sheet enhanced (GM-grade), big-file performance fix
- **Shared PivotCache (`excel_com._build_plan`)**: ONE cache for all pivots instead
  of one-per-pivot. An 80k-row book went from ~46 MB + very slow (and pivots left
  EMPTY when per-pivot caches exhausted resources) to a few MB and fast, with all
  pivots populated. This was the "dashboard takes too long / pivots without values"
  root cause.
- **Date in the COLUMNS panel (`excel_com._build_one_pivot`, `XL_COLUMN_FIELD`)**:
  the grouped Year/Month is placed as a COLUMN field (months spread left-to-right
  beside the row titles) whenever the pivot has a row dimension; a date-only pivot
  keeps the date in rows. Grouped once on the shared cache (`Years` field reused).
- **KPI sheet broadened (`analyzers/kpi.py`)**: a full scorecard (total, avg/month,
  best/lowest month, MoM, YoY, active months, top dimension) + a month-by-month
  trend table (Δ%/share/cumulative) with a trend chart + Top-N decoded breakdowns
  with % of total + a currency split. Emitted with or without Excel.
- **Dashboard never blank (`analyzers/dashboard.py`)**: a real one-page dashboard
  (tiles + up to 5 openpyxl charts: monthly trend, by-dimension bar, composition
  pie, Pareto, year-over-year), decoded names. The COM `_build_dashboard_charts`
  rebuild (which used to DELETE these and leave the sheet blank when pivots failed)
  is no longer called.
- **Smart Tables months-across (`aggregate.crosstab_period`)**: each decoded
  dimension is a MONTHS-ACROSS cross-tab (names down the rows, months across the
  columns, a Total column), Top-N + "Others" rollup, month heat color-scale.

### v1.13.0 — account categories, purpose detection, decoded display, revenue sign
Driven by the hospital's chart-of-accounts catalog (code / description / category):
- **`CodeMap.categories`** (`core/library/store.py`): each account code now carries
  a category (revenues / purchases / salaries / cash / …). `tools/ingest_account_categories.py`
  parses the wide "code+description under a category banner" catalog into the
  `account` map (descriptions + categories); `category_of(code)` resolves it.
- **Purpose detection** (`core/semantic.py` `_detect_purpose`): the engine sums the
  decoded account categories (money-weighted) to infer what the workbook is ABOUT
  — REVENUE vs EXPENSES — sets `SemanticModel.purpose`/`purpose_kind`, re-tags the
  money measures to that nature, and the Insights sheet states it ("this looks
  like a REVENUE report"). Applies to Auto and Custom.
- **Decoded descriptions everywhere** (`core/decode.py` `friendly_name`): all sheets
  show the glossary meaning of a header (`ACTTNUMB` → "Department Account Name")
  and decode codes to names; the raw CODE column is dropped from groupings
  (`is_decoded_helper`/`decoded_helper` aware) in insights, smart tables, exec.
- **Revenue sign reversal** (`pipeline._apply_revenue_sign`): Lebanese revenue books
  store amounts NEGATIVE. For a revenue purpose, LBP money columns are flipped to
  POSITIVE in-memory (openpyxl sheets) and via a HIDDEN positive helper column
  `"<col> (+)"` that the COM pivots aggregate (so SUM/AVG/MIN/MAX are correct);
  originals are untouched. USD/$ columns are detected and left positive.
- **USD detection**: if the sheet already has a USD/$ value column, the dollar
  prompt is skipped and no extra dollar column/calc is added (`pipeline` + UI).
- **Helpers stay hidden** (`excel_com._hide_helper_columns`): COM re-hides every
  `… (Name)` / `… (+)` helper after re-saving, fixing a visible-helper bug.

### v1.12.1 — generation-hang fixes (frozen progress bar)
Two causes of a "stuck loading bar" during Auto-Generate were fixed:
- **Excel automation now uses `win32com.DispatchEx`** (a SEPARATE, dedicated Excel
  process) instead of `Dispatch`. Plain `Dispatch` attaches to the user's
  ALREADY-OPEN Excel; `Visible=False` then hides their window and any modal
  dialog they (or we) trigger blocks forever — a frozen bar. The isolated
  instance can't be disturbed by, or disturb, the user's interactive Excel.
  `Workbooks.Open` also passes `UpdateLinks=0, IgnoreReadOnlyRecommended=True`
  and we set `AskToUpdateLinks/EnableEvents=False` + `AutomationSecurity=3` so no
  link-update / macro / read-only prompt can stall the run.
- **Groq fast-fail.** `GroqNarrator` now uses a short separate CONNECT timeout
  (`connect_timeout=6s`, read `20s`, `retries=0`) via a `(connect, read)` tuple.
  A hospital firewall that silently DROPS packets to `api.groq.com` previously
  stalled the "Building Executive Summary…" stage for up to ~60s; it now falls
  back to the deterministic summary in ~6s.

### v1.12.0 — the intelligence layer (semantic read + insight engine + Insights sheet)
The big upgrade from "aggregates everything" to "reasons about what matters".
Three new pure-`core/` layers feed a new headline sheet:

- **Semantic layer (`core/semantic.py`).** Adds *meaning* on top of the
  profiler's *shape*: `classify_metric` tags each money/numeric/percent column as
  `MetricKind` (REVENUE / COST / BALANCE / VOLUME / RATIO / AMOUNT / GENERIC),
  and `analyze()` builds a `SemanticModel` (with `primary_money`, `revenue`,
  `cost`, `balance` accessors) and detects a `ReportType`
  (FINANCIAL / RECEIVABLES / CENSUS / OPERATIONS / GENERIC). Keyword banks match
  BOTH the raw header and its library meaning, so it works decoded or not.
  Everything falls back to GENERIC — never errors.
- **Insight engine (`core/insights/`).** `detect_insights(profile, semantic)`
  returns ranked, typed `Insight` objects from explainable statistics — NO
  training data, runs offline on one workbook: period-over-period **variance**
  (with the top driver), Pareto **concentration**, MAD/z-score **anomaly**,
  least-squares **trend + one-step forecast**, receivables **ageing** buckets
  (0–30/31–60/61–90/90+), and negative-record **losses**. Each carries a
  `Severity` (HIGH/WATCH/INFO), a 0–1 `score`, `good` (RAG), and `evidence`
  (driver, buckets, items, series) used to build charts. Detectors are wrapped so
  one failure can't sink the run; findings are de-duped and severity-ranked.
- **Insights sheet (`core/analyzers/insights.py`, `SHEET_INSIGHTS`).** The new
  FIRST tab (`writer._move_insights_first` moves it to index 0). Renders a RAG
  **KPI scorecard**, a **Bottom Line**, ranked **"What to look at"**, amber
  **Risks**, red **Recommended actions** (derived from the insight kinds), a
  **Findings** Smart Table (priority · finding · measure · impact, with a data
  bar on impact), plus a **Pareto** chart (bars + cumulative-% line on a
  secondary axis) and a **trend + forecast** line. Always produces something —
  shows "stable, keep monitoring" when nothing material is found.
- **Smart Tables 2.0 rendering (`render.DataTable.bar_columns` /
  `scale_columns`).** `writer._apply_table_visuals` adds openpyxl **data bars**
  and green→red **color-scale heatmaps** to chosen columns. Smart Tables put a
  bar on each value column; the Insights findings table bars the impact score.
- **Charts (`render.ChartKind.PARETO` / `COMBO`, `ChartSpec.line_values`).** The
  writer renders a true Pareto/combo as a `BarChart` with a `LineChart` overlay
  on an independent secondary axis. All chart work is **openpyxl** so it renders
  WITH OR WITHOUT Excel (matches the product's offline-fallback philosophy). A
  native Excel **waterfall** is the one deferred COM-only enhancement.
- **Wiring.** `AnalysisOptions.insights` (default True), `_all_analyzers` /
  `_selected_analyzers` put `InsightsAnalyzer` first, the pipeline captures
  `AnalysisResult.insights`, and the UI gains an **Insights** checkbox + shows the
  top finding in the completion dialog. Tests: `tests/test_insights.py`
  (semantic classification, variance/concentration detection, ranking, empties).

### v1.11.0 — library everywhere, LBP default, date-grouped scenarios
- **Library on every sheet (not just Smart Tables).** After profiling, the engine
  detects code columns the library can decode and injects HIDDEN decoded-name
  helper columns (`<CODE> (Name)`) onto the data sheet (`core/decode.py` +
  `writer.inject_hidden_helpers`). Originals are untouched; the result `notes`
  list the hidden columns. Code columns are then grouped BY the helper, so the
  real PivotTables, Dashboard, KPI, Summary and Smart Tables all show real names.
  `ColumnProfile.decoded_helper` / `is_decoded_helper` carry this through the
  model; `TableProfile.dimensions`/`pivot_dimensions` skip the raw-code column.
- **LBP is the currency default.** `formatting.is_dollar_column` (shared with
  `pivot_plan`) → money is `… LBP` unless the column is explicitly dollars
  (`USD`/`$`). Fixes KPI/Dashboard tiles that previously showed `$` on LBP.
  `render.NumberFormat.LBP` (`#,##0" LBP"`) added for tables.
- **BASIC RULE: never a pivot/smart table without a date.** `pivot_plan` nests the
  date (Month/Year) as the OUTER row field on every category pivot; Smart Tables
  and the static Pivot fallback group by month (`aggregate.group_period_dim`).
- **Smart Tables is a scenario generator** (`analyzers/smart_tables.py`): many
  tables (≤14) — each readable dimension × value measure, month-grouped, decoded,
  in LBP. Appears only when a helper was actually injected.
- **Custom wizard shows the library description** per title
  (`Engine.describe_columns` adds `description`; `ui/custom_dialog` renders it).
- **Output Mode adds a "Smart Tables" checkbox** (`ui/main_window`).

---

## 2. Golden rules

1. **`core/` must NEVER import PySide6 (or any UI).** The engine is UI-agnostic. The UI
   (`ui/`) depends on `core/`, never the reverse. Every front-end calls one entry point:
   `core.pipeline.Engine().run(path, AnalysisOptions, progress_cb) -> AnalysisResult`.
2. **Never use `ws.cell(row, col)` random access on a read-only openpyxl worksheet.** It
   re-streams the sheet from the top each call → O(rows²). Always iterate with
   `ws.iter_rows()` (see `core/loader.py`). A 500-row file once took >120 s because of this.
3. **Secrets (`local_secrets.py`, `buildinfo.py`) are gitignored.** Never commit a key.
4. **Don't break the four output sheet names** (`core/constants.py`): the loader skips them
   on re-runs, and re-runs regenerate them cleanly.
5. After changing engine logic, run the test suite: `python -m pytest tests/test_engine.py -q`
   (6 tests; ~30–40 s when Excel is responsive, longer if Excel is busy).

---

## 3. Repository layout

```
the_excel_agent/
├── main.py                  # launcher -> ui.app.run()
├── config.py                # APP_NAME, ORG_NAME, APP_VERSION, Groq + update config, bundled key resolver
├── buildinfo.py             # GENERATED + GITIGNORED: BUILD_DATE + BUNDLED_GROQ_KEY
├── local_secrets.py         # GITIGNORED: GROQ_API_KEY source for dev/build
├── local_secrets.example.py # committed template (no key)
├── build.spec               # PyInstaller ONEDIR spec
├── installer.iss            # Inno Setup script -> Setup.exe
├── build.bat                # one-click: build + package + publish
├── requirements.txt
├── .gitignore
├── core/                    # ENGINE (no UI imports)
│   ├── pipeline.py          # Engine.run(), describe_columns(), analyzer wiring
│   ├── loader.py            # open workbook, single-pass detect tables/headers/formats
│   ├── profiler.py          # column type inference (numeric/currency/percent/date/categorical/text/IDENTIFIER)
│   ├── pivot_detect.py      # detect sheets that already contain a PivotTable (zip/xml, no Excel)
│   ├── models.py            # dataclasses: ColumnProfile, TableProfile, WorkbookProfile, AnalysisOptions, CustomSelection, ...
│   ├── pivot_plan.py        # declarative plan of PivotTables (auto + custom + combinations)
│   ├── excel_com.py         # ExcelFinalizer: drives Excel via win32com (pivots, CF, sort, charts, save)
│   ├── writer.py            # openpyxl renderer for KPI/Dashboard/Exec sheets (+ static fallback)
│   ├── render.py            # framework-free SheetSpec/DataTable/ChartSpec/KpiTile/TextBlock
│   ├── aggregate.py         # group_sum, time_series, period_over_period_growth
│   ├── formatting.py        # human value formatting
│   ├── constants.py         # output sheet names
│   ├── updater.py           # GitHub-release update check + silent installer launch
│   ├── llm/                 # Groq narrative layer (optional)
│   │   ├── groq_client.py   # GroqNarrator (OpenAI-compatible, JSON mode, soft-fail)
│   │   └── prompts.py       # consultant-grade system + JSON-schema user prompt
│   └── analyzers/           # base.py, kpi.py, pivot.py, dashboard.py, executive_summary.py
├── ui/                      # PySide6 front-end
│   ├── app.py               # QApplication bootstrap, stylesheet, window icon
│   ├── main_window.py       # main screen, mode buttons, update check, closeEvent (silent update)
│   ├── custom_dialog.py     # Custom Generate wizard (singles + combination builder)
│   ├── settings_dialog.py   # Groq key + model (Credential Manager), warning label
│   ├── update_dialog.py     # AutoUpdateDialog (kept; current flow is silent, see §9)
│   ├── update_worker.py     # threaded update check/download
│   ├── worker.py            # threaded Engine.run wrapper
│   ├── settings_store.py    # keyring-backed user key/model; effective_key()
│   └── resources/           # style.qss, app.ico
├── tools/
│   ├── stamp_build.py       # writes buildinfo.py (date + bundled key)
│   └── make_icon.py         # generates ui/resources/app.ico (Pillow)
└── tests/
    ├── test_engine.py       # 6 engine tests
    └── make_sample.py       # synthetic fixture generator
```

---

## 4. The engine pipeline (`core/pipeline.py`)

`Engine.run(path, options, progress_cb)`:
1. `load_workbook_profile(path)` → `WorkbookProfile` (all data tables, column roles, number
   formats, and any sheets that already contain pivots — those are left untouched).
2. If `options.custom` is set, target the chosen sheet and record preferred measures.
3. Build the selected analyzers' `SheetSpec`s and write KPI/Dashboard/Executive Summary via
   **openpyxl** (`writer.py`). The Pivot Analysis sheet is NOT written by openpyxl when Excel
   is available — COM builds real pivots there.
4. If `excel_available()`: build the **pivot plan** (`pivot_plan.build_pivot_plan(profile, custom)`)
   and run `ExcelFinalizer.finalize(path, profile, plan)` (COM). Else: note that Excel is absent
   and leave static tables.
5. Return `AnalysisResult` (sheets created, notes, whether the LLM was used).

`Engine.describe_columns(profile, sheet)` powers the Custom wizard (lists groupable
dimensions and value measures with a recommended pre-selection).

---

## 5. Column classification (`core/profiler.py`, `core/models.py`)

Detection is **data-driven**, not keyword-bound, so it works on any headers/language:

- **Number format is the primary signal:** a cell format containing `$ € £ ¥ ₹ [$` →
  `CURRENCY`; containing `%` → `PERCENT`.
- **Type + cardinality:** dates by value type; low-distinct text → `CATEGORICAL`; high-distinct
  text → `TEXT`.
- **IDENTIFIER:** integer columns that are row ids / codes (header like `id/no/#/index/code`,
  the first column of unique ints, or any near-unique integer ratio > 0.98) → **excluded from
  measures and dimensions** so the agent never totals a row number.
- **Header keywords** are only a fallback hint for currency/percent when the format is `General`.
- `TableProfile` exposes roles: `value_measures` (money/number, ranked by `value_score` so PNL
  beats price), `percent_measures`, `primary_value_measure`, `key_measures`, `dimensions`
  (categorical), `pivot_dimensions` (categorical + moderate/high-card text, excluding single-value
  and near-unique), `date_columns`, `identifier_column`.

`pivot_detect.detect_pivot_sheets(path)` parses the .xlsx zip rels to find sheets that already
contain a PivotTable (no Excel/openpyxl needed); those are reported on the Dashboard and never
modified.

---

## 6. Pivot plan (`core/pivot_plan.py`)

A declarative list of `PivotSpec`s; `excel_com` renders them. Two modes:

### Auto (`build_pivot_plan(profile)`)
- Per date column (ENTRY AT / EXIT AT…): grouped Month+Year, **Record Count + Total value ($)**.
- Per categorical/groupable dimension: **Total value ($)** and **Total percent (%)**.
- Date × dimension cross-breakdowns (date outer, dim inner).
- **Combined pivots:** date period × pairs of top dimensions, with **Sum + % of total**.
- KPI sheet: **Measure Statistics** pivot (Sum/Avg/Count/Min/Max) on the primary value.
- Wide dimensions (distinct > 25, e.g. TRIGGER DETAIL) are limited to **Top-N (20) by value**.

### Custom (`build_custom_plan(table, CustomSelection)`)
- **Singles** (always): one pivot per selected title × each chosen measure.
- **Combinations** (`CustomSelection.combinations: list[list[str]]`): each list nests its titles
  into ONE pivot (date outer if present, then categories) with Sum of each value + % of total.
- Measure display format chosen by the user per `MeasureChoice.format_kind`:
  `usd` → `"$"#,##0.00`, `lbp` → `#,##0" LBP"`, `number` → `#,##0.00`, `percent` →
  `0.00%`/`0.00"%"`, `auto` → inherit source.

### Number-format rules
- Currency: inherit source `$`/`€`/`£`/`[$` format, else default `"$"#,##0.00`.
- Percent: if the source already uses an Excel `%` format the value is a fraction → `0.00%`;
  if it's a percent-by-header on a `General` column the value is already a percent → literal
  `0.00"%"` (do NOT multiply by 100). See PNL PCT.
- `% of total` data field: `DataFieldSpec.calculation = XL_PERCENT_OF_TOTAL (8)`, format `0.00%`.

---

## 7. Excel COM finalizer (`core/excel_com.py`) — the tricky part

`ExcelFinalizer.finalize(path, profile, plan)` drives a hidden Excel instance via `win32com`.
**Excel COM needs an ABSOLUTE path** for `Workbooks.Open`. Order of operations matters a lot
because **date grouping and refresh are PivotCache-level operations that refresh ALL pivots
sharing the cache**, which silently wipes inline sort/conditional-formatting.

Sequence:
1. Convert the source data range to a real Excel **Table (ListObject)** if it isn't one already
   (detect first; never re-convert).
2. **Build every pivot first** (create, add row fields with the grouped DATE outermost, group the
   date by Month+Year via `.Group(Periods=(F,F,F,F,True,F,True))`, add data fields + number
   formats + `% of total` calculation, set grand totals). Collect `(pivot, spec)`.
3. **THEN, in separate sub-passes over all pivots (so nothing wipes earlier work):**
   - `_disable_subtotals` (subtotals off → each data field's `DataRange` is pure detail cells →
     this is how "exclude subtotals" is satisfied).
   - **Sort:** `_sort_and_limit` uses `PivotField.AutoSort(2, dataFieldName)` (xlDescending). Use
     AutoSort — it's a field-definition property that survives cache refresh; manual
     `PivotItem.Position` gets reset. For wide dims, hide non-Top-N items first
     (`pf.PivotItems()` — **must be CALLED with parentheses**).
   - **Conditional formatting (`_cf_pivot`):** for each data field, apply a plain
     `DataRange.FormatConditions.AddTop10()` rank 1 highlight. Apply to the **full DataRange**
     (NOT `Cells(1,1)`, NOT `ScopeType` — both fail/are unreliable). DataRange excludes grand
     totals; subtotals are off → highlights exactly the top value per column.
4. Build **Dashboard charts** (`_build_dashboard_charts`): delete the old static chart, then add a
   column chart per single-dimension value (`$`) category pivot (skip percent, skip wide >25
   items), laid out 2-per-row, `SetSourceData(pt.TableRange1)`.
5. Conditional-format any static ListObjects (skip the raw source table and pivot-containing sheets).
6. **AutoFit** all columns on all produced sheets.
7. `wb.Save()` (in place, same path/name), close, quit, `CoUninitialize`.

**Critical: `RefreshOnFileOpen` is deliberately OFF.** Refresh-on-open re-runs the pivot and wipes
the sort order and conditional formatting. Pivots are fully populated at creation, so the saved
file looks correct immediately. (Trade-off: if the user later changes source data they refresh
manually. This was a conscious decision to keep CF/sorting persistent.)

**Inspecting CF in tests:** read it via `DataField.DataRange.FormatConditions`, NOT `.Cells(1,1)`.

The branding "SAHEL GENERAL HOSPITAL" lives in the **app window** (`config.ORG_NAME`), NOT in the
Excel Dashboard (whose heading is the generic "Executive Dashboard").

---

## 8. Groq LLM (Executive Summary) & key handling

- `core/llm/groq_client.GroqNarrator`: OpenAI-compatible call to Groq, `response_format` JSON
  object, model default `config.DEFAULT_GROQ_MODEL` (`llama-3.3-70b-versatile`). Soft-fails to
  None on any error (no network, 401, 429) so the summary is ALWAYS produced.
- `executive_summary.py` builds rich metrics (leaders/laggards, concentration with a positive-mass
  base so PnL nets≈0 don't blow up the %, trend best/worst, negatives), asks Groq for a structured
  JSON briefing, and renders styled sections (Bottom Line / Overview / Key Findings / Risks (amber)
  / Opportunities / **Recommended Action Plan (red)**). Deterministic fallback builds the same
  structure offline.

### Bundled Groq key — resolution order (`config.bundled_groq_key()` → `_raw_bundled_key()`)
1. `buildinfo.BUNDLED_GROQ_KEY` — **stamped at build time by `tools/stamp_build.py`**. `config`
   imports `buildinfo` UNCONDITIONALLY, which guarantees PyInstaller bundles it into the exe's PYZ.
2. `local_secrets.GROQ_API_KEY` — gitignored, used for dev runs from source.
3. `GROQ_API_KEY` environment variable.
Accepts a raw `gsk_...` key or a base64 blob. **Per-user keys** entered in Settings (stored in
Windows Credential Manager via `keyring`, `ui/settings_store.py`) take precedence at runtime.

> GOTCHA: a *conditionally* imported module (like `local_secrets`) may not be statically detected
> by PyInstaller. `buildinfo` is the reliable carrier because `config` imports it at top level.
> To verify a key is in a build: extract the exe's `PYZ.pyz` with PyInstaller's
> `CArchiveReader`/`ZlibArchiveReader` and check for `buildinfo` — do NOT just `ls _internal`
> (modules live compressed inside the PYZ, not as loose files).

> SECURITY: the bundled key is extractable from any distributed exe. It's a free-tier Groq key; if
> it leaks/abused, regenerate at console.groq.com and rebuild. The repo is public, so the published
> installer exposes the key — accepted by the project owner.

A diagnostic for "AI not working" on a colleague's PC: **Settings → Test Connection** shows the real
error (network/firewall block of `api.groq.com`, invalid key, or rate limit). A working key returns
"Connection successful."

---

## 8b. Reference library ("the brain") & Smart Tables

A persistent reference store that teaches the engine the hospital's own vocabulary so
the **Smart Tables** sheet (and, later, Executive Summary / KPI) produce hospital-relevant,
decoded output. It is plain Python + JSON — **no UI imports, part of `core/`.**

Layout (`core/library/`):
- `store.py` — `Library` dataclass + `HeaderEntry` / `CodeMap`; tolerant lookup, `load/save`,
  `get_library()` (process-cached). `data/*.json` holds the knowledge.
- `ingest.py` — `ingest_excel(path, kind?, category?)` merges one reference workbook into the
  library (incremental, idempotent). Auto-detects sheet kind; forgiving column heuristics.
- `data/headers.json` — abbreviation glossary `{ABBREV: {meaning, category}}`.
- `data/codes.json` — per-domain code maps `{domain: {label, entries:{code: definition}}}`.
- `data/meta.json` — version + list of ingested sources.

Two knowledge kinds:
1. **Header glossary** — abbreviated header → real meaning (`ADTH → Date of Admission`),
   optionally a `category` linking the column to a code domain.
2. **Code maps** — one domain each (guarantor, department, supplier, doctor, …):
   `code → definition` (`001 → Private patient`).

**Tolerant matching** (`store.py`): codes normalize so `1`, `001`, `"001 "` all resolve
(Excel often reads a code column as ints). A data column is matched to a code map either by
its glossary `category` or by **value-overlap auto-detection** (`best_map_for_values`, ≥50%
coverage). Headers match by normalized form, then alphanumeric-squashed fallback.

**Smart Tables analyzer** (`core/analyzers/smart_tables.py`): emits plain openpyxl DataTables
(NOT pivots, no COM), one "Total \<measure\> by \<decoded dimension\>" per decodable column.
`applies_to()` returns **False while the library is empty**, so the sheet only appears once
reference files are ingested — nothing breaks before then. Wired into `pipeline.py`
(`_all_analyzers` / `_selected_analyzers`) and `AnalysisOptions.smart_tables` (default True).

**Ingesting new files** (run per file the user sends; updates the committed JSON, which the
next build ships to colleagues):
```
python tools/ingest_library.py "Headers Glossary.xlsx"                # auto-detected
python tools/ingest_library.py "Guarantor Codes.xlsx" --kind codes --category guarantor
python tools/ingest_library.py --show                                 # library stats
```
The JSON files are bundled via `build.spec` `datas`. Tests: `tests/test_library.py`.
Column-name heuristics in `ingest.py` (`_*_HINTS`) may need tuning as real files arrive.

**Populated library (v1.10.0).** The SAHEL reference set is now ingested (`data/*.json`):
- **Header glossary:** 183 headers (`ACTTNUMB → Department Account Name`, `SRCC → Source Code`, …).
- **`account` code map:** 2,687 chart-of-accounts codes (`101300100001 → Subscribed Capital (Called & Paid Up)`).
- **`fld1` code map:** 961 entity/payer codes (`10001 → NATIONAL SOCIAL SECURITY`, `10003 → LEBANESE ARMY`, `30047 → BAHMAN HOSPITAL`).

Ingested with explicit flags because both code files carry a banner header row
(`Header : ACTTNUMB` / `HEADER : FLD1`):
```
python tools/ingest_library.py "HEADER library.xlsx" --kind headers
python tools/ingest_library.py "HEADER subtabs account ....xlsx" --kind codes --category account
python tools/ingest_library.py "HEADER subtabs ... FLD1 ....xlsx" --kind codes --category fld1
```

**Ingest gotcha fixed (`_ingest_codes_sheet`).** The FLD1 file's definition column is titled
"Description of **codes** of this header"; the substring "codes" matched `_CODE_HINTS` and stole
the code-column pick, collapsing code+definition onto one column (entries became `name → name`).
A collision guard now falls back to first=code / second=definition when the two picks coincide.

---

## 9. How updates are handled (the colleague experience is INVISIBLE)

Updates flow through **GitHub Releases** of `config.GITHUB_REPO`
(`alamehmazen123/the_excel_agent`). `config.UPDATE_MANIFEST_URL` is that repo's
**latest-release API** (`https://api.github.com/repos/<repo>/releases/latest`).

Runtime flow (`ui/main_window.py` + `core/updater.py`):
1. On launch, a background thread calls `updater.check_for_update(UPDATE_MANIFEST_URL, APP_VERSION)`.
   For a GitHub URL it parses `tag_name` (strip leading `v`) as the version and finds the `.exe`
   asset's `browser_download_url`. Returns `UpdateInfo` only if the release is newer; soft-fails to
   None (404 = no releases yet).
2. If newer, the app **downloads the installer QUIETLY in the background** (no dialog, no progress
   popup) and stores `_pending_installer`. A subtle footer line says "An update will be applied
   automatically."
3. **On app close (`closeEvent`)**, `updater.launch_installer_silent(path)` runs the installer
   detached with `/VERYSILENT /SUPPRESSMSGBOXES /CLOSEAPPLICATIONS /NORESTART`. It installs in the
   background; **no relaunch** (the user reopens later to the new version). The colleague sees no
   "updating…" UI at all.

The installer's `[Run]` launch entry has `skipifsilent`, so **silent updates don't relaunch** while
**fresh interactive installs do** launch the app at the end.

Self-update only runs in the frozen/installed app (`updater.is_frozen()`); it's a no-op from source.

`update_dialog.AutoUpdateDialog` (a visible "updating now, please wait…" variant) still exists but
is NOT wired in — the current flow is the silent close-time install above.

---

## 10. How the Setup.exe is built and handled

The deliverable is **`Output\ExcelIntelligenceAgent-Setup.exe`** — a per-user installer (no admin),
NOT a bare single-file exe.

> Why not single-file? A single-file PyInstaller exe glues a compressed archive onto the launcher;
> if that file is corrupted in transit or stripped by antivirus it fails with *"Could not load
> PyInstaller's embedded PKG archive."* The **onedir + installer** format has no such appended
> archive, is antivirus-friendly, and installs cleanly.

Pipeline (each step also runs standalone; `build.bat` chains them):
1. `python tools/make_icon.py` → `ui/resources/app.ico` (Pillow; Excel-green grid + AI spark).
2. `python tools/stamp_build.py` → writes `buildinfo.py` with today's `BUILD_DATE` and the resolved
   `BUNDLED_GROQ_KEY`.
3. `python -m PyInstaller build.spec --noconfirm --clean` → **onedir** folder
   `dist\ExcelIntelligenceAgent\` (exe + `_internal\`). `build.spec` is `exclude_binaries=True` +
   `COLLECT`, UPX off (AV-friendly), bundles `style.qss` + `app.ico`, hidden-imports for
   keyring/pywin32 and (best-effort) `local_secrets`.
4. Inno Setup compiles `installer.iss` →
   `"%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" /DAppVersion=<APP_VERSION> installer.iss`.
   Installs to `%LOCALAPPDATA%\Programs\ExcelIntelligenceAgent`, per-user
   (`PrivilegesRequired=lowest`), Start-Menu + optional Desktop shortcut, `CloseApplications=yes`,
   `RestartApplications=no`.

Install location on a colleague PC:
`%LOCALAPPDATA%\Programs\ExcelIntelligenceAgent\ExcelIntelligenceAgent.exe`.

### Build gotchas
- **Kill any running `ExcelIntelligenceAgent.exe` / `EXCEL` before building**, or PyInstaller
  `COLLECT` fails copying `_internal` (file lock), and the installer can't be replaced.
- **Inno error 110 "EndUpdateResource failed … exclude the Output folder from your antivirus":**
  antivirus momentarily locked `Setup.exe` during icon embedding. It's intermittent — just retry
  the ISCC step, or add a Windows Defender exclusion for the project folder
  (`Add-MpPreference -ExclusionPath <folder>`, needs admin).
- Run `gh`/`ISCC` from cmd.exe or PowerShell, NOT Git-bash (Git-bash mangles `/D...` ISCC defines
  and `/`-prefixed args).

---

## 11. How publishing is done (so colleagues auto-update)

`build.bat` step 4 publishes the freshly built installer to GitHub Releases via the `gh` CLI:
```
gh release create v<APP_VERSION> "Output\ExcelIntelligenceAgent-Setup.exe" \
   --title "v<APP_VERSION>" --notes "..." --repo alamehmazen123/the_excel_agent
```
`build.bat` locates `gh` even if not on PATH (checks `%ProgramFiles%\GitHub CLI\gh.exe` and
`%LOCALAPPDATA%\Programs\GitHub CLI\gh.exe`). If `gh` is missing or not logged in, it prints manual
instructions and continues.

### One-time setup (developer machine)
1. `winget install GitHub.cli` (already installed at `C:\Program Files\GitHub CLI\gh.exe`).
2. `gh auth login` → GitHub.com → HTTPS → **Paste an authentication token** (a **classic** token
   with the **`repo`** scope — a read-only/fine-grained token gives `HTTP 401 Bad credentials` on
   asset upload). The token is saved in the OS keyring; `build.bat` then publishes with no prompt.
   - GOTCHA: the device-flow code is **printed in the terminal**, not emailed. At the
     "How would you like to authenticate" menu, you must arrow to "Paste an authentication token"
     and press Enter BEFORE pasting.

### Releasing a new version
1. Bump `APP_VERSION` in `config.py`.
2. Run `build.bat` → builds, packages, and publishes `v<APP_VERSION>` to GitHub.
3. The FIRST build that contains the GitHub-updater must be sent to colleagues **manually once**
   (their old build can't auto-update *to* the first release). After that, every release is
   automatic — they update silently on next close-and-reopen.

### Manual publish alternatives (no `gh` / no token)
- Web UI: `https://github.com/<repo>/releases` → Draft a new release → tag `v<version>` → drag in
  `Output\ExcelIntelligenceAgent-Setup.exe` → Publish.
- With a stored `gh` login you can also publish a pre-built installer without rebuilding:
  `gh release create v<version> "Output\ExcelIntelligenceAgent-Setup.exe" --repo <repo>`.

### Verifying a publish (do this after every release)
```
python -c "import requests,config; from core import updater; \
 r=requests.get(config.UPDATE_MANIFEST_URL,headers={'Accept':'application/vnd.github+json'}); \
 print('status',r.status_code, r.json().get('tag_name'), [a['name'] for a in r.json().get('assets',[])]); \
 print('update for old user:', bool(updater.check_for_update(config.UPDATE_MANIFEST_URL,'1.0.0')))"
```
Expect status 200, the new tag, the `.exe` asset, and `update for old user: True`. (`/latest`
returns 404 while there are no published releases; drafts are invisible to unauthenticated reads, so
also check the Releases web page if unsure.)

---

## 12. Versioning, footer, branding

- `APP_VERSION` in `config.py` is the single source of truth; the installer version comes from it
  (`/DAppVersion`), and GitHub tags are `v<APP_VERSION>`.
- `buildinfo.BUILD_DATE` (stamped) shows in the app footer next to the version.
- `config.ORG_NAME` = "SAHEL GENERAL HOSPITAL" — shown as a blue banner in the app window header.
- The Settings (gear) button is intentionally small (30×30, icon-only) with an amber warning above
  the key field: "DO NOT CHANGE, unless you are aware of what you are doing (have a new key)".

---

## 13. Responsive UI

`ui/main_window.py` wraps all content in a `QScrollArea` and `_size_to_screen()` opens the window at
~62% width × 88% height of the available screen (clamped), centered, with a small 480×460 minimum,
so it autofits any screen and never clips on small laptops.

---

## 14. Testing

- `python -m pytest tests/test_engine.py -q` — 6 tests: type detection, multi-sheet detection,
  sheet creation + original-data integrity, offline summary fallback, clean re-run regeneration,
  empty-workbook error. The suite exercises the real Excel COM path when Excel is available.
- Offscreen UI smoke tests: set `QT_QPA_PLATFORM=offscreen` and construct `MainWindow` / dialogs.
- COM inspection of produced workbooks: open with `win32com`, but add a retry loop — a freshly
  closed Excel instance can transiently throw `RPC_E_DISCONNECTED (0x80010108)` /
  "Call was rejected by callee". Single product runs are stable; rapid test instances can race.

---

## 15. Distribution summary (what to send / requirements)

- Give colleagues **`Output\ExcelIntelligenceAgent-Setup.exe`** (zip it for transfer to survive
  email/AV). They double-click; first launch may show SmartScreen "Run anyway" (unsigned).
- **Microsoft Excel must be installed** for real pivots/CF/charts; otherwise static tables.
- Updates and the Groq AI both require network access (`github.com` / `api.groq.com`). A hospital
  firewall blocking either will silently fall back (no auto-update / template summaries). The
  app never errors because of it.
