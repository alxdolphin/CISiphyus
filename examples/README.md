# downstream applications

Prototypical, fully functional applications that consume CISiphyus exports. Clone this
repo, bootstrap auth, and run — no private monorepo checkout is required.n

Each application retrieves its CIS export via cisiphyus when you omit `--workbook` /
`--metric-workbook` (age-gated refresh when missing or stale). Pass those flags only to
pin a local file for CI or reproducible reruns.

| Application | CISiphyus report | Output |
|---|---|---|
| Accreditation monitoring | `accreditation` | Site flag CSVs / summaries |
| Student metrics audit | `student_metrics_summary` | Row-level metric/scale flags |
| Goal achievement audit | `goal_tracking_student_goals` + `student_metrics_summary` + `accreditation` | GAR exceptions + drilldown coverage flags |
| Goal achievement summary | `goal_achievement` | Aggregate goal-area counts |
| QPR provisioning | `student_metrics_summary` | Per-site `Q{n}_{Site}_QPR.xlsx` |
| Longitudinal trend tracker | `accreditation` + `student_metrics_summary` | QoQ movement + regression CSVs / summaries |

`audit goal-achievement` runs GAR logic in [`examples/audit/audit.py`](examples/audit/audit.py) on **goal_tracking_student_goals** (`CIS_StudentProgress_Detail`), then cross-checks **accreditation Student Drilldown** against **student metrics**. Row-level metric QA is `audit metrics`. The aggregate Goal Achievement Summary pull is not used for this audit.

## goal achievement audit (`goal_achievement/`)

GAR on goal tracking plus drilldown cross-check.

```bash
cisiphyus pull goal_tracking_student_goals --school-year SY25-26
cisiphyus audit goal-achievement
cisiphyus audit goal-achievement --school-year SY25-26
```

Pin local workbooks:

```bash
cisiphyus audit goal-achievement \\
  --goal-progress-workbook path/to/GoalTracking.xlsx \\
  --student-metrics-workbook path/to/StudentMetrics.xlsx \\
  --accreditation-workbook path/to/Accreditation_Report.xlsx
```

Outputs: `goal_achievement_summary.md`, `goal_achievement_summary.json`, `gar_exceptions.csv`, `gar_audit_report.md`, `drilldown_exceptions.csv`. Per-SY archive: `artifacts/goal_achievement_audit/{SY}_GoalAchievement_AUDIT.csv`.

## goal achievement summary (`goal_achievement` pull)

Aggregate counts by goal area (`CIS_GoalAchievement_Summary`, FormID `1000002372`).

```bash
cisiphyus pull goal_achievement
```

## accreditation site monitoring (`accreditation/`)

Retrieves the `accreditation` report (`ui_export` strategy) and flags site-level
compliance issues: site-coordination reporting gaps, Tier I support counts, and
case-management completeness.

```bash
python examples/accreditation/accreditation_monitor.py --output-dir /tmp/accreditation-out
```

Force a fresh export even when the cached workbook is still within the 24h gate:

```bash
python examples/accreditation/accreditation_monitor.py \
  --output-dir /tmp/accreditation-out \
  --force-fetch
```

Outputs: `all_flags.csv`, `sc_flags.csv`, `cm_flags.csv`, `tier1_flags.csv`,
`monitoring_summary.json`, `monitoring_summary.md`.

## qpr workbook provisioning (`qpr/`)

Retrieves `student_metrics_summary` (`direct_url` strategy), then generates per-site QPR
import workbooks (`.xlsx`) using bundled [`qpr_tools.py`](qpr/qpr_tools.py) provision
logic and default fixtures under `examples/qpr/fixtures/`.

```bash
cisiphyus qpr --grading-period 2.0
```

Single site:

```bash
cisiphyus qpr --grading-period 2.0 --school-name "Lincoln HS"
```

Override bundled template, deadlines, or site-staff inputs when your affiliate uses
local workbooks:

```bash
cisiphyus qpr --grading-period 2.0 \
  --template /path/to/QPR_Import_Template.xlsx \
  --deadlines /path/to/ReportingDeadlines.xlsx \
  --site-staff-list /path/to/Site_Staff_List.xlsx
```

Outputs: `artifacts/qpr/{school-year}/Q{n}/Q{n}_{Site}_QPR.xlsx` per site.

## longitudinal trend tracker (`trends/`)

Discovers available reporting periods for each school year, captures snapshots, and
compares consecutive quarters.

```bash
cisiphyus trend SY25-26
cisiphyus trend SY25-26 SY24-25
cisiphyus trend --all
```

### period discovery

Periods come from `artifacts/qpr/{school-year}/Q{n}/` folders (same layout as QPR output),
cached archives under `artifacts/archives/{school-year}/Q{n}/`, and existing snapshots.
`--all` also includes every school year listed in `school_year_programs` inside
`config/reports.yaml`. Years without QPR quarter folders are pulled from CISDM,
archived under `artifacts/archives/{school-year}/EOY/`, and captured as
`artifacts/snapshots/{school-year}/EOY/` snapshots (metrics-only for historical years).

`--all` also runs cross-year compares (`artifacts/trends/cross_year/`) between
consecutive EOY snapshots (prior school year EOY vs next school year EOY). Years
without an EOY snapshot are skipped until one is captured.

`cisiphyus trend cross-year` (alias `cisiphyus trend eoy`) pulls fresh CISDM exports,
refreshes EOY snapshots, runs cross-year compares, and writes
`artifacts/trends/cross_year/index.html`. Use `--no-fetch` to regenerate HTML only.

```bash
cisiphyus trend cross-year
cisiphyus trend eoy
cisiphyus trend eoy --no-fetch
```

Trend pulls **student metrics only** by default. Accreditation is skipped unless
`ACCREDITATION_FETCH_FROM_CISDM=1` or you pass `--accreditation` on `trend` / `eoy`.

```bash
cisiphyus trend eoy --accreditation
ACCREDITATION_FETCH_FROM_CISDM=1 cisiphyus trend SY25-26
```

The cross-year HTML report has two sections:

- **Student Metrics summary totals** — aggregate audit totals compared between EOY files (different rosters each year; not the same students tracked).
- **Student goal progress** — on-track rate and year-over-year goals improved / fell behind for matched student goal rows.

Snapshots store `metrics/progress_rollup.json` for goal progress. Re-run `backfill_progress_rollups` or `cisiphyus trend --all` to populate historical snapshots.

Compare output warns when period snapshots share identical workbook hashes (CISDM exports
are school-year-scoped, not quarter-scoped — run trend at each period close for meaningful QoQ).

Per-year status is printed: `snapshotted`, `compared`, `skipped`, or `failed`.

Run at the end of each grading period so each quarter gets its own CISDM pull. Older
quarters without a cached pull are skipped until you have snapshot history for them.

```bash
cisiphyus trend SY25-26          # pull + capture latest missing period, compare pairs
cisiphyus trend SY25-26 --force-fetch
```

Snapshots: `artifacts/snapshots/{school-year}/{period}/`  
Compare outputs: `artifacts/trends/{school-year}/{Q1_vs_Q2}/`

## environment variables

| Variable | Purpose |
|---|---|
| `CISIPHYUS_ROOT` | Override the cisiphyus repo root (default: this repo) |
| `CISPHYUS_HEADED` | `1` to show the browser during cisiphyus retrieval |
| `ACCREDITATION_WORKBOOK` / `ACCREDITATION_LOCAL_INPUTS_DIR` | Accreditation workbook destination |
| `ACCREDITATION_MAX_AGE_HOURS` | Freshness gate for the accreditation workbook (default 24) |
| `GOAL_ACHIEVEMENT_GOAL_PROGRESS_WORKBOOK` / `GOAL_ACHIEVEMENT_STUDENT_METRICS_WORKBOOK` / `GOAL_ACHIEVEMENT_ACCREDITATION_WORKBOOK` | Input workbook paths |
| `GOAL_ACHIEVEMENT_LOCAL_INPUTS_DIR` | Cached CISDM pull copies (default `artifacts/goal_achievement/`) |
| `GOAL_ACHIEVEMENT_OUTPUT_DIR` | Audit output root (default `artifacts/goal_achievement/`) |
| `QPR_STUDENT_METRICS_WORKBOOK` / `QPR_LOCAL_INPUTS_DIR` | Student metrics workbook destination |
| `QPR_OUTPUT_DIR` | QPR provision output root (default `artifacts/qpr/`) |
| `QPR_FETCH_STUDENT_METRICS` | `1` to force a refresh even when fresh |
| `TREND_SNAPSHOTS_DIR` | Snapshot archive root (default `artifacts/snapshots/`) |
| `TREND_INPUTS_DIR` | Cached CISDM pull copies (default `artifacts/archives/`) |
| `TREND_OUTPUT_DIR` | Trend compare output root (default `artifacts/trends/`) |
| `TREND_SCHOOL_YEAR` | Default school year when none passed to `cisiphyus trend` |

## tests

```bash
pip install -e .[dev]
pytest tests/
```

Tests pass `--workbook` / `--metric-workbook` fixtures explicitly; no live CISDM access is required.
