# downstream applications

Prototypical, fully functional applications that consume CISiphyus exports. Each one
retrieves its CIS export via cisiphyus when you omit `--workbook` / `--metric-workbook`
(age-gated refresh when missing or stale). Pass those flags only to pin a local file—for
CI or reproducible reruns.

Production counterparts live in the CIS monorepo; Python modules here are trimmed
copies kept in sync with `./sync_from_cis.sh` (CIS stays canonical).

| Application | CISiphyus report | Output |
|---|---|---|
| Accreditation monitoring | `accreditation` | Site flag CSVs / summaries |
| QPR provisioning | `student_metrics_summary` | Per-site `Q{n}_{Site}_QPR.xlsx` |

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
import workbooks (`.xlsx`) using production [`tools/qpr/qpr_tools.py`](../../tools/qpr/qpr_tools.py)
provision logic (`CIS_MONOREPO_ROOT` defaults to `../..` when nested at `tools/CISiphyus`).

```bash
python examples/qpr/qpr_provision.py /tmp/qpr-provision-out \
  --grading-period 2.0 \
  --no-site-staff-filter
```

Single site:

```bash
python examples/qpr/qpr_provision.py /tmp/qpr-provision-out \
  --grading-period 2.0 \
  --school-name "Lincoln HS" \
  --no-site-staff-filter
```

Outputs: `{output}/{school-year}/Q{n}/Q{n}_{Site}_QPR.xlsx` per site.

## environment variables

| Variable | Purpose |
|---|---|
| `CISIPHYUS_ROOT` | Override the cisiphyus repo root (default: this repo) |
| `CISPHYUS_HEADED` | `1` to show the browser during cisiphyus retrieval |
| `ACCREDITATION_WORKBOOK` / `ACCREDITATION_LOCAL_INPUTS_DIR` | Accreditation workbook destination |
| `ACCREDITATION_MAX_AGE_HOURS` | Freshness gate for the accreditation workbook (default 24) |
| `QPR_STUDENT_METRICS_WORKBOOK` / `QPR_LOCAL_INPUTS_DIR` | Student metrics workbook destination |
| `QPR_FETCH_STUDENT_METRICS` | `1` to force a refresh even when fresh |
| `CIS_MONOREPO_ROOT` | CIS checkout containing `tools/qpr/qpr_tools.py` (required for provisioning) |

## keeping copies in sync

```bash
CIS_MONOREPO_ROOT=/path/to/CIS ./examples/sync_from_cis.sh
```

The script copies the vendored manifest and warns when files changed; re-apply the
contract adjustments listed at the top of the script after a sync.

## tests

```bash
pip install -e .[dev]
pytest tests/
```

Tests pass `--workbook` / `--metric-workbook` fixtures explicitly; no live CISDM access is required.
QPR provisioning tests require `CIS_MONOREPO_ROOT` (or the default nested layout).
