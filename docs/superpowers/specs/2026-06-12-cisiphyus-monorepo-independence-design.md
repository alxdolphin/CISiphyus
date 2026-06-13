# CISiphyus monorepo independence

**Status:** Approved for implementation | **Date:** 2026-06-12

## goal

Affiliate data teams clone CISiphyus to learn programmatic CISDM exports. The repo and its
examples must run after clone + bootstrap with no dependency on an unpublished CIS monorepo.

## problem

QPR provisioning previously loaded `qpr_tools.py` from `CIS_MONOREPO_ROOT` (defaulting to
`../..`). Accreditation was already self-contained; QPR was not. Documentation and tests
referenced monorepo paths and skipped when the monorepo was absent.

## decision

- Vendor full `qpr_tools.py` under `examples/qpr/` with path constants repointed to bundled fixtures
- Load QPR tools from `examples/qpr/qpr_tools.py` only; remove `CIS_MONOREPO_ROOT`
- Bundle default template, deadlines, and minimal site-staff fixtures under `examples/qpr/fixtures/`
- Delete `examples/sync_from_cis.sh`; CISiphyus is the sole source of truth for affiliates
- Site-staff filter: use bundled fixture when present; otherwise disable filter without error

## layout changes

| Before | After |
|---|---|
| `qpr_tools_loader` → monorepo `tools/qpr/qpr_tools.py` | `qpr_tools_loader` → `examples/qpr/qpr_tools.py` |
| Template/deadlines from monorepo evaluation paths | Defaults from `examples/qpr/fixtures/` |
| `sync_from_cis.sh` required for maintainers | Removed |
| Tests skip when monorepo missing | Tests always use bundled module |

## affiliate workflow

1. Clone CISiphyus
2. `pip install -e .[dev]` and `python run.py --bootstrap`
3. Run examples with live cisiphyus retrieval (no `--workbook` / `--metric-workbook`)
4. Optionally override `--template`, `--deadlines`, `--site-staff-list` with affiliate-local files

## out of scope

- Publishing the private CIS monorepo
- Trimming unused `qpr_tools.py` CLI commands (validate, docx, Drive, pipeline)
- Replacing bundled fixtures with affiliate-specific school-year files by default
