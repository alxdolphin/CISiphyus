# PR draft: feat/downstream-examples

Local review copy for [PR #3](https://github.com/alxdolphin/CISiphyus/pull/3). Update the GitHub PR body with this text when you push (`git push` then paste or `gh pr edit 3 --body-file docs/pr-feat-downstream-examples.md`).

---

## Summary

Add self-contained downstream examples that demonstrate programmatic CISDM exports for affiliate data teams. After clone, bootstrap, and `pip install -e .[dev]`, both examples run with **live cisiphyus retrieval** and **no private monorepo checkout**.

- **Accreditation monitoring** — live fetch of `accreditation` report, site-level compliance flag CSVs/summaries
- **QPR provisioning** — live fetch of `student_metrics_summary`, per-site QPR import workbooks via **bundled** `examples/qpr/qpr_tools.py`
- **CLI dispatch** — `cisiphyus accreditation …` and `cisiphyus qpr …` route to example entrypoints
- **Monorepo independence** — removed `CIS_MONOREPO_ROOT`, deleted `examples/sync_from_cis.sh`; bundled template, deadlines, and site-staff fixtures under `examples/qpr/fixtures/`
- **Nix dev shell** — `flake.nix`, `shell.nix`, `flake.lock` for Playwright-on-NixOS

## Affiliate workflow

```bash
git clone https://github.com/alxdolphin/CISiphyus.git
cd CISiphyus
pip install -e .[dev] && playwright install chromium
python run.py --bootstrap

cisiphyus accreditation \
  --output-dir examples/accreditation/local_inputs/monitor-out

cisiphyus qpr examples/qpr/output --grading-period 2.0
```

Optional overrides for affiliate-local workbooks: `--template`, `--deadlines`, `--site-staff-list`.

## Design specs

- [`docs/superpowers/specs/2026-06-12-downstream-examples-design.md`](docs/superpowers/specs/2026-06-12-downstream-examples-design.md)
- [`docs/superpowers/specs/2026-06-12-cisiphyus-monorepo-independence-design.md`](docs/superpowers/specs/2026-06-12-cisiphyus-monorepo-independence-design.md)

## Commits (newest first, vs `origin/main`)

| Commit | Message |
|---|---|
| `443bbf2` | docs: document affiliate self-contained workflow |
| `45ef922` | test(examples): require bundled qpr_tools in QPR tests |
| `48252e3` | refactor(examples): remove CIS monorepo loader and sync script |
| `cbefca8` | feat(examples): vendor qpr_tools for self-contained QPR provisioning |
| `921f45b` | docs(examples): add design spec and usage documentation |
| `7cb8246` | feat(cli): dispatch cisiphyus accreditation and qpr to examples |
| `98e2910` | test(examples): add fixture-driven accreditation and QPR tests |
| `d495d8b` | feat(examples): add QPR workbook provisioning application |
| `3e3237a` | feat(examples): add accreditation site monitoring application |
| `dfabb68` | chore(examples): track fixtures and example tests in gitignore |
| `88adb67` | chore(nix): add flake.lock |
| `00225eb` | feat(nix): add Nix flake and shell configuration… |

## Test plan

- [x] `env -u CIS_MONOREPO_ROOT .venv/bin/python -m pytest tests/ -q` (15 passed)
- [x] QPR provision smoke with bundled fixtures (no monorepo env)
- [ ] `cisiphyus accreditation --output-dir …` (live CISDM)
- [ ] `cisiphyus qpr examples/qpr/output --grading-period 2.0` (live CISDM metrics fetch + provision)

## Breaking / removal notes

- **`CIS_MONOREPO_ROOT`** is no longer used by QPR provisioning
- **`examples/sync_from_cis.sh`** removed; CISiphyus is the sole source of truth for affiliates
- QPR default site-staff filter uses bundled `site_staff_list_minimal.xlsx` when present; omit with `--no-site-staff-filter`
