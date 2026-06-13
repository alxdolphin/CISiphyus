#!/usr/bin/env bash
# refresh prototypical downstream applications from a CIS monorepo checkout.
#
# the CIS monorepo is canonical; these copies are practical reference
# implementations that stay aligned with production fetch/monitor/provision logic.
# run this after changing production logic, then re-apply the documented contract
# adjustments (see docs/superpowers/specs/2026-06-12-downstream-examples-design.md):
# - _default_cisiphyus_root() resolves to this repo root
# - _cisiphyus_cmd() passes only <report_id> and --headed
# - help text references examples/ paths
#
# usage: CIS_MONOREPO_ROOT=/path/to/CIS ./examples/sync_from_cis.sh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(dirname "$SCRIPT_DIR")"
# default assumes this repo is nested at <CIS>/tools/CISiphyus
CIS_MONOREPO_ROOT="${CIS_MONOREPO_ROOT:-$(dirname "$(dirname "$REPO_ROOT")")}"

if [[ ! -d "$CIS_MONOREPO_ROOT/tools" ]]; then
    echo "error: CIS monorepo not found at $CIS_MONOREPO_ROOT (set CIS_MONOREPO_ROOT)" >&2
    exit 1
fi

# manifest: <source relative to CIS monorepo> -> <destination relative to this repo>
MANIFEST=(
    "tools/accreditation/scripts/accreditation_cisiphyus_fetch.py:examples/accreditation/accreditation_cisiphyus_fetch.py"
    "tools/accreditation/scripts/accreditation_workbook.py:examples/accreditation/accreditation_workbook.py"
    "tools/accreditation/scripts/accreditation_monitor.py:examples/accreditation/accreditation_monitor.py"
    "tools/qpr/qpr_cisiphyus_prefetch.py:examples/qpr/qpr_cisiphyus_prefetch.py"
    "tools/tests/fixtures/accreditation_monitor_minimal.xlsx:examples/accreditation/fixtures/accreditation_monitor_minimal.xlsx"
)

status=0
changed=0
for entry in "${MANIFEST[@]}"; do
    src="$CIS_MONOREPO_ROOT/${entry%%:*}"
    dest="$REPO_ROOT/${entry##*:}"
    if [[ ! -f "$src" ]]; then
        echo "MISSING  $src" >&2
        status=1
        continue
    fi
    if [[ -f "$dest" ]] && cmp -s "$src" "$dest"; then
        echo "ok       ${entry##*:}"
        continue
    fi
    mkdir -p "$(dirname "$dest")"
    cp "$src" "$dest"
    echo "updated  ${entry##*:}"
    changed=$((changed + 1))
done

if (( changed > 0 )); then
    echo
    echo "WARNING: $changed file(s) updated from production copies." >&2
    echo "re-apply the contract adjustments listed at the top of this script before committing." >&2
fi
exit "$status"
