#!/usr/bin/env bash
# Run the two DOCX PDF matrices as PDF/A-2u:
#
#   * curated: `pdf_matrix.py --prepare` (24 fixtures)
#   * full:    every checked-in DOCX under crates/docx-edit/tests/fixtures
#
# Both exports pass `--pdfa-2u` to `export_pdf_batch`; documents the writer
# cannot represent conformantly (for example an uncovered emoji scalar) are
# recorded as `unsupported` and reported separately instead of failing the
# matrix. Both checks pass `--require-pdfa-2u` to `pdf_matrix.py`, so every
# produced PDF must validate with the pinned veraPDF release. The export step
# is skipped with a warning when the batch example does not expose
# `--pdfa-2u` yet.
#
# Usage:
#   bash scripts/office-quality/pdfa_matrix.sh
#
# Environment:
#   VERAPDF              veraPDF executable (default: PDFA_MATRIX_DIR/verapdf/bin/verapdf)
#   VERAPDF_VERSION      pinned release (default: 1.30.4)
#   PDFA_MATRIX_DIR      work directory (default: .source/pdf-matrix)
#   PYTHON               interpreter with pypdfium2 (default: harness venv or python3)
#   CARGO                cargo executable (default: cargo)
#
# Reports:
#   $PDFA_MATRIX_DIR/report-curated.json
#   $PDFA_MATRIX_DIR/report-full.json
#   $PDFA_MATRIX_DIR/verapdf/{curated,full}/<fixture>.verapdf.json

set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

work="${PDFA_MATRIX_DIR:-.source/pdf-matrix}"
python="${PYTHON:-}"
if [ -z "$python" ]; then
    if [ -x ".source/office-quality/venv/bin/python" ]; then
        python=".source/office-quality/venv/bin/python"
    else
        python="python3"
    fi
fi
cargo="${CARGO:-cargo}"
verapdf="${VERAPDF:-$work/verapdf/bin/verapdf}"
version="${VERAPDF_VERSION:-1.30.4}"
example="crates/betteroffice-docx/examples/export_pdf_batch.rs"
full_input="crates/docx-edit/tests/fixtures"

if ! grep -q -- '--pdfa-2u' "$example"; then
    mkdir -p "$work"
    cat >&2 <<EOF
pdfa_matrix: $example does not expose --pdfa-2u yet.
pdfa_matrix: skipped both PDF/A-2u matrices; standard PDF export is unchanged.
EOF
    {
        echo "skipped: $example does not expose --pdfa-2u"
        date -u +"checked at %Y-%m-%dT%H:%M:%SZ"
    } > "$work/skipped.txt"
    exit 0
fi

if [ ! -x "$verapdf" ]; then
    cat >&2 <<EOF
pdfa_matrix: veraPDF is not executable at $verapdf
  install the pinned release first:
    bash scripts/office-quality/install-verapdf.sh $work/verapdf
EOF
    exit 2
fi

mkdir -p "$work"
reported_version="$("$verapdf" --version 2>&1 | sed -n 's/^veraPDF \([0-9][0-9.]*\).*/\1/p' | sed -n '1p')"
if [ "$reported_version" != "$version" ]; then
    echo "pdfa_matrix: veraPDF $reported_version found, require $version" >&2
    exit 2
fi

export VERAPDF="$verapdf"
export VERAPDF_VERSION="$version"

status=0
run_matrix() {
    local name="$1"
    local input="$2"
    local output="$3"
    local report="$4"
    local verapdf_reports="$5"

    echo "pdfa_matrix: $name matrix"
    echo "  input:   $input"
    echo "  output:  $output"
    echo "  report:  $report"
    echo "  veraPDF: $verapdf_reports"

    if ! "$cargo" run --quiet -p betteroffice-docx --example export_pdf_batch -- \
        --pdfa-2u --allow-unsupported "$input" "$output"; then
        echo "pdfa_matrix: $name export failed; checking the manifest it wrote" >&2
        status=1
    fi
    if ! "$python" scripts/office-quality/pdf_matrix.py \
        "$output/manifest.json" \
        --require-pdfa-2u \
        --allow-unsupported \
        --strict-fonts \
        --verapdf "$verapdf" \
        --verapdf-version "$version" \
        --verapdf-report-dir "$verapdf_reports" \
        --json "$report"; then
        echo "pdfa_matrix: $name validation failed" >&2
        status=1
    fi
}

"$python" scripts/office-quality/pdf_matrix.py --prepare "$work/fixtures"

curated_count="$(find "$work/fixtures" -type f -name '*.docx' | wc -l | tr -d ' ')"
full_count="$(find "$full_input" -type f -name '*.docx' | wc -l | tr -d ' ')"
if [ "$curated_count" != "24" ]; then
    echo "pdfa_matrix: warning: curated matrix has $curated_count fixtures, expected 24" >&2
fi
if [ "$full_count" != "103" ]; then
    echo "pdfa_matrix: warning: full matrix has $full_count fixtures, expected 103" >&2
fi
echo "pdfa_matrix: curated=$curated_count fixture(s), full=$full_count fixture(s)"

run_matrix curated "$work/fixtures" "$work/pdfs-curated" \
    "$work/report-curated.json" "$work/verapdf/curated"
run_matrix full "$full_input" "$work/pdfs-full" \
    "$work/report-full.json" "$work/verapdf/full"

if [ "$status" -ne 0 ]; then
    echo "pdfa_matrix: PDF/A-2u matrices failed" >&2
    exit "$status"
fi
echo "pdfa_matrix: both PDF/A-2u matrices passed"
