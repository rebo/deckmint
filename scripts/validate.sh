#!/usr/bin/env bash
set -e
cd "$(dirname "$0")/.."
echo "Building features.pptx..."
cargo run --example features --quiet
echo ""
echo "=== Schema validation (OOXML SDK) ==="
npx --yes @xarsh/ooxml-validator features.pptx
echo ""
echo "=== Repair-pattern validation ==="
python3 scripts/check_pptx.py features.pptx
