#!/usr/bin/env bash
# 导出/渲染 pptx -> PDF + 逐页PNG(macOS/Linux,LibreOffice + Poppler)
# 用法: bash export_deck.sh deck.pptx [输出目录]
# 依赖: soffice(LibreOffice)、pdftoppm(Poppler)。缺哪个先装:apt install libreoffice poppler-utils
set -euo pipefail
PPTX="$1"
OUTDIR="${2:-$(dirname "$PPTX")/export}"
mkdir -p "$OUTDIR"

command -v soffice >/dev/null || { echo "缺少 LibreOffice(soffice)"; exit 1; }
command -v pdftoppm >/dev/null || { echo "缺少 Poppler(pdftoppm)"; exit 1; }

soffice --headless --convert-to pdf --outdir "$OUTDIR" "$PPTX"
PDF="$OUTDIR/$(basename "${PPTX%.*}").pdf"
mkdir -p "$OUTDIR/png"
pdftoppm -png -r 150 "$PDF" "$OUTDIR/png/page"

echo "完成: $PDF ; $OUTDIR/png/*.png"
