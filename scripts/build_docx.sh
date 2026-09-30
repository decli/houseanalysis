#!/usr/bin/env bash
# Markdown -> Word。需要 pandoc；reference.docx 提供中文字体与表格样式。
set -e
cd "$(dirname "$0")/../report"
export LANG=C.UTF-8 LC_ALL=C.UTF-8
pandoc "国创光谷上城_持有还是出售_分析报告.md" \
  -o "国创光谷上城_持有还是出售_分析报告.docx" \
  --reference-doc=../scripts/reference.docx \
  --resource-path=.:..
echo "done: report/国创光谷上城_持有还是出售_分析报告.docx"
