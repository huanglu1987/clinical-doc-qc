#!/usr/bin/env bash
# clinical-doc-qc 安装脚本
# 用法：bash install.sh            → 装到 Codex（~/.codex/skills）
#      bash install.sh --claude   → 装到 Claude Code（~/.claude/skills）
set -euo pipefail

NAME="clinical-doc-qc"
SRC="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
TARGET_ROOT="$HOME/.codex/skills"
[[ "${1:-}" == "--claude" ]] && TARGET_ROOT="$HOME/.claude/skills"
DEST="$TARGET_ROOT/$NAME"

echo "▸ 安装 $NAME"
echo "  源目录：$SRC"
echo "  目标：  $DEST"
echo

# ── 1. Python 与依赖 ─────────────────────────────────────────
command -v python3 >/dev/null 2>&1 || { echo "✗ 未找到 python3，请先安装 Python 3.8+"; exit 1; }
echo "▸ 检查 Python 依赖…"
MISSING=()
for mod in docx pypdf fitz; do
  python3 -c "import $mod" >/dev/null 2>&1 || MISSING+=("$mod")
done
if [ ${#MISSING[@]} -gt 0 ]; then
  echo "  缺少：${MISSING[*]}，正在安装…"
  python3 -m pip install --quiet --upgrade python-docx pypdf pymupdf \
    || { echo "✗ 依赖安装失败，请手动执行：pip3 install python-docx pypdf pymupdf"; exit 1; }
fi
echo "  ✓ 依赖就绪"

# ── 2. 复制文件 ──────────────────────────────────────────────
mkdir -p "$TARGET_ROOT"
if [ -d "$DEST" ]; then
  BAK="$DEST.bak.$(date +%Y%m%d%H%M%S)"
  echo "▸ 已存在旧版本，备份到 $BAK"
  mv "$DEST" "$BAK"
fi
mkdir -p "$DEST"
for item in SKILL.md README.md scripts references templates; do
  [ -e "$SRC/$item" ] && cp -R "$SRC/$item" "$DEST/"
done
chmod +x "$DEST/scripts/docqc.py" 2>/dev/null || true
echo "  ✓ 文件已复制"

# ── 3. 自检 ──────────────────────────────────────────────────
echo "▸ 自检…"
python3 "$DEST/scripts/docqc.py" --help >/dev/null 2>&1 \
  || { echo "✗ 自检失败：docqc.py 无法运行"; exit 1; }
echo "  ✓ 自检通过"

cat <<EOF

════════════════════════════════════════════════════════════
安装完成。

命令行直接用：
  python3 $DEST/scripts/docqc.py check 方案.docx --profile cde
  python3 $DEST/scripts/docqc.py check protocol.docx --profile fda

在 Codex / Claude Code 里直接说：
  用 clinical-doc-qc 按 CDE 规则审核这套资料：<文件路径...>

正式提交的材料，记得按 templates/review-prompt.md 做双通道核查
（两个互不共享历史的会话各跑一遍再合并，实测 30 条发现只有 1 条重合）。
════════════════════════════════════════════════════════════
EOF
