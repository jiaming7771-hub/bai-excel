#!/usr/bin/env bash
# 在 macOS 上打包给普通用户直接用的 .app
set -euo pipefail
cd "$(dirname "$0")/.."

if [[ ! -d .venv ]]; then
  python3 -m venv .venv
fi
# shellcheck disable=SC1091
source .venv/bin/activate
pip install -q -r requirements.txt

pkill -f "Excel小工具箱.app/Contents/MacOS" 2>/dev/null || true
pyinstaller --noconfirm --clean excel_toolbox.spec

OUT_DIR="${HOME}/Desktop/Excel小工具箱发布"
mkdir -p "${OUT_DIR}"
rm -rf "${OUT_DIR}/Excel小工具箱.app"
cp -R "dist/Excel小工具箱.app" "${OUT_DIR}/Excel小工具箱.app"
xattr -cr "${OUT_DIR}/Excel小工具箱.app" || true

# 同步一份到桌面方便双击
rm -rf "${HOME}/Desktop/Excel小工具箱.app"
cp -R "${OUT_DIR}/Excel小工具箱.app" "${HOME}/Desktop/Excel小工具箱.app"
xattr -cr "${HOME}/Desktop/Excel小工具箱.app" || true

# 打成 zip 方便发给用户
rm -f "${OUT_DIR}/Excel小工具箱-macOS.zip"
ditto -c -k --keepParent "${OUT_DIR}/Excel小工具箱.app" "${OUT_DIR}/Excel小工具箱-macOS.zip"

echo "OK: ${OUT_DIR}/Excel小工具箱.app"
echo "OK: ${OUT_DIR}/Excel小工具箱-macOS.zip"
