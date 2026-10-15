#!/usr/bin/env bash
# =====================================================================
# PDF2Excel · macOS 一键构建脚本
# 产出：dist/PDF2Excel.app（Apple Silicon / Intel 取决于当前机器架构）
#
# 用法：bash build_mac.sh
# =====================================================================
set -euo pipefail
cd "$(dirname "$0")"

PY="${PY:-python3}"
if [ -x /opt/homebrew/bin/python3.14 ]; then PY=/opt/homebrew/bin/python3.14; fi

echo "==> 使用 Python：$($PY --version 2>&1)"

# 1) 虚拟环境与依赖
if [ ! -d .venv ]; then
  echo "==> 创建虚拟环境 .venv"
  "$PY" -m venv .venv
fi
.venv/bin/pip install --quiet --upgrade pip
.venv/bin/pip install --quiet -r requirements-dev.txt

# 2) 图标（缺失时生成）
if [ ! -f assets/icon.icns ]; then
  echo "==> 生成应用图标"
  .venv/bin/python tools/make_icons.py
fi

# 3) PyInstaller 打包（onedir .app，启动快）
echo "==> PyInstaller 打包中…"
.venv/bin/pyinstaller \
  --noconfirm --clean --windowed \
  --name "PDF2Excel" \
  --osx-bundle-identifier "com.pdf2excel.app" \
  --icon assets/icon.icns \
  --collect-all pywebview \
  --add-data "ui:ui" \
  --add-data "assets:assets" \
  main.py

# 4) ad-hoc 签名（缓解 Gatekeeper 提示；正式分发请使用 Apple 开发者证书）
if command -v codesign >/dev/null 2>&1; then
  echo "==> ad-hoc 签名"
  codesign --force --deep --sign - "dist/PDF2Excel.app" >/dev/null 2>&1 || true
fi

echo
echo "构建完成：$(pwd)/dist/PDF2Excel.app"
echo "首次打开方式：右键点击 PDF2Excel.app →「打开」（或系统设置中允许）"
