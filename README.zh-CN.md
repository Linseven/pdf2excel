<div align="center">

<img src="assets/icon.png" alt="PDF2Excel" width="96">

# PDF2Excel

**PDF 表格转 Excel —— 精美、精准、完全本地运行。**

[English](README.md) | 简体中文

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-macOS%20%7C%20Windows-lightgrey)](#下载)
[![Python](https://img.shields.io/badge/python-3.10%2B-3776AB?logo=python&logoColor=white)](#从源码构建)
[![Release](https://img.shields.io/github/v/release/Linseven/pdf2excel)](https://github.com/Linseven/pdf2excel/releases)

<img src="docs/screenshots/app.png" alt="PDF2Excel 界面截图" width="820">

</div>

## 功能特性

- **精准表格提取** —— 三种模式：自动识别（逐页在线框/无框策略间回退）、线框表格、无框表格，基于 [pdfplumber](https://github.com/jsvine/pdfplumber)
- **批量转换** —— 拖拽添加多个 PDF，页级进度条、随时取消，单个失败不阻塞队列
- **页码范围** —— 只转需要的页，如 `1-5,8,11-13`
- **工作表模式** —— 所有页合并为一个工作表，或每页一个工作表
- **加密 PDF** —— 界面内按文件输入密码，无需借助其他工具
- **智能回退** —— 无表格页自动提取纯文本；扫描件（纯图片页）给出明确提示
- **精美 Excel 输出** —— 表头加粗底纹、冻结首行、列宽自适应（中文双宽感知）、细边框
- **100% 本地离线** —— 文件不出本机，无需账号、绝不上传
- **内置命令行** —— 无界面批处理，便于脚本集成

## 下载

从 [**Releases**](https://github.com/Linseven/pdf2excel/releases) 获取最新构建：

| 平台 | 文件 | 备注 |
|---|---|---|
| macOS（Apple Silicon） | `PDF2Excel-macOS-arm64.zip` | 首次打开：右键 → **打开** |
| Windows 10/11 | `PDF2Excel.exe` | 本机一键构建，见 [docs/BUILD_WINDOWS.md](docs/BUILD_WINDOWS.md) |

> Intel Mac 与 Windows 预编译包暂未发布，可按下文一键从源码构建。

## 从源码构建

需要 **Python 3.10+**。

**macOS**

```bash
git clone https://github.com/Linseven/pdf2excel.git
cd pdf2excel
bash build_mac.sh          # → dist/PDF2Excel.app
```

**Windows**

```bat
git clone https://github.com/Linseven/pdf2excel.git
cd pdf2excel
build_windows.bat          :: → dist\PDF2Excel.exe
```

两个脚本都会自动创建虚拟环境、安装依赖、生成图标并用 PyInstaller 打包，无需任何手动步骤。

## 免打包运行（开发模式）

```bash
python3 -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements-dev.txt
python main.py
```

## 命令行用法

```bash
python main.py --cli input.pdf -o output.xlsx
python main.py --cli report.pdf --mode text --pages 1-5,8 --sheet pages
python main.py --cli secret.pdf --password 123456
```

| 参数 | 取值 | 默认值 | 说明 |
|---|---|---|---|
| `-o, --output` | 路径 | 同名 `.xlsx` | 输出文件 |
| `--mode` | `auto` / `lines` / `text` | `auto` | 表格提取策略 |
| `--pages` | 如 `1-5,8` | 全部 | 页码范围 |
| `--sheet` | `merge` / `pages` | `merge` | 工作表布局 |
| `--password` | 字符串 | — | 加密 PDF 的打开密码 |

## 工作原理

```
PDF ──▶ pdfplumber ──▶ 表格单元格（线框 / 文字对齐）
      ──▶ 纯文本回退（未检测到表格）
      ──▶ 扫描件提示（无文字层）
                          │
                          ▼
              openpyxl ──▶ 带样式的 .xlsx
```

界面采用原生 [pywebview](https://pywebview.flowrl.com/) 窗口渲染手工打造的零依赖 HTML/CSS/JS 界面 —— 拥有 Electron 级的观感，体积却只有零头。

## 项目结构

```
├── main.py              # 入口：pywebview 桥接（Api）+ CLI
├── converter.py         # 纯转换引擎（无 GUI 依赖）
├── ui/                  # 前端（HTML/CSS/JS，零构建步骤）
├── assets/              # 图标（tools/make_icons.py 生成）
├── tools/               # 图标生成器、截图工具
├── tests/               # 引擎端到端测试 + 真实窗口 GUI 冒烟测试
├── build_mac.sh         # macOS 一键构建 → .app
└── build_windows.bat    # Windows 一键构建 → .exe
```

## 测试

```bash
python tests/run_tests.py    # 24 项引擎断言（线框/无框/加密/多页/文本回退…）
python tests/gui_smoke.py    # 18 项真实窗口 GUI 断言
```

## 路线图

- [ ] 扫描件 OCR 支持
- [ ] Releases 提供 Windows 预编译安装包
- [ ] 英文界面切换
- [ ] CSV / Markdown 导出

## 参与贡献

欢迎 Issue 与 PR —— 参见 [CONTRIBUTING.md](CONTRIBUTING.md)。

## 开源协议

[MIT](LICENSE) © 2026 Linseven
