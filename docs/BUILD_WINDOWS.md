# Windows 构建说明

`build_windows.bat` 一键将本项目打包为 Windows 平台可直接运行的独立程序 **`dist\PDF2Excel.exe`**（单文件、双击即用、无需安装 Python）。

---

## 一、前置条件

| 条件 | 说明 |
|---|---|
| Windows 10 / 11 | 64 位系统（Windows 7/8 未验证） |
| Python 3.10 – 3.14 | 从 [python.org/downloads](https://www.python.org/downloads/) 安装，**安装时务必勾选 “Add python.exe to PATH”** |
| 网络 | 首次构建需联网下载依赖（约 100MB） |
| 磁盘 | 至少 1GB 可用空间（虚拟环境 + 构建缓存 + 产物） |

> 注意：PyInstaller 不支持跨平台编译，**必须在 Windows 机器上运行本脚本**（macOS 上构建的是 .app，由 `build_mac.sh` 负责）。

---

## 二、使用方法

1. 将整个项目文件夹复制到 Windows 电脑（U盘 / Git / 压缩包均可，**无需复制 `.venv/`、`build/`、`dist/` 目录**）。
2. 双击 `build_windows.bat`，或在命令行中执行：
   ```bat
   cd /d 项目路径
   build_windows.bat
   ```
3. 等待构建完成（首次约 3–10 分钟，取决于网速；之后重建约 1–3 分钟）。
4. 构建成功后，产物位于：**`dist\PDF2Excel.exe`**，可复制到任意位置运行或分发给他人。

---

## 三、脚本执行流程

双击脚本后，它按顺序完成以下步骤，任一步失败都会提示原因并暂停（不会静默出错）：

| 步骤 | 动作 | 说明 |
|---|---|---|
| 1 | 检查 `python` 命令 | 找不到则提示安装 Python 并勾选 PATH |
| 2 | 创建 `.venv` 虚拟环境 | 已存在则跳过，不污染系统 Python |
| 3 | 安装依赖 | `pip install -r requirements-dev.txt`（pdfplumber、openpyxl、pywebview、pyinstaller 等） |
| 4 | 生成图标 | `assets\icon.ico` 不存在时自动生成（首次需 Pillow，已含在依赖中） |
| 5 | PyInstaller 打包 | 单文件 `--onefile` + 无控制台 `--noconsole`，打入 `ui\` 与 `assets\` 资源 |
| 6 | 输出 | `dist\PDF2Excel.exe`（约 40–60MB） |

---

## 四、常见问题

### 1. 双击 exe 后窗口空白 / 闪退
程序界面依赖 **WebView2 运行时**。Windows 10/11 一般已内置；若缺失，从微软官网安装：
https://developer.microsoft.com/microsoft-edge/webview2/
（脚本结束时的提示中也包含此链接）

### 2. 弹出 "Windows 已保护你的电脑"（SmartScreen）
未签名的 exe 会触发此提示。点击 **“更多信息” → “仍要运行”** 即可。正式分发建议购买代码签名证书。

### 3. 杀毒软件误报
PyInstaller 单文件 exe 偶发被误报，属常见现象。可将 `dist\PDF2Excel.exe` 加入白名单；如需彻底规避，把脚本中 `--onefile` 删除后重建，改用 `dist\PDF2Excel\PDF2Excel.exe`（目录版，误报率更低但分发需带整个文件夹）。

### 4. 依赖安装失败
- 多为公司网络/代理所致，可为 pip 配置代理或换源：
  ```bat
  .venv\Scripts\python.exe -m pip config set global.index-url https://pypi.tuna.tsinghua.edu.cn/simple
  ```
  然后重新运行脚本。
- 若 Python 版本过新（刚发布）导致个别包无预编译 wheel，换用 3.12/3.13 稳定版。

### 5. 想清理后全新构建
```bat
rmdir /s /q .venv build dist
build_windows.bat
```

---

## 五、产物验证清单

构建完成后建议快速验证：

1. 双击 `dist\PDF2Excel.exe` → 出现 “PDF2Excel” 主窗口（浅色界面、虚线拖放区）。
2. 拖入一个含表格的 PDF → 显示文件卡片与页数。
3. 点击 “开始转换” → 进度条推进，完成后点击 “打开” 能定位到生成的 `.xlsx`。

如需无界面批量处理，也可使用命令行模式：
```bat
PDF2Excel.exe --cli 输入.pdf -o 输出.xlsx
```

---

## 六、相关文件

| 文件 | 作用 |
|---|---|
| `build_windows.bat` | Windows 一键构建脚本（本文档所述） |
| `build_mac.sh` | macOS 构建脚本（产出 .app） |
| `requirements.txt` | 运行时依赖 |
| `requirements-dev.txt` | 构建 + 测试依赖（含上一文件） |
| `PDF2Excel.spec` | PyInstaller 自动生成的打包配置（重建时会被脚本覆盖） |
