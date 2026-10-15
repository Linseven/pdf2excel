#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""GUI 冒烟测试：真实 pywebview 窗口 + 真实页面逻辑 + 真实转换流程。

仅覆写「文件选择对话框 / 输出目录对话框」两个无法自动化的原生交互，
其余全部走真实代码路径（app.js 渲染 + Api 桥接 + converter 转换）。

运行：.venv/bin/python tests/gui_smoke.py
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import webview  # noqa: E402

import main as m  # noqa: E402
from converter import probe_pdf  # noqa: E402

SAMPLES = os.path.join(os.path.dirname(os.path.abspath(__file__)), "samples")
RESULTS: list[tuple[str, bool, str]] = []


def check(name: str, cond, detail: str = "") -> None:
    RESULTS.append((name, bool(cond), str(detail)))
    print(f"  {'✓' if cond else '✗'} {name}" + (f"  [{detail}]" if detail and not cond else ""))


def main() -> int:
    out_dir = tempfile.mkdtemp(prefix="p2x_gui_test_")
    api = m.Api()
    window = webview.create_window(
        "GUI 冒烟测试", m.resource_path("ui/index.html"),
        js_api=api, width=980, height=700, hidden=True,
    )
    api.window = window

    def js(code: str):
        return window.evaluate_js(code)

    def run() -> None:
        try:
            # 0) 等待 pywebview API 注入完成
            t0 = time.time()
            while time.time() - t0 < 20:
                if js("!!(window.pywebview && window.pywebview.api)"):
                    break
                time.sleep(0.3)
            check("pywebview API 注入", js("!!(window.pywebview && window.pywebview.api)"))
            time.sleep(1.0)  # 等 app.js init 完成

            check("初始无文件卡片", js("document.querySelectorAll('.file-card').length") == 0)
            check("初始转换按钮禁用", js("document.getElementById('btnConvert').disabled") is True)

            # 1) 覆写文件选择对话框 → 返回真实样例 PDF（含加密件；multipage 拉长转换时长便于观察运行态）
            samples = [probe_pdf(os.path.join(SAMPLES, "multipage.pdf")),
                       probe_pdf(os.path.join(SAMPLES, "encrypted.pdf"))]
            payload = json.dumps(samples, ensure_ascii=False)
            js(f"window.pywebview.api.select_pdfs = () => Promise.resolve({payload})")
            js(f"window.pywebview.api.select_output_dir = () => Promise.resolve({json.dumps(out_dir)})")
            js("document.getElementById('btnOutDir').click()")
            time.sleep(0.6)
            check("输出目录已设置", os.path.basename(str(js("document.getElementById('outDirText').textContent")))
                  == os.path.basename(out_dir))

            # 2) 添加文件（真实 addFiles 逻辑）
            js("document.getElementById('dropzone').click()")
            time.sleep(1.0)
            check("生成 2 个文件卡片", js("document.querySelectorAll('.file-card').length") == 2)
            check("普通文件显示等待转换",
                  js("document.querySelector('.file-card:not(.needs-pw) .badge').textContent") == "等待转换")
            check("加密文件显示需要密码",
                  js("document.querySelector('.file-card.needs-pw .badge').textContent") == "需要密码")
            check("转换按钮可用", js("document.getElementById('btnConvert').disabled") is False)

            # 3) 输入密码（真实密码行逻辑）
            js("document.querySelector('.pw-row input').value = 'secret123';"
               "document.querySelector('.pw-ok').click()")
            time.sleep(0.5)
            check("密码确认后卡片进入等待", js("document.querySelectorAll('.pw-row').length") == 0)
            check("按钮计数 2 个文件",
                  "2 个" in str(js("document.getElementById('ctaLabel').textContent")))

            # 4) 开始转换（真实 start_conversion → converter → 事件 → 渲染）
            # click() 同步执行 handler 中 state.running=true + render()（首个 await 之前），
            # 因此在同一段 JS 里即可确定性地读到按钮的取消状态
            r = js("document.getElementById('btnConvert').click();"
                   "document.getElementById('btnConvert').classList.contains('is-running')")
            check("点击后按钮立即进入取消状态", r is True, str(r))

            saw_run = False
            t0 = time.time()
            done = False
            while time.time() - t0 < 90:
                if not saw_run and js("document.querySelectorAll('.file-card.is-run').length") >= 1:
                    saw_run = True
                if "全部完成" in str(js("document.getElementById('summary').textContent")):
                    done = True
                    break
                time.sleep(0.05)
            check("批量转换完成（真实转换）", done)
            check("观察到转换中的卡片", saw_run)
            time.sleep(0.5)

            badges = js("Array.from(document.querySelectorAll('.badge')).map(b => b.textContent)")
            check("全部卡片已完成", badges == ["已完成", "已完成"], str(badges))
            check("每卡均有打开按钮", js("document.querySelectorAll('.fc-open').length") == 2)
            check("完成行含表格统计",
                  "个表格" in str(js("document.querySelector('.ok-line').textContent")))
            xlsx = [f for f in os.listdir(out_dir) if f.endswith(".xlsx")]
            check("真实生成 2 个 xlsx", len(xlsx) == 2, str(xlsx))

            # 5) 移除文件（真实 removeFile 逻辑）
            js("document.querySelector('.fc-remove').click()")
            time.sleep(0.4)
            check("移除后剩 1 个卡片", js("document.querySelectorAll('.file-card').length") == 1)
        except Exception as e:  # noqa: BLE001
            check("冒烟测试执行", False, repr(e))
        finally:
            # 还原全局设置中的输出目录，避免污染用户配置
            m.save_settings(outputDir=m.DEFAULT_OUT_DIR)
            window.destroy()

    webview.start(run, http_server=True, debug=False)

    failed = [r for r in RESULTS if not r[1]]
    print(f"\n结果：{len(RESULTS) - len(failed)} 通过 / {len(failed)} 失败")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
