#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PDF2Excel — PDF 转 Excel 桌面应用入口。

GUI 模式（默认）：pywebview 原生窗口渲染 ui/ 下的 Web 界面。
CLI 模式：python main.py --cli input.pdf [-o out.xlsx] [选项]
"""

from __future__ import annotations

import argparse
import base64
import json
import os
import re
import subprocess
import sys
import tempfile
import threading
import uuid

from converter import ConversionCancelled, ConversionError, convert, probe_pdf, unique_output_path

APP_TITLE = "PDF2Excel · PDF 转 Excel"
SETTINGS_DIR = os.path.join(os.path.expanduser("~"), ".pdf2excel")
SETTINGS_PATH = os.path.join(SETTINGS_DIR, "settings.json")
DEFAULT_OUT_DIR = os.path.join(os.path.expanduser("~"), "Downloads", "PDF2Excel")
DROP_DIR = os.path.join(tempfile.gettempdir(), "pdf2excel_drops")
MAX_DROP_SIZE = 300 * 1024 * 1024  # 300MB


def resource_path(rel: str) -> str:
    """兼容开发环境与 PyInstaller 打包（sys._MEIPASS）的资源路径。"""
    root = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.abspath(__file__))
    return os.path.join(root, *rel.split("/"))


# ------------------------------------------------------------------ 设置持久化


def load_settings() -> dict:
    try:
        with open(SETTINGS_PATH, encoding="utf-8") as f:
            s = json.load(f)
        return s if isinstance(s, dict) else {}
    except Exception:  # noqa: BLE001
        return {}


def save_settings(**kv) -> None:
    s = load_settings()
    s.update({k: v for k, v in kv.items() if v is not None})
    try:
        os.makedirs(SETTINGS_DIR, exist_ok=True)
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(s, f, ensure_ascii=False, indent=2)
    except Exception:  # noqa: BLE001
        pass


# ------------------------------------------------------------------ JS API


class Api:
    """暴露给前端 JS 的桥接接口（pywebview 自动挂载为 window.pywebview.api）。"""

    def __init__(self):
        self.window = None  # 由 run_gui() 注入
        self._cancel = threading.Event()
        self._worker: threading.Thread | None = None
        self._drops: dict[str, dict] = {}
        os.makedirs(DROP_DIR, exist_ok=True)

    # ---------- 事件推送（worker 线程 → 前端） ----------
    def _emit(self, event: str, data) -> None:
        if self.window is None:
            return
        payload = json.dumps({"event": event, "data": data}, ensure_ascii=False)
        try:
            self.window.evaluate_js(f"window.__bridge && window.__bridge({payload})")
        except Exception:  # noqa: BLE001
            pass

    # ---------- 基础 ----------
    def get_defaults(self) -> dict:
        s = load_settings()
        return {
            "outputDir": s.get("outputDir") or DEFAULT_OUT_DIR,
            "mode": s.get("mode", "auto"),
            "sheetMode": s.get("sheetMode", "merge"),
            "pages": s.get("pages", ""),
        }

    def save_settings(self, outputDir=None, mode=None, sheetMode=None, pages=None):
        save_settings(outputDir=outputDir, mode=mode, sheetMode=sheetMode, pages=pages)
        return True

    # ---------- 文件选择 ----------
    def select_pdfs(self) -> list[dict]:
        import webview  # 延迟导入，CLI 模式无需加载 GUI 依赖

        try:
            paths = self.window.create_file_dialog(
                webview.OPEN_DIALOG,
                allow_multiple=True,
                file_types=("PDF 文件 (*.pdf)", "所有文件 (*.*)"),
            )
        except Exception:  # noqa: BLE001
            return []
        if not paths:
            return []
        if isinstance(paths, str):
            paths = [paths]
        return [probe_pdf(p) for p in paths if str(p).lower().endswith(".pdf")]

    def select_output_dir(self) -> str | None:
        import webview

        try:
            d = self.window.create_file_dialog(webview.FOLDER_DIALOG)
        except Exception:  # noqa: BLE001
            return None
        if not d:
            return None
        return d[0] if isinstance(d, (list, tuple)) else str(d)

    # ---------- 拖拽文件（WebView 不暴露本地路径，经 base64 分块传输） ----------
    _SAFE_NAME = re.compile(r"[^0-9A-Za-z_.\-\u4e00-\u9fff （）()\[\]【】]+")

    def drop_begin(self, name: str, size: int) -> dict:
        if size > MAX_DROP_SIZE:
            return {"ok": False, "error": "文件过大（超过 300MB），请使用“点击选择文件”"}
        drop_id = uuid.uuid4().hex
        safe = self._SAFE_NAME.sub("_", os.path.basename(name)) or "dropped.pdf"
        path = os.path.join(DROP_DIR, f"{drop_id}_{safe}")
        try:
            self._drops[drop_id] = {"path": path, "fh": open(path, "wb")}
        except OSError as e:
            return {"ok": False, "error": f"无法创建临时文件：{e}"}
        return {"ok": True, "dropId": drop_id}

    def drop_chunk(self, drop_id: str, b64: str) -> dict:
        d = self._drops.get(drop_id)
        if d is None:
            return {"ok": False, "error": "无效的传输会话"}
        try:
            d["fh"].write(base64.b64decode(b64))
            return {"ok": True}
        except Exception as e:  # noqa: BLE001
            self._drop_abort(drop_id)
            return {"ok": False, "error": str(e)}

    def drop_finish(self, drop_id: str) -> dict:
        d = self._drops.pop(drop_id, None)
        if d is None:
            return {"ok": False, "error": "无效的传输会话"}
        try:
            d["fh"].close()
        except Exception:  # noqa: BLE001
            pass
        info = probe_pdf(d["path"])
        info["ok"] = not info["error"] and not info["encrypted"]
        return info

    def _drop_abort(self, drop_id: str) -> None:
        d = self._drops.pop(drop_id, None)
        if d is None:
            return
        try:
            d["fh"].close()
            os.remove(d["path"])
        except Exception:  # noqa: BLE001
            pass

    # ---------- 转换 ----------
    def start_conversion(self, payload: dict) -> dict:
        if self._worker is not None and self._worker.is_alive():
            return {"ok": False, "error": "已有转换任务正在进行"}
        files = [f for f in (payload or {}).get("files", []) if f.get("path")]
        if not files:
            return {"ok": False, "error": "没有待转换的文件"}
        out_dir = payload.get("outputDir") or DEFAULT_OUT_DIR
        try:
            os.makedirs(out_dir, exist_ok=True)
        except OSError as e:
            return {"ok": False, "error": f"输出目录不可用：{e}"}
        opts = {
            "mode": payload.get("mode") or "auto",
            "sheet_mode": payload.get("sheetMode") or "merge",
            "pages": payload.get("pages") or "",
        }
        save_settings(outputDir=out_dir, **opts)
        self._cancel.clear()
        self._worker = threading.Thread(target=self._run_batch, args=(files, out_dir, opts), daemon=True)
        self._worker.start()
        return {"ok": True}

    def _run_batch(self, files: list[dict], out_dir: str, opts: dict) -> None:
        cancelled = False
        for f in files:
            if self._cancel.is_set():
                cancelled = True
                break
            fid = f.get("id")
            self._emit("file_start", {"id": fid})
            try:
                stem = os.path.splitext(os.path.basename(f["path"]))[0]
                xlsx = unique_output_path(out_dir, stem)
                res = convert(
                    f["path"],
                    xlsx_path=xlsx,
                    password=f.get("password") or None,
                    progress_cb=lambda done, total, _id=fid: self._emit(
                        "file_progress", {"id": _id, "done": done, "total": total}
                    ),
                    cancel_event=self._cancel,
                    **opts,
                )
                self._emit(
                    "file_done",
                    {
                        "id": fid,
                        "output": res["output"],
                        "pages": res["pages"],
                        "tables": res["tables"],
                        "rows": res["rows"],
                        "warnings": res["warnings"],
                    },
                )
            except ConversionCancelled:
                cancelled = True
                self._emit("file_error", {"id": fid, "code": "cancelled", "error": "已取消"})
                break
            except ConversionError as e:
                self._emit("file_error", {"id": fid, "code": e.code, "error": str(e)})
            except Exception as e:  # noqa: BLE001 —— 兜底，避免 worker 静默崩溃
                self._emit("file_error", {"id": fid, "code": "unknown", "error": f"发生未预期的错误：{e}"})
        self._emit("batch_done", {"cancelled": cancelled})

    def cancel_conversion(self) -> bool:
        self._cancel.set()
        return True

    # ---------- 系统 ----------
    def open_path(self, path: str) -> bool:
        """在系统文件管理器中定位文件。"""
        try:
            if sys.platform == "darwin":
                subprocess.Popen(["open", "-R", path])
            elif os.name == "nt":
                subprocess.Popen(["explorer", f"/select,{os.path.normpath(path)}"])
            else:
                subprocess.Popen(["xdg-open", os.path.dirname(path) or path])
            return True
        except Exception:  # noqa: BLE001
            return False


# ------------------------------------------------------------------ 入口


def run_gui() -> None:
    import webview

    api = Api()
    window = webview.create_window(
        APP_TITLE,
        resource_path("ui/index.html"),
        js_api=api,
        width=980,
        height=700,
        min_size=(840, 620),
        background_color="#F6F7FB",
    )
    api.window = window
    # 本地资源经内置 HTTP 服务加载，规避 WKWebView / WebView2 的本地文件跨资源限制
    webview.start(http_server=True, debug=False)


def run_cli(argv: list[str]) -> int:
    p = argparse.ArgumentParser(prog="pdf2excel", description="PDF → Excel 转换工具（无界面模式）")
    p.add_argument("input", help="输入 PDF 文件路径")
    p.add_argument("-o", "--output", default=None, help="输出 xlsx 路径（默认：PDF 同目录同名 .xlsx）")
    p.add_argument("--mode", choices=["auto", "lines", "text"], default="auto",
                   help="表格提取模式：auto 自动识别（默认）/ lines 线框表格 / text 无框表格")
    p.add_argument("--pages", default="", help="页码范围，如 1-5,8（默认全部页）")
    p.add_argument("--sheet", choices=["merge", "pages"], default="merge",
                   help="merge 所有页合并到一个工作表（默认）/ pages 每页一个工作表")
    p.add_argument("--password", default=None, help="加密 PDF 的打开密码")
    args = p.parse_args(argv)

    def progress(done: int, total: int) -> None:
        print(f"\r  进度：{done}/{total} 页", end="", flush=True)

    try:
        res = convert(
            args.input,
            xlsx_path=args.output,
            mode=args.mode,
            pages=args.pages,
            sheet_mode=args.sheet,
            password=args.password,
            progress_cb=progress,
        )
    except ConversionError as e:
        print(f"\n[失败] {e}", file=sys.stderr)
        return 1
    except ConversionCancelled:
        print("\n[取消]", file=sys.stderr)
        return 1

    print(f"\n[完成] {res['output']}")
    print(f"  共 {res['pages']} 页 · {res['tables']} 个表格 · {res['rows']} 行")
    for w in res["warnings"]:
        print(f"  提示：{w}")
    return 0


def main() -> None:
    if len(sys.argv) > 1 and sys.argv[1] == "--cli":
        sys.exit(run_cli(sys.argv[2:]))
    run_gui()


if __name__ == "__main__":
    main()
