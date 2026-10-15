/* PDF2Excel 前端逻辑：文件管理 / 拖放传输 / 转换流程 / 状态渲染 */
"use strict";

(() => {
  const $ = (s) => document.querySelector(s);

  const els = {
    dropzone: $("#dropzone"),
    btnBrowse: $("#btnBrowse"),
    selMode: $("#selMode"),
    selSheet: $("#selSheet"),
    inpPages: $("#inpPages"),
    btnOutDir: $("#btnOutDir"),
    outDirText: $("#outDirText"),
    btnClear: $("#btnClear"),
    fileList: $("#fileList"),
    emptyState: $("#emptyState"),
    fileCount: $("#fileCount"),
    summary: $("#summary"),
    btnConvert: $("#btnConvert"),
    ctaLabel: $("#ctaLabel"),
    dropOverlay: $("#dropOverlay"),
    toast: $("#toast"),
  };

  const state = {
    files: new Map(),   // id → file 对象
    outputDir: "",
    running: false,
    toastTimer: null,
  };

  /* ---------------------------------------------------------- 小工具 */

  const esc = (s) =>
    String(s ?? "").replace(/[&<>"']/g, (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  const fmtSize = (n) => {
    if (!n && n !== 0) return "";
    if (n < 1024) return n + " B";
    if (n < 1048576) return (n / 1024).toFixed(1) + " KB";
    if (n < 1073741824) return (n / 1048576).toFixed(1) + " MB";
    return (n / 1073741824).toFixed(2) + " GB";
  };

  const genId = () =>
    (crypto.randomUUID ? crypto.randomUUID() : "f" + Date.now() + Math.random().toString(36).slice(2));

  function toast(msg, isError = false) {
    els.toast.textContent = msg;
    els.toast.classList.toggle("error", isError);
    els.toast.classList.remove("hidden");
    clearTimeout(state.toastTimer);
    state.toastTimer = setTimeout(() => els.toast.classList.add("hidden"), 3200);
  }

  /* ---------------------------------------------------------- 图标 */

  const ICONS = {
    pdf: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"
      stroke-linecap="round" stroke-linejoin="round">
      <path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8l-5-5z"/>
      <path d="M14 3v5h5M8.5 13.5h2.2m2.6 0h2.2M8.5 17h2.2m2.6 0h2.2"/></svg>`,
    x: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"
      stroke-linecap="round"><path d="M18 6 6 18M6 6l12 12"/></svg>`,
    folderOpen: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"
      stroke-linecap="round" stroke-linejoin="round">
      <path d="m3 8 4-4 2 2h10a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8z" transform="translate(0 1)"/></svg>`,
  };

  /* ---------------------------------------------------------- 文件管理 */

  function addFiles(infos) {
    let added = 0;
    for (const info of infos) {
      if (!info || !info.path) continue;
      if (info.error) { toast(`无法读取 ${info.name}：${info.error}`, true); continue; }
      // 去重：同一路径且未转换过的不再重复添加
      const dup = [...state.files.values()].find(
        (f) => f.path === info.path && (f.status === "wait" || f.status === "needs-pw"));
      if (dup) continue;
      const id = info.id || genId();  // 只生成一次，Map 键与对象 id 必须一致
      state.files.set(id, {
        id,
        path: info.path,
        name: info.name || "document.pdf",
        size: info.size || 0,
        pages: info.pages,
        encrypted: !!info.encrypted,
        password: null,
        status: info.encrypted ? "needs-pw" : "wait",
        result: null,
        error: null,
      });
      added++;
    }
    if (added) render();
  }

  function removeFile(id) {
    state.files.delete(id);
    render();
  }

  function waitingFiles() {
    return [...state.files.values()].filter((f) => f.status === "wait" || f.status === "needs-pw");
  }

  function convertTargetFiles() {
    // 参与转换的文件：等待中 + 需要密码（已输入密码的）
    return [...state.files.values()].filter(
      (f) => f.status === "wait" || (f.status === "needs-pw" && f.password));
  }

  /* ---------------------------------------------------------- 渲染 */

  function render() {
    els.emptyState.style.display = state.files.size ? "none" : "flex";
    els.btnClear.classList.toggle("hidden", !state.files.size || state.running);

    // 移除 DOM 中已不存在的卡片
    els.fileList.querySelectorAll(".file-card").forEach((el) => {
      if (!state.files.has(el.dataset.id)) el.remove();
    });

    for (const f of state.files.values()) {
      let card = els.fileList.querySelector(`.file-card[data-id="${f.id}"]`);
      if (!card) { card = buildCard(f); els.fileList.appendChild(card); }
      updateCard(f, card);
    }

    const n = state.files.size;
    els.fileCount.textContent = n ? `· ${n} 个` : "";
    updateActionbar();
  }

  function buildCard(f) {
    const card = document.createElement("div");
    card.className = "file-card";
    card.dataset.id = f.id;
    card.innerHTML = `
      <div class="fc-icon">${ICONS.pdf}</div>
      <div class="fc-body">
        <div class="fc-name"></div>
        <div class="fc-meta"></div>
        <div class="fc-progress"><div class="fc-bar"></div></div>
        <div class="fc-extra"></div>
      </div>
      <div class="fc-side">
        <span class="badge"></span>
        <button class="icon-btn fc-remove" title="移除">${ICONS.x}</button>
      </div>`;
    card.querySelector(".fc-remove").addEventListener("click", () => {
      if (state.running) return;
      removeFile(f.id);
    });
    return card;
  }

  function updateCard(f, card) {
    card.className = "file-card is-" + (f.status === "needs-pw" ? "wait" : f.status);
    card.classList.toggle("needs-pw", f.status === "needs-pw");

    card.querySelector(".fc-name").textContent = f.name;
    card.querySelector(".fc-name").title = f.path;

    const meta = [];
    if (f.size) meta.push(fmtSize(f.size));
    if (f.status === "run") {
      meta.push(f.progress ? `正在转换 ${f.progress.done}/${f.progress.total} 页` : "正在读取…");
    } else if (f.pages) {
      meta.push(`${f.pages} 页`);
    } else if (f.encrypted) {
      meta.push("已加密");
    }
    card.querySelector(".fc-meta").textContent = meta.join(" · ");

    if (f.status === "run") {
      const pct = f.progress && f.progress.total ? (f.progress.done / f.progress.total) * 100 : 0;
      card.querySelector(".fc-bar").style.width = Math.max(pct, 4) + "%";
    }

    const badge = card.querySelector(".badge");
    const extra = card.querySelector(".fc-extra");
    const removeBtn = card.querySelector(".fc-remove");
    removeBtn.disabled = state.running;

    if (f.status === "wait") {
      badge.textContent = "等待转换";
      extra.innerHTML = "";
    } else if (f.status === "needs-pw") {
      badge.textContent = "需要密码";
      if (!extra.querySelector(".pw-row")) {
        extra.innerHTML = "";
        const row = document.createElement("div");
        row.className = "pw-row";
        row.innerHTML = `
          <input type="password" placeholder="输入 PDF 打开密码" spellcheck="false">
          <button class="mini-btn pw-ok">确定</button>`;
        row.querySelector("input").addEventListener("keydown", (e) => {
          if (e.key === "Enter") row.querySelector(".pw-ok").click();
        });
        row.querySelector(".pw-ok").addEventListener("click", () => {
          const v = row.querySelector("input").value.trim();
          if (!v) { toast("请先输入密码", true); return; }
          f.password = v;
          f.status = "wait";
          render();
        });
        extra.appendChild(row);
      }
    } else if (f.status === "run") {
      badge.textContent = "转换中";
      extra.innerHTML = "";
    } else if (f.status === "done") {
      badge.textContent = "已完成";
      const out = f.result ? f.result.output : "";
      let html = `<span class="ok-line">已保存 <b>${esc(shortName(out))}</b>`;
      if (f.result && f.result.tables != null) {
        html += `<span>· ${f.result.tables} 个表格 / ${f.result.rows} 行</span>`;
      }
      html += `<button class="mini-btn fc-open">打开</button></span>`;
      if (f.result && f.result.warnings && f.result.warnings.length) {
        const w = f.result.warnings;
        html += `<span class="warn-line" title="${esc(w.join("\n"))}">提示：${esc(w[0])}${w.length > 1 ? ` 等 ${w.length} 条` : ""}</span>`;
      }
      extra.innerHTML = html;
      const btn = extra.querySelector(".fc-open");
      if (btn) btn.addEventListener("click", () => pywebview.api.open_path(out));
    } else if (f.status === "fail") {
      badge.textContent = f.errorCode === "cancelled" ? "已取消" : "失败";
      let html = `<span class="err-line">${esc(f.error || "转换失败")}</span>`;
      if (f.errorCode === "password") {
        f.status = "needs-pw";
        return updateCard(f, card); // 复用密码输入 UI
      }
      extra.innerHTML = html;
    }
  }

  function shortName(path) {
    if (!path) return "";
    const parts = String(path).replace(/\\/g, "/").split("/");
    return parts[parts.length - 1] || path;
  }

  function updateActionbar() {
    const targets = convertTargetFiles();
    const n = state.files.size;

    if (state.running) {
      els.btnConvert.classList.add("is-running");
      els.btnConvert.disabled = false;
      els.ctaLabel.textContent = "取消转换";
      const run = [...state.files.values()].filter((f) => f.status === "run");
      els.summary.textContent = run.length
        ? `正在转换：${run[0].name}`
        : "正在处理…";
      return;
    }

    els.btnConvert.classList.remove("is-running");
    els.ctaLabel.textContent = targets.length ? `开始转换${targets.length > 1 ? `（${targets.length} 个）` : ""}` : "开始转换";

    if (!n) {
      els.btnConvert.disabled = true;
      els.summary.textContent = "准备就绪";
    } else if (!targets.length) {
      els.btnConvert.disabled = true;
      const done = [...state.files.values()].filter((f) => f.status === "done").length;
      els.summary.textContent = done === n ? `全部完成 · ${n} 个文件` : "没有可转换的文件（请检查文件状态）";
    } else {
      els.btnConvert.disabled = false;
      const pages = targets.reduce((s, f) => s + (f.pages || 0), 0);
      els.summary.textContent = `${n} 个文件 · ${pages ? `${pages} 页 · ` : ""}输出至 ${state.outputDir || "默认目录"}`;
    }
  }

  /* ---------------------------------------------------------- Python 事件 */

  const handlers = {
    file_start(d) {
      const f = state.files.get(d.id);
      if (!f) return;
      f.status = "run"; f.progress = null;
      updateCard(f, cardOf(f.id));
      updateActionbar();
    },
    file_progress(d) {
      const f = state.files.get(d.id);
      if (!f || f.status !== "run") return;
      f.progress = d;
      const card = cardOf(d.id);
      if (card) {
        card.querySelector(".fc-meta").textContent =
          `${fmtSize(f.size)} · 正在转换 ${d.done}/${d.total} 页`;
        card.querySelector(".fc-bar").style.width = Math.max((d.done / d.total) * 100, 4) + "%";
      }
    },
    file_done(d) {
      const f = state.files.get(d.id);
      if (!f) return;
      f.status = "done"; f.result = d; f.progress = null;
      updateCard(f, cardOf(f.id));
      updateActionbar();
    },
    file_error(d) {
      const f = state.files.get(d.id);
      if (!f) return;
      f.status = "fail"; f.error = d.error; f.errorCode = d.code; f.progress = null;
      updateCard(f, cardOf(f.id));
      updateActionbar();
    },
    batch_done(d) {
      state.running = false;
      const files = [...state.files.values()];
      const done = files.filter((f) => f.status === "done").length;
      const fail = files.filter((f) => f.status === "fail" && f.errorCode !== "cancelled").length;
      if (d && d.cancelled) {
        els.summary.textContent = "已取消";
        toast("转换已取消");
      } else if (fail === 0) {
        toast(`全部完成：${done} 个文件已转换`);
      } else {
        toast(`完成 ${done} 个，失败 ${fail} 个`, true);
      }
      render();
    },
  };

  window.__bridge = (msg) => {
    try {
      const h = handlers[msg && msg.event];
      if (h) h(msg.data);
    } catch (e) {
      console.error(e);
    }
  };

  const cardOf = (id) => els.fileList.querySelector(`.file-card[data-id="${id}"]`);

  /* ---------------------------------------------------------- 拖放（分块 base64 传输） */

  function b64encode(u8) {
    let s = "";
    const CH = 0x8000;
    for (let i = 0; i < u8.length; i += CH) {
      s += String.fromCharCode.apply(null, u8.subarray(i, i + CH));
    }
    return btoa(s);
  }

  async function uploadDroppedFiles(webFiles) {
    for (const f of webFiles) {
      if (!/\.pdf$/i.test(f.name)) { toast(`已跳过非 PDF 文件：${f.name}`, true); continue; }
      if (f.size > 300 * 1024 * 1024) { toast(`文件过大（>300MB）：${f.name}`, true); continue; }
      try {
        const begin = await pywebview.api.drop_begin(f.name, f.size);
        if (!begin.ok) { toast(begin.error || "无法接收文件", true); continue; }
        const buf = new Uint8Array(await f.arrayBuffer());
        const CHUNK = 512 * 1024;
        for (let off = 0; off < buf.length; off += CHUNK) {
          const r = await pywebview.api.drop_chunk(begin.dropId, b64encode(buf.subarray(off, off + CHUNK)));
          if (!r.ok) { toast(r.error || "传输失败", true); break; }
        }
        const info = await pywebview.api.drop_finish(begin.dropId);
        if (!info.ok && info.error) toast(`${f.name}：${info.error}`, true);
        addFiles([info]);
      } catch (e) {
        toast(`读取 ${f.name} 失败`, true);
      }
    }
  }

  // 全窗口拖放
  let dragDepth = 0;
  document.addEventListener("dragenter", (e) => {
    if (!e.dataTransfer || ![...e.dataTransfer.types].includes("Files")) return;
    e.preventDefault();
    dragDepth++;
    els.dropOverlay.classList.remove("hidden");
  });
  document.addEventListener("dragover", (e) => e.preventDefault());
  document.addEventListener("dragleave", (e) => {
    e.preventDefault();
    if (--dragDepth <= 0) { dragDepth = 0; els.dropOverlay.classList.add("hidden"); }
  });
  document.addEventListener("drop", (e) => {
    e.preventDefault();
    dragDepth = 0;
    els.dropOverlay.classList.add("hidden");
    if (state.running) { toast("转换进行中，请稍候再添加文件", true); return; }
    if (e.dataTransfer && e.dataTransfer.files.length) uploadDroppedFiles(e.dataTransfer.files);
  });

  /* ---------------------------------------------------------- 交互绑定 */

  els.btnBrowse.addEventListener("click", (e) => { e.stopPropagation(); pickFiles(); });
  els.dropzone.addEventListener("click", pickFiles);
  els.dropzone.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); pickFiles(); }
  });

  async function pickFiles() {
    if (state.running) { toast("转换进行中，请稍候再添加文件", true); return; }
    const infos = await pywebview.api.select_pdfs();
    if (infos && infos.length) addFiles(infos);
  }

  els.btnClear.addEventListener("click", () => {
    state.files.clear();
    render();
  });

  els.btnOutDir.addEventListener("click", async () => {
    const d = await pywebview.api.select_output_dir();
    if (d) { state.outputDir = d; els.outDirText.textContent = d; els.btnOutDir.title = d; updateActionbar(); }
  });

  els.btnConvert.addEventListener("click", async () => {
    if (state.running) {
      els.ctaLabel.textContent = "正在取消…";
      els.btnConvert.disabled = true;
      await pywebview.api.cancel_conversion();
      return;
    }
    const files = convertTargetFiles();
    if (!files.length) return;

    const missingPw = [...state.files.values()].filter((f) => f.status === "needs-pw" && !f.password);
    if (missingPw.length) toast(`${missingPw.length} 个加密文件尚未输入密码，本次将跳过`);

    state.running = true;
    render(); // 冻结列表操作
    const r = await pywebview.api.start_conversion({
      outputDir: state.outputDir,
      mode: els.selMode.value,
      sheetMode: els.selSheet.value,
      pages: els.inpPages.value.trim(),
      files: files.map((f) => ({ id: f.id, path: f.path, password: f.password || null })),
    });
    if (!r || !r.ok) {
      state.running = false;
      toast((r && r.error) || "无法开始转换", true);
      render();
    }
  });

  /* ---------------------------------------------------------- 初始化 */

  window.addEventListener("pywebviewready", async () => {
    try {
      const d = await pywebview.api.get_defaults();
      state.outputDir = d.outputDir || "";
      els.outDirText.textContent = state.outputDir;
      els.btnOutDir.title = state.outputDir;
      els.selMode.value = d.mode || "auto";
      els.selSheet.value = d.sheetMode || "merge";
      els.inpPages.value = d.pages || "";
    } catch (e) {
      console.error(e);
    }
    render();
  });

  render(); // pywebview 未就绪时先渲染静态骨架
})();
