"""
流水线面板 v3
正确顺序：01 文档解析 → 02 数据标注 → 03 MD转Word
支持批量多文件输入（最多10个，每个≤50MB）
"""
import os, threading, traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from gui.theme import COLORS, FONTS, PADDING
from gui.widgets import StyledButton, LogView, Divider

_MAX_FILES        = 10
_MAX_SIZE_MB      = 50
_MAX_SIZE_BYTES   = _MAX_SIZE_MB * 1024 * 1024


def _fmt_size(b: int) -> str:
    return f"{b/1024:.0f} KB" if b < 1024*1024 else f"{b/1024/1024:.1f} MB"


class PipelinePanel(tk.Frame):
    def __init__(self, parent, shared_state, status_bar,
                 switch_to_tool_cb=None, **kw):
        bg = kw.pop("bg", COLORS["bg_main"])
        super().__init__(parent, bg=bg, **kw)
        self._shared        = shared_state
        self._status_bar    = status_bar
        self._switch_cb     = switch_to_tool_cb
        self._is_running    = False
        self._current_files: list = []
        self._build()

    # ── 构建 UI ───────────────────────────────────────────────

    def _build(self):
        # 顶部标题栏（固定，不随滚动）
        hdr = tk.Frame(self, bg=COLORS["bg_card"])
        hdr.pack(fill="x")
        if self._switch_cb:
            tk.Button(hdr, text="◈  主页",
                      bg=COLORS["bg_card"], fg=COLORS["accent"],
                      activebackground=COLORS["bg_hover"],
                      activeforeground=COLORS["accent"],
                      relief="flat", bd=0, cursor="hand2",
                      font=FONTS["sm"], padx=PADDING["md"],
                      command=lambda: self._switch_cb("home")
                      ).pack(side="right", pady=PADDING["sm"],
                             padx=PADDING["md"])
        tk.Label(hdr, text="流水线",
                 bg=COLORS["bg_card"], fg=COLORS["pipeline_color"],
                 font=FONTS["h1"]).pack(side="left",
                                        padx=PADDING["xl"], pady=PADDING["lg"])
        tk.Label(hdr, text="一键串联三个工具，自动完成文档全流程处理",
                 bg=COLORS["bg_card"], fg=COLORS["text_secondary"],
                 font=FONTS["md"]).pack(side="left", pady=PADDING["lg"])
        Divider(self).pack(fill="x")

        # 滚动区域
        canvas = tk.Canvas(self, bg=COLORS["bg_main"], highlightthickness=0)
        sb = ttk.Scrollbar(self, orient="vertical",
                            style="Dark.Vertical.TScrollbar",
                            command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        body = tk.Frame(canvas, bg=COLORS["bg_main"])
        win_id = canvas.create_window((0, 0), window=body, anchor="nw")
        body.bind("<Configure>", lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(
            win_id, width=e.width))

        # 鼠标滚轮 —— 进入面板全局绑定，离开解绑
        import sys as _sys
        def _scroll(event):
            if _sys.platform == "win32":
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            elif _sys.platform == "darwin":
                canvas.yview_scroll(int(-1 * event.delta), "units")
            else:
                canvas.yview_scroll(-1 if event.num == 4 else 1, "units")

        def _on_enter(e):
            if _sys.platform == "linux":
                self.winfo_toplevel().bind_all("<Button-4>", _scroll)
                self.winfo_toplevel().bind_all("<Button-5>", _scroll)
            else:
                self.winfo_toplevel().bind_all("<MouseWheel>", _scroll)

        def _on_leave(e):
            if _sys.platform == "linux":
                self.winfo_toplevel().unbind_all("<Button-4>")
                self.winfo_toplevel().unbind_all("<Button-5>")
            else:
                self.winfo_toplevel().unbind_all("<MouseWheel>")

        self.bind("<Enter>", _on_enter)
        self.bind("<Leave>", _on_leave)

        pad = PADDING["xl"]

        # ── 文件选择 ─────────────────────────────────────────
        self._section(body, "输入文件", pad)
        file_card = tk.Frame(body, bg=COLORS["bg_card"])
        file_card.pack(fill="x", padx=pad, pady=(0, PADDING["lg"]))
        fi = tk.Frame(file_card, bg=COLORS["bg_card"])
        fi.pack(fill="x", padx=PADDING["lg"], pady=PADDING["lg"])

        # 两个选择按钮
        btn_row = tk.Frame(fi, bg=COLORS["bg_card"])
        btn_row.pack(fill="x")
        StyledButton(btn_row, text="📄 选择文件",
                     style="secondary",
                     command=self._pick_files).pack(
            side="left", fill="x", expand=True, padx=(0, 4))
        StyledButton(btn_row, text="📁 选择文件夹",
                     style="secondary",
                     command=self._pick_folder).pack(
            side="left", fill="x", expand=True)

        self._file_count_lbl = tk.Label(
            fi,
            text=f"未选择文件（最多 {_MAX_FILES} 个，每个 ≤ {_MAX_SIZE_MB} MB）",
            bg=COLORS["bg_card"], fg=COLORS["text_muted"],
            font=FONTS["xs"], anchor="w",
        )
        self._file_count_lbl.pack(fill="x", pady=(6, 0))

        # 文件列表区
        list_wrap = tk.Frame(fi, bg=COLORS["bg_active"], pady=2)
        list_wrap.pack(fill="x", pady=(6, 0))
        self._file_list_frame = tk.Frame(list_wrap, bg=COLORS["bg_active"])
        self._file_list_frame.pack(fill="x", padx=2)
        self._list_placeholder = tk.Label(
            self._file_list_frame,
            text="点击上方按钮选择文件或文件夹",
            bg=COLORS["bg_active"], fg=COLORS["text_muted"],
            font=FONTS["xs"], pady=10,
        )
        self._list_placeholder.pack()

        # ── 执行步骤（正确顺序）──────────────────────────────
        self._section(body, "执行步骤", pad)

        self._step1_var = tk.BooleanVar(value=True)
        self._step2_var = tk.BooleanVar(value=True)
        self._step3_var = tk.BooleanVar(value=True)

        self._step_card(body, pad, "01", "文档解析",
                        "PDF / Word  →  Markdown + 图片",
                        COLORS["tool1_color"], self._step1_var)
        self._step_card(body, pad, "02", "数据标注",
                        "Markdown  →  打标签 + 表格转QA",
                        COLORS["tool3_color"], self._step2_var)
        self._step_card(body, pad, "03", "MD → Word",
                        "标注后的 Markdown  →  Word 文档",
                        COLORS["tool2_color"], self._step3_var)

        # ── 按钮行 ───────────────────────────────────────────
        btn_row = tk.Frame(body, bg=COLORS["bg_main"])
        btn_row.pack(padx=pad, pady=PADDING["lg"], anchor="w")

        self._run_btn = StyledButton(btn_row, text="启动流水线",
                                     style="primary", command=self._run)
        self._run_btn.pack(side="left", padx=(0, PADDING["md"]))

        self._stop_btn = StyledButton(btn_row, text="停止",
                                      style="danger", command=self._stop)
        self._stop_btn.pack(side="left")
        self._stop_btn.config(state="disabled")

        # ── 执行进度 ─────────────────────────────────────────
        self._section(body, "执行进度", pad)
        prog_card = tk.Frame(body, bg=COLORS["bg_card"])
        prog_card.pack(fill="x", padx=pad, pady=(0, PADDING["lg"]))
        pi = tk.Frame(prog_card, bg=COLORS["bg_card"])
        pi.pack(fill="x", padx=PADDING["lg"], pady=PADDING["lg"])

        # 当前文件提示
        self._cur_file_lbl = tk.Label(
            pi, text="", bg=COLORS["bg_card"],
            fg=COLORS["pipeline_color"], font=FONTS["xs"], anchor="w")
        self._cur_file_lbl.pack(fill="x", pady=(0, PADDING["sm"]))

        self._p1 = self._prog_row(pi, "Step 01  文档解析",
                                   "Blue.Horizontal.TProgressbar")
        self._p2 = self._prog_row(pi, "Step 02  数据标注",
                                   "Orange.Horizontal.TProgressbar")
        self._p3 = self._prog_row(pi, "Step 03  MD → Word",
                                   "Green.Horizontal.TProgressbar")

        # ── 执行日志 ─────────────────────────────────────────
        self._section(body, "执行日志", pad)
        log_card = tk.Frame(body, bg=COLORS["bg_card"])
        log_card.pack(fill="x", padx=pad, pady=(0, PADDING["xl"]))
        self._log = LogView(log_card, height=10)
        self._log.pack(fill="x")

    # ── 辅助 UI ───────────────────────────────────────────────

    def _section(self, parent, title, pad):
        row = tk.Frame(parent, bg=COLORS["bg_main"])
        row.pack(fill="x", padx=pad, pady=(PADDING["md"], PADDING["sm"]))
        tk.Frame(row, bg=COLORS["pipeline_color"], width=4).pack(
            side="left", fill="y")
        tk.Label(row, text=title, bg=COLORS["bg_main"],
                 fg=COLORS["text_primary"],
                 font=FONTS["h2"], padx=10).pack(side="left")

    def _step_card(self, parent, pad, num, title, desc, color, var):
        card = tk.Frame(parent, bg=COLORS["bg_card"])
        card.pack(fill="x", padx=pad, pady=(0, PADDING["sm"]))
        tk.Frame(card, bg=color, width=5).pack(side="left", fill="y")
        inner = tk.Frame(card, bg=COLORS["bg_card"])
        inner.pack(side="left", fill="both", expand=True,
                   padx=PADDING["lg"], pady=PADDING["md"])
        row = tk.Frame(inner, bg=COLORS["bg_card"])
        row.pack(fill="x")
        tk.Label(row, text=num, bg=COLORS["bg_card"], fg=color,
                 font=(FONTS["family"], 26, "bold"),
                 width=3).pack(side="left")
        txt = tk.Frame(row, bg=COLORS["bg_card"])
        txt.pack(side="left", padx=(PADDING["sm"], 0))
        tk.Label(txt, text=title, bg=COLORS["bg_card"],
                 fg=COLORS["text_primary"],
                 font=FONTS["h2"], anchor="w").pack(anchor="w")
        tk.Label(txt, text=desc, bg=COLORS["bg_card"],
                 fg=COLORS["text_secondary"],
                 font=FONTS["sm"], anchor="w").pack(anchor="w")
        tk.Checkbutton(row, variable=var, text="启用",
                       bg=COLORS["bg_card"], fg=COLORS["text_primary"],
                       activebackground=COLORS["bg_card"],
                       selectcolor=COLORS["bg_active"],
                       font=FONTS["md"], cursor="hand2"
                       ).pack(side="right", padx=PADDING["lg"])

    def _prog_row(self, parent, label, style):
        frame = tk.Frame(parent, bg=COLORS["bg_card"])
        frame.pack(fill="x", pady=(0, PADDING["sm"]))
        lbl = tk.Label(frame, text=label, bg=COLORS["bg_card"],
                        fg=COLORS["text_secondary"],
                        font=FONTS["sm"], anchor="w")
        lbl.pack(fill="x", pady=(0, 3))
        bar_row = tk.Frame(frame, bg=COLORS["bg_card"])
        bar_row.pack(fill="x")
        var = tk.DoubleVar(value=0)
        bar = ttk.Progressbar(bar_row, variable=var, maximum=100,
                               style=style, mode="determinate")
        bar.pack(side="left", fill="x", expand=True)
        pct = tk.Label(bar_row, text="0%", bg=COLORS["bg_card"],
                        fg=COLORS["text_secondary"],
                        font=FONTS["sm"], width=6)
        pct.pack(side="right", padx=(8, 0))

        class _P:
            def __init__(s, v, b, p, l):
                s._var, s._bar, s._pct, s._lbl = v, b, p, l
            def update(s, value, label=None):
                def _do():
                    s._var.set(value)
                    s._pct.config(text=f"{int(value)}%")
                    if label: s._lbl.config(text=label)
                try: s._bar.after(0, _do)
                except Exception: pass
            def reset(s, label=""):
                s.update(0, label)
        return _P(var, bar, pct, lbl)

    # ── 文件选择 ──────────────────────────────────────────────

    def _pick_files(self):
        paths = filedialog.askopenfilenames(
            title=f"选择文档（最多 {_MAX_FILES} 个）",
            filetypes=[("支持的文档", "*.pdf *.docx *.doc"),
                       ("所有文件", "*.*")])
        if paths:
            self._apply_files(list(paths))

    def _pick_folder(self):
        folder = filedialog.askdirectory(title="选择文件夹（自动扫描支持格式）")
        if not folder:
            return
        files = []
        for ext in ("*.pdf", "*.docx", "*.doc"):
            files.extend(Path(folder).glob(ext))
        files = sorted(set(files))
        if not files:
            messagebox.showinfo("提示",
                "该文件夹中没有找到支持的文档（PDF / DOCX / DOC）", parent=self)
            return
        self._apply_files([str(f) for f in files])

    def _apply_files(self, paths: list):
        if len(paths) > _MAX_FILES:
            messagebox.showwarning("文件数量超限",
                f"最多支持 {_MAX_FILES} 个文件，已自动截取前 {_MAX_FILES} 个。\n"
                f"（共选择了 {len(paths)} 个文件）", parent=self)
            paths = paths[:_MAX_FILES]

        valid, oversized = [], []
        for p in paths:
            sz = os.path.getsize(p)
            if sz > _MAX_SIZE_BYTES:
                oversized.append(f"{Path(p).name}  ({_fmt_size(sz)})")
            else:
                valid.append(p)

        if oversized:
            messagebox.showwarning("文件过大",
                f"以下文件超过 {_MAX_SIZE_MB} MB 限制，已跳过：\n\n"
                + "\n".join(oversized), parent=self)

        if not valid:
            return

        self._current_files = valid
        self._refresh_file_list()

    def _refresh_file_list(self):
        for w in self._file_list_frame.winfo_children():
            w.destroy()
        n = len(self._current_files)
        self._file_count_lbl.config(
            text=f"已选择 {n} / {_MAX_FILES} 个文件",
            fg=COLORS["text_primary"])
        for p in self._current_files:
            name = Path(p).name
            sz   = _fmt_size(os.path.getsize(p))
            row  = tk.Frame(self._file_list_frame, bg=COLORS["bg_active"])
            row.pack(fill="x", pady=1, padx=1)
            tk.Label(row, text="📄", bg=COLORS["bg_active"],
                     font=FONTS["xs"]).pack(side="left", padx=(4, 0))
            display = name if len(name) <= 28 else name[:25] + "…"
            tk.Label(row, text=display, bg=COLORS["bg_active"],
                     fg=COLORS["text_primary"], font=FONTS["xs"],
                     anchor="w").pack(side="left", fill="x", expand=True)
            tk.Label(row, text=sz, bg=COLORS["bg_active"],
                     fg=COLORS["text_muted"],
                     font=FONTS["xs"]).pack(side="right", padx=(0, 6))
        self._status_bar.set(f"已选择 {n} 个文件", "normal")

    # ── 执行控制 ──────────────────────────────────────────────

    def _run(self):
        if not self._current_files:
            messagebox.showwarning("提示", "请先选择输入文件", parent=self)
            return
        steps = [self._step1_var.get(),
                 self._step2_var.get(),
                 self._step3_var.get()]
        if not any(steps):
            messagebox.showwarning("提示", "请至少选择一个执行步骤", parent=self)
            return
        self._is_running = True
        self._run_btn.config(state="disabled")
        self._stop_btn.config(state="normal")
        self._log.clear()
        for p in [self._p1, self._p2, self._p3]:
            p.reset()
        self._cur_file_lbl.config(text="")
        self._status_bar.set("流水线执行中…", "running")
        threading.Thread(target=self._run_all_files,
                         args=(list(self._current_files), steps),
                         daemon=True).start()

    def _stop(self):
        self._is_running = False
        self._log.append("用户请求停止", "WARNING")
        self._stop_btn.config(state="disabled")
        self._status_bar.set("已停止", "warning")

    # ── 批量驱动（逐文件顺序执行）───────────────────────────

    def _run_all_files(self, file_paths: list, steps: list):
        n = len(file_paths)
        success_count = 0
        for i, fp in enumerate(file_paths):
            if not self._is_running:
                break
            fname = Path(fp).name
            self.after(0, lambda f=fname, idx=i:
                       self._cur_file_lbl.config(
                           text=f"处理中 [{idx+1}/{n}]：{f}"))
            self._log.append("─" * 40, "DIM")
            self._log.append(f"[{i+1}/{n}]  {fname}", "GOLD")
            for p in [self._p1, self._p2, self._p3]:
                p.reset()
            ok = self._worker(fp, steps)
            if ok:
                success_count += 1

        if self._is_running:
            self._log.append("═" * 40, "DIM")
            self._log.append(
                f"批量流水线完成：{success_count}/{n} 个文件处理成功", "GOLD")
            self._status_bar.set(
                f"流水线完成 ✓  ({success_count}/{n})", "success")
            self.after(0, lambda: messagebox.showinfo(
                "完成",
                f"所有步骤执行完毕，共处理 {n} 个文件（{success_count} 成功）。\n"
                "输出文件保存在各输入文件的同级目录中。",
                parent=self))

        self._is_running = False
        self.after(0, self._reset_ui)

    # ── 流水线核心（正确顺序：解析→标注→转Word）────────────

    def _worker(self, file_path, steps) -> bool:
        """处理单个文件，返回 True 表示整体成功（无致命异常）"""
        run_parse, run_annotate, run_word = steps
        md_path    = None
        images_dir = None
        _ok        = True

        try:
            # ══ Step 01：文档解析 PDF/Word → Markdown ════════
            if run_parse and self._is_running:
                self._log.append("Step 01  文档解析 开始", "GOLD")
                self._p1.update(5, "Step 01 — 解析中…")
                try:
                    from config.settings import get_config
                    cfg = get_config()
                    vlm = cfg.get("vlm", cfg.get("online", {}))
                    if not vlm.get("api_key"):
                        self._log.append("未配置 VLM API Key，跳过步骤01", "WARNING")
                        self._p1.update(0, "Step 01 — 跳过（未配置VLM）")
                    else:
                        from core.pipeline import run_pipeline
                        result = run_pipeline(
                            file_path,
                            progress_cb=lambda c, t, m: (
                                self._p1.update(5 + (c/max(t,1))*90,
                                                f"Step 01 — {m}"),
                                self._log.append(m, "INFO")
                            ),
                        )
                        if result.success:
                            md_path    = result.md_path
                            images_dir = result.images_dir
                            self._shared["last_md_path"]    = md_path
                            self._shared["last_images_dir"] = images_dir
                            self._p1.update(100, "Step 01 — ✓ 完成")
                            self._log.append(f"Step 01 完成：{md_path}", "SUCCESS")
                        else:
                            self._p1.update(0, "Step 01 — ✗ 失败")
                            self._log.append(
                                f"Step 01 失败：{result.error[:200]}", "ERROR")
                except Exception as e:
                    self._p1.update(0, "Step 01 — ✗ 失败")
                    self._log.append(f"Step 01 异常：{e}", "ERROR")
            else:
                self._p1.update(0, "Step 01 — 已跳过")
                md_path    = self._shared.get("last_md_path", "")
                images_dir = self._shared.get("last_images_dir", "")

            # ══ Step 02：数据标注 Markdown → 带tags的Markdown ═
            annotated_md = md_path   # 默认用原始md，若标注成功则更新
            if run_annotate and self._is_running:
                if not md_path or not Path(md_path).exists():
                    self._p2.update(0, "Step 02 — 跳过（无MD文件）")
                    self._log.append("Step 02 跳过：无可用 MD 文件", "WARNING")
                else:
                    self._log.append("Step 02  数据标注 开始", "GOLD")
                    self._p2.update(5, "Step 02 — 准备中…")
                    try:
                        annotated_md = self._run_annotate(md_path)
                        if annotated_md:
                            self._p2.update(100, "Step 02 — ✓ 完成")
                            self._log.append(
                                f"Step 02 完成：{annotated_md}", "SUCCESS")
                        else:
                            self._p2.update(0, "Step 02 — 跳过（未配置LLM）")
                            annotated_md = md_path
                    except Exception as e:
                        self._p2.update(0, "Step 02 — ✗ 失败")
                        self._log.append(f"Step 02 失败：{e}", "ERROR")
                        annotated_md = md_path   # 失败时回退用原始md
            else:
                self._p2.update(0, "Step 02 — 已跳过")

            # ══ Step 03：MD → Word（用标注后的MD）═══════════
            if run_word and self._is_running:
                src_md = annotated_md or md_path
                if not src_md or not Path(src_md).exists():
                    self._p3.update(0, "Step 03 — 跳过（无MD文件）")
                    self._log.append("Step 03 跳过：无可用 MD 文件", "WARNING")
                else:
                    self._log.append("Step 03  MD → Word 开始", "GOLD")
                    self._p3.update(20, "Step 03 — 转换中…")
                    try:
                        from tools.tool2_converter import run_conversion
                        out_path = str(Path(src_md).with_suffix(".docx"))
                        if Path(out_path).exists():
                            out_path = str(Path(src_md).parent /
                                          (Path(src_md).stem + "_output.docx"))
                        imgs = images_dir or str(Path(src_md).parent / "images")
                        run_conversion(
                            src_md, imgs, out_path,
                            log_cb=lambda m: self._log.append(m.strip(), "INFO"))
                        self._p3.update(100, "Step 03 — ✓ 完成")
                        self._log.append(f"Step 03 完成：{out_path}", "SUCCESS")
                    except Exception as e:
                        self._p3.update(0, "Step 03 — ✗ 失败")
                        self._log.append(f"Step 03 失败：{e}", "ERROR")
            else:
                self._p3.update(0, "Step 03 — 已跳过")

        except Exception as e:
            self._log.append(f"流水线异常：{traceback.format_exc()[:300]}", "ERROR")
            _ok = False

        return _ok

    def _run_annotate(self, md_path: str) -> str | None:
        """
        流水线调用的标注方法。
        输出 _annotated.md，格式：正文 + @@@tags@@@ + #### 分隔符
        未配置 API Key 返回 None。
        """
        import json as _j, re as _r
        import requests as _rq
        from tools.tool3_annotator import (
            _parse_units, _rebuild, PROMPT_TEXT, PROMPT_TABLE_BATCH,
            _table_row_to_text)
        from config.settings import get_config

        cfg = get_config()
        llm = cfg.get("llm", {})
        api_key    = llm.get("api_key", "")
        base_url   = llm.get("base_url", "https://api.siliconflow.cn/v1")
        model_name = llm.get("model", "Pro/deepseek-ai/DeepSeek-V3")
        model_url  = base_url.rstrip("/") + "/chat/completions"

        if not api_key:
            self._log.append("未配置 LLM API Key，跳过步骤02", "WARNING")
            return None

        def _call(sys_p, usr_m):
            h = {"Content-Type": "application/json",
                 "Authorization": f"Bearer {api_key}"}
            p = {"model": model_name,
                 "messages": [{"role": "system", "content": sys_p},
                               {"role": "user",   "content": usr_m}],
                 "max_tokens": 2048, "temperature": 0}
            r = _rq.post(model_url, headers=h, json=p, timeout=300)
            r.raise_for_status()
            raw = r.json()["choices"][0]["message"]["content"].strip()
            raw = _r.sub(r'^```(?:json)?\s*', '', raw)
            raw = _r.sub(r'\s*```$', '', raw).strip()
            return raw

        with open(md_path, "r", encoding="utf-8") as f:
            content = f.read()
        orig_lines = content.split("\n")
        units  = _parse_units(content, 30)
        total  = len(units)
        results = [None] * total

        for idx, unit in enumerate(units):
            if not self._is_running:
                break
            self._p2.update(5 + (idx / max(total, 1)) * 90,
                            f"Step 02 — {idx+1}/{total}")
            try:
                if unit["type"] == "text":
                    sys_p = (PROMPT_TEXT
                             .replace("__HEADING__",   unit["heading"])
                             .replace("__PARAGRAPH__", unit["para"]))
                    raw = _call(sys_p, '只输出JSON对象{"tags":[]}')
                    try:
                        tags = _j.loads(raw).get("tags", [])
                    except Exception:
                        tags = ["解析失败"]
                    results[idx] = {"type": "text", "tags": tags[:6]}
                else:
                    fmt = unit.get("fmt", "")
                    headers   = unit.get("headers", [])
                    data_rows = unit.get("data_rows", [])

                    # png_ref 类型：从旁边的 .json 文件加载 grid 数据
                    if fmt == "png_ref" and not data_rows:
                        import json as _j2
                        from tools.tool3_annotator import _is_sub_header as _ish, _merge_header_rows as _mhr
                        tbl_ref = unit.get("tbl_ref", "")
                        if tbl_ref:
                            json_path = str(Path(md_path).parent / tbl_ref.replace(".png", ".json"))
                            if _r.search(r'\.png$', tbl_ref) and Path(json_path).exists():
                                try:
                                    grid = _j2.loads(Path(json_path).read_text(encoding="utf-8"))
                                    if grid and len(grid) > 1:
                                        raw_headers = grid[0]
                                        if len(grid) > 2 and _ish(grid[1]):
                                            headers   = grid[1]
                                            data_rows = grid[2:]
                                        elif len(grid) > 2 and not _ish(grid[1]) \
                                                and not any(
                                                    _r.fullmatch(r'[\d.,]+', v)
                                                    for v in grid[1] if v.strip()):
                                            headers   = _mhr(raw_headers, grid[1])
                                            data_rows = grid[2:]
                                        else:
                                            headers   = raw_headers
                                            data_rows = grid[1:]
                                except Exception:
                                    pass

                    if not data_rows:
                        results[idx] = {"type": "table", "row_results": []}
                        continue
                    # 整表一次批量请求
                    table_text = "\n".join(unit["lines"]) if fmt != "png_ref" else \
                                 "\n".join(["\t".join(r) for r in [headers] + list(data_rows)])
                    row_count  = len(data_rows)
                    sys_p = (PROMPT_TABLE_BATCH
                             .replace("__TITLE__", unit["title"])
                             .replace("__TABLE__", table_text))
                    try:
                        raw = _call(sys_p,
                                    f'只输出JSON数组，共{row_count}个对象，每行数据对应一个，完整保留所有字段。')
                        batch = _j.loads(raw)
                        if not isinstance(batch, list):
                            raise ValueError("not a list")
                    except Exception:
                        batch = [{"question": "", "answer": _table_row_to_text(headers, row),
                                   "tags": [unit["title"]]} for row in data_rows]
                    if len(batch) < row_count:
                        for row in data_rows[len(batch):]:
                            batch.append({"question": "", "answer": _table_row_to_text(headers, row),
                                           "tags": [unit["title"]]})
                    row_results = []
                    for item in batch:
                        if not isinstance(item, dict):
                            continue
                        raw_tags = item.get("tags", [unit["title"]])
                        clean_tags = [t for t in raw_tags
                                      if isinstance(t, str) and 0 < len(t) <= 25][:5]
                        row_results.append({
                            "question": item.get("question", "").strip(),
                            "answer":   (item.get("answer", "") or
                                         item.get("description", "")).strip(),
                            "tags":     clean_tags,
                        })
                    results[idx] = {"type": "table", "row_results": row_results}
                    self._log.append(
                        f"  表格 {idx+1}：{len(row_results)} 行已处理", "SUCCESS")
            except Exception as e:
                self._log.append(f"  单元{idx+1}失败: {e}", "ERROR")
                results[idx] = None

        out_lines = _rebuild(orig_lines, units, results)
        out_path  = str(Path(md_path).parent / (Path(md_path).stem + "_annotated.md"))
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("\n".join(out_lines))

        chunk_count = out_lines.count("####")
        self._log.append(f"标注完成：{chunk_count} 个 chunk → {out_path}", "GOLD")
        return out_path

    def _reset_ui(self):
        self._run_btn.config(state="normal")
        self._stop_btn.config(state="disabled")