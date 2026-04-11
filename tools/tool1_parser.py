"""
Tool 1 — 文档解析面板
PDF / Word  →  Markdown（支持批量多文件 + 并行处理）
"""
import os
import threading
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from gui.theme import COLORS, FONTS, PADDING
from gui.widgets import (
    StyledButton, LogView, ProgressRow,
    SectionHeader, Divider,
)

MAX_FILES = 10
MAX_FILE_SIZE_MB = 50
MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024


def _size_str(size_bytes: int) -> str:
    if size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.0f} KB"
    return f"{size_bytes / 1024 / 1024:.1f} MB"


class _FileProgressRow(tk.Frame):
    """右侧面板中单个文件的进度行"""

    _ICONS = {"waiting": "⏳", "running": "⚙", "done": "✓", "error": "✗"}
    _COLORS = {
        "waiting": COLORS["text_muted"],
        "running": COLORS["tool1_color"],
        "done":    COLORS["success"],
        "error":   COLORS["error"],
    }
    _PB_STYLES = {
        "waiting": "Blue.Horizontal.TProgressbar",
        "running": "Blue.Horizontal.TProgressbar",
        "done":    "Green.Horizontal.TProgressbar",
        "error":   "Orange.Horizontal.TProgressbar",
    }

    def __init__(self, parent, filename: str, size_str: str, row_bg: str, **kw):
        super().__init__(parent, bg=row_bg, pady=3, **kw)
        self._row_bg = row_bg

        # 状态图标
        self._icon_lbl = tk.Label(self, text="⏳", bg=row_bg,
                                   fg=COLORS["text_muted"],
                                   font=FONTS["sm"], width=2)
        self._icon_lbl.pack(side="left", padx=(6, 2))

        # 文件名 + 消息
        info = tk.Frame(self, bg=row_bg)
        info.pack(side="left", fill="x", expand=True, padx=4)

        display = filename if len(filename) <= 30 else filename[:27] + "…"
        tk.Label(info, text=display, bg=row_bg, fg=COLORS["text_primary"],
                 font=FONTS["xs"], anchor="w").pack(fill="x")
        self._msg_lbl = tk.Label(info,
                                   text=f"等待中  ·  {size_str}",
                                   bg=row_bg, fg=COLORS["text_muted"],
                                   font=FONTS["xs"], anchor="w")
        self._msg_lbl.pack(fill="x")

        # 进度条（右侧固定宽度）
        pb_wrap = tk.Frame(self, bg=row_bg, width=170)
        pb_wrap.pack(side="right", padx=(0, 10))
        pb_wrap.pack_propagate(False)
        self._pb_var = tk.DoubleVar(value=0)
        self._pb = ttk.Progressbar(pb_wrap, variable=self._pb_var,
                                    length=160, mode="determinate",
                                    style="Blue.Horizontal.TProgressbar")
        self._pb.pack(fill="x", pady=6)

    def set_status(self, status: str, pct: float, msg: str):
        icon  = self._ICONS.get(status, "⚙")
        color = self._COLORS.get(status, COLORS["text_muted"])
        style = self._PB_STYLES.get(status, "Blue.Horizontal.TProgressbar")
        self._icon_lbl.config(text=icon, fg=color)
        self._pb_var.set(pct)
        self._pb.config(style=style)
        self._msg_lbl.config(text=msg[:70], fg=color)


class Tool1Panel(tk.Frame):
    """文档解析：PDF/Word → Markdown（支持批量 + 并行）"""

    def __init__(self, parent, shared_state: dict, status_bar,
                 navigate_cb=None, **kw):
        bg = kw.pop("bg", COLORS["bg_main"])
        super().__init__(parent, bg=bg, **kw)
        self._shared      = shared_state
        self._status_bar  = status_bar
        self._navigate    = navigate_cb
        self._current_files: list = []          # 已选文件路径列表
        self._results:       list = []          # ParseResult 列表（与文件一一对应）
        self._cancel_event        = None
        self._file_row_frames:list = []         # 右侧每文件进度行
        self._done_count          = 0
        self._build()

    # ── UI 构建 ───────────────────────────────────────────────

    def _build(self):
        hdr = tk.Frame(self, bg=COLORS["bg_card"])
        hdr.pack(fill="x")
        if self._navigate:
            tk.Button(hdr, text="◈  主页",
                      bg=COLORS["bg_card"], fg=COLORS["accent"],
                      activebackground=COLORS["bg_hover"],
                      activeforeground=COLORS["accent"],
                      relief="flat", bd=0, cursor="hand2",
                      font=FONTS["sm"], padx=PADDING["md"],
                      command=lambda: self._navigate("home")
                      ).pack(side="right", pady=PADDING["sm"],
                             padx=PADDING["md"])
        tk.Label(hdr, text="⬢  文档解析",
                 bg=COLORS["bg_card"], fg=COLORS["tool1_color"],
                 font=FONTS["h1"]).pack(side="left",
                                         padx=PADDING["xl"], pady=PADDING["lg"])
        tk.Label(hdr, text="PDF / Word  →  Markdown  ·  支持批量并行处理（最多10个）",
                 bg=COLORS["bg_card"], fg=COLORS["text_muted"],
                 font=FONTS["sm"]).pack(side="left", pady=PADDING["lg"])
        Divider(self).pack(fill="x")

        body = tk.Frame(self, bg=COLORS["bg_main"])
        body.pack(fill="both", expand=True)
        self._build_left(body)
        self._build_right(body)

    def _build_left(self, parent):
        left = tk.Frame(parent, bg=COLORS["bg_sidebar"], width=300)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        Divider(left, color=COLORS["border_light"]).pack(fill="x")

        # ── 文件选择区 ──
        fcard = tk.Frame(left, bg=COLORS["bg_sidebar"])
        fcard.pack(fill="x", padx=PADDING["md"], pady=PADDING["md"])

        SectionHeader(fcard, "选择文件",
                      color=COLORS["tool1_color"],
                      bg=COLORS["bg_sidebar"]).pack(fill="x", pady=(0, PADDING["sm"]))

        btn_row = tk.Frame(fcard, bg=COLORS["bg_sidebar"])
        btn_row.pack(fill="x", pady=(0, PADDING["sm"]))
        StyledButton(btn_row, text="📄 选择文件",
                     style="secondary",
                     command=self._pick_files).pack(side="left", fill="x",
                                                     expand=True, padx=(0, 4))
        StyledButton(btn_row, text="📁 选择文件夹",
                     style="secondary",
                     command=self._pick_folder).pack(side="left", fill="x",
                                                      expand=True)

        self._file_count_lbl = tk.Label(
            fcard,
            text=f"未选择文件（最多 {MAX_FILES} 个，每个 ≤ {MAX_FILE_SIZE_MB} MB）",
            bg=COLORS["bg_sidebar"], fg=COLORS["text_muted"],
            font=FONTS["xs"], anchor="w", wraplength=260,
        )
        self._file_count_lbl.pack(fill="x")

        # 文件列表（固定高度区域）
        list_wrap = tk.Frame(fcard, bg=COLORS["bg_active"], pady=2)
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

        Divider(left, color=COLORS["border_light"]).pack(fill="x", pady=PADDING["sm"])

        # ── 操作按钮 ──
        ctrl = tk.Frame(left, bg=COLORS["bg_sidebar"])
        ctrl.pack(fill="x", padx=PADDING["md"])

        self._start_btn = StyledButton(ctrl, text="▶  开始批量解析",
                                       style="blue",
                                       command=self._start_parse)
        self._start_btn.pack(fill="x", pady=(0, PADDING["sm"]))

        self._cancel_btn = StyledButton(ctrl, text="⏹  取消全部",
                                        style="secondary",
                                        command=self._cancel_parse)
        self._cancel_btn.pack(fill="x")
        self._cancel_btn.config(state="disabled")

        # ── 总体进度 ──
        prog_frame = tk.Frame(left, bg=COLORS["bg_sidebar"])
        prog_frame.pack(fill="x", padx=PADDING["md"], pady=PADDING["md"])
        self._progress = ProgressRow(prog_frame,
                                     bar_style="Blue.Horizontal.TProgressbar",
                                     bg=COLORS["bg_sidebar"])
        self._progress.pack(fill="x")

        self._task_counter = tk.Label(
            left, text="",
            bg=COLORS["bg_sidebar"], fg=COLORS["text_muted"],
            font=FONTS["xs"],
        )
        self._task_counter.pack(padx=PADDING["md"], anchor="w")

        Divider(left, color=COLORS["border_light"]).pack(fill="x")

        # ── 统计信息 ──
        stats_frame = tk.LabelFrame(
            left, text=" 解析统计 ",
            bg=COLORS["bg_sidebar"], fg=COLORS["text_muted"],
            font=FONTS["xs"], bd=1, relief="groove",
        )
        stats_frame.pack(fill="x", padx=PADDING["md"], pady=PADDING["md"])

        self._stats = {}
        for key, label in [("files",    "成功文件"),
                            ("blocks",   "内容块"),
                            ("figures",  "图片"),
                            ("tables",   "表格"),
                            ("formulas", "公式")]:
            row = tk.Frame(stats_frame, bg=COLORS["bg_sidebar"])
            row.pack(fill="x", padx=PADDING["sm"], pady=1)
            tk.Label(row, text=label, bg=COLORS["bg_sidebar"],
                     fg=COLORS["text_secondary"],
                     font=FONTS["xs"]).pack(side="left")
            v = tk.Label(row, text="—", bg=COLORS["bg_sidebar"],
                         fg=COLORS["tool1_color"], font=FONTS["xs"])
            v.pack(side="right")
            self._stats[key] = v

        # ── 导出 & 传递按钮（底部）──
        dl_frame = tk.Frame(left, bg=COLORS["bg_sidebar"])
        dl_frame.pack(fill="x", padx=PADDING["md"], side="bottom",
                      pady=PADDING["md"])

        self._dl_zip_btn = StyledButton(dl_frame, text="📦  导出全部 ZIP",
                                        style="blue",
                                        command=self._save_all_zips)
        self._dl_zip_btn.pack(fill="x", pady=(0, PADDING["sm"]))
        self._dl_zip_btn.config(state="disabled")

        Divider(left, color=COLORS["border_light"]).pack(fill="x",
                                                          pady=(PADDING["sm"], 0))
        self._pass_btn = StyledButton(
            left, text="→  传递给 MD→Word / 数据标注",
            style="ghost",
            command=self._pass_to_next,
        )
        self._pass_btn.pack(fill="x", padx=PADDING["md"], pady=PADDING["sm"])
        self._pass_btn.config(state="disabled")

    def _build_right(self, parent):
        right = tk.Frame(parent, bg=COLORS["bg_main"])
        right.pack(side="left", fill="both", expand=True)

        # 文件进度表头
        prog_hdr = tk.Frame(right, bg=COLORS["bg_card"], height=32)
        prog_hdr.pack(fill="x")
        prog_hdr.pack_propagate(False)
        tk.Label(prog_hdr, text="📊  文件处理进度",
                 bg=COLORS["bg_card"], fg=COLORS["text_secondary"],
                 font=FONTS["sm"]).pack(side="left", padx=PADDING["md"], pady=6)
        Divider(right).pack(fill="x")

        # 可滚动的文件进度行容器
        rows_outer = tk.Frame(right, bg=COLORS["bg_main"])
        rows_outer.pack(fill="x", padx=PADDING["md"], pady=(PADDING["sm"], 0))

        canvas = tk.Canvas(rows_outer, bg=COLORS["bg_main"],
                            highlightthickness=0, height=240)
        vsb = ttk.Scrollbar(rows_outer, orient="vertical",
                             command=canvas.yview,
                             style="Dark.Vertical.TScrollbar")
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        self._rows_canvas = canvas
        self._rows_inner  = tk.Frame(canvas, bg=COLORS["bg_main"])
        self._rows_window = canvas.create_window(
            (0, 0), window=self._rows_inner, anchor="nw")

        def _on_frame_configure(e):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(e):
            canvas.itemconfig(self._rows_window, width=e.width)

        self._rows_inner.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        # 滚轮支持
        import sys as _sys
        def _scroll(event):
            if _sys.platform == "win32":
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            else:
                canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")
        canvas.bind("<MouseWheel>", _scroll)

        # 占位文字
        self._rows_placeholder = tk.Label(
            self._rows_inner,
            text="\n  选择文件后点击「开始批量解析」，此处将实时显示每个文件的处理进度\n",
            bg=COLORS["bg_main"], fg=COLORS["text_muted"],
            font=FONTS["sm"],
        )
        self._rows_placeholder.pack(fill="x")

        Divider(right).pack(fill="x", pady=(PADDING["sm"], 0))

        # 日志区（下半部分，可伸缩）
        log_frame = tk.Frame(right, bg=COLORS["bg_main"])
        log_frame.pack(fill="both", expand=True,
                       padx=PADDING["md"], pady=PADDING["md"])
        tk.Label(log_frame, text="运行日志",
                 bg=COLORS["bg_main"], fg=COLORS["text_muted"],
                 font=FONTS["xs"]).pack(anchor="w")
        self._log = LogView(log_frame, height=10)
        self._log.pack(fill="both", expand=True, pady=(2, 0))

    # ── 文件选择 ──────────────────────────────────────────────

    def _pick_files(self):
        paths = filedialog.askopenfilenames(
            title=f"选择文档（最多 {MAX_FILES} 个）",
            filetypes=[
                ("支持的文档", "*.pdf *.docx *.doc"),
                ("PDF 文件",   "*.pdf"),
                ("Word 文档",  "*.docx *.doc"),
                ("所有文件",   "*.*"),
            ],
        )
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
            messagebox.showinfo(
                "提示",
                "该文件夹中没有找到支持的文档（PDF / DOCX / DOC）",
                parent=self,
            )
            return
        self._apply_files([str(f) for f in files])

    def _apply_files(self, paths: list):
        """校验并设置文件列表"""
        if len(paths) > MAX_FILES:
            messagebox.showwarning(
                "文件数量超限",
                f"最多支持 {MAX_FILES} 个文件，已自动截取前 {MAX_FILES} 个。\n"
                f"（共选择了 {len(paths)} 个文件）",
                parent=self,
            )
            paths = paths[:MAX_FILES]

        valid, oversized = [], []
        for p in paths:
            size = os.path.getsize(p)
            if size > MAX_FILE_SIZE_BYTES:
                oversized.append(f"{Path(p).name}  ({_size_str(size)})")
            else:
                valid.append(p)

        if oversized:
            messagebox.showwarning(
                "文件过大",
                f"以下文件超过 {MAX_FILE_SIZE_MB} MB 限制，已跳过：\n\n"
                + "\n".join(oversized),
                parent=self,
            )

        if not valid:
            return

        self._current_files = valid
        self._refresh_file_list_ui()
        self._rebuild_progress_rows()

    def _refresh_file_list_ui(self):
        """刷新左侧文件列表区域"""
        for w in self._file_list_frame.winfo_children():
            w.destroy()

        n = len(self._current_files)
        self._file_count_lbl.config(
            text=f"已选择 {n} / {MAX_FILES} 个文件",
            fg=COLORS["text_primary"],
        )

        for p in self._current_files:
            name = Path(p).name
            sz   = _size_str(os.path.getsize(p))
            row  = tk.Frame(self._file_list_frame, bg=COLORS["bg_active"])
            row.pack(fill="x", pady=1, padx=1)
            tk.Label(row, text="📄", bg=COLORS["bg_active"],
                     font=FONTS["xs"]).pack(side="left", padx=(4, 0))
            display = name if len(name) <= 22 else name[:19] + "…"
            tk.Label(row, text=display, bg=COLORS["bg_active"],
                     fg=COLORS["text_primary"], font=FONTS["xs"],
                     anchor="w").pack(side="left", fill="x", expand=True)
            tk.Label(row, text=sz, bg=COLORS["bg_active"],
                     fg=COLORS["text_muted"],
                     font=FONTS["xs"]).pack(side="right", padx=(0, 6))

        self._status_bar.set(f"已选择 {n} 个文件", "normal")

    def _rebuild_progress_rows(self):
        """在右侧重建每文件进度行（文件选择变化时调用）"""
        for w in self._rows_inner.winfo_children():
            w.destroy()
        self._file_row_frames = []

        if not self._current_files:
            self._rows_placeholder = tk.Label(
                self._rows_inner,
                text="\n  选择文件后点击「开始批量解析」，此处将实时显示进度\n",
                bg=COLORS["bg_main"], fg=COLORS["text_muted"],
                font=FONTS["sm"],
            )
            self._rows_placeholder.pack(fill="x")
            return

        for i, p in enumerate(self._current_files):
            name = Path(p).name
            sz   = _size_str(os.path.getsize(p))
            row_bg = COLORS["bg_card"] if i % 2 == 0 else COLORS["bg_main"]
            row = _FileProgressRow(self._rows_inner, name, sz, row_bg)
            row.pack(fill="x", pady=1)
            self._file_row_frames.append(row)

    # ── 解析控制 ──────────────────────────────────────────────

    def _start_parse(self):
        if not self._current_files:
            messagebox.showwarning("提示", "请先选择文件", parent=self)
            return

        try:
            from config.settings import get_config
            cfg  = get_config()
            mode = cfg.get("model_mode", "online")
            if mode == "online" and not cfg.get("online", {}).get("api_key"):
                messagebox.showwarning(
                    "提示", "请先在「模型设置」中填写 API Key", parent=self)
                return
        except Exception:
            pass

        n = len(self._current_files)
        self._results    = [None] * n
        self._done_count = 0

        # 重建进度行（保证状态干净）
        self._rebuild_progress_rows()

        # UI 状态
        self._start_btn.config(state="disabled")
        self._cancel_btn.config(state="normal")
        self._dl_zip_btn.config(state="disabled")
        self._pass_btn.config(state="disabled")
        self._progress.reset(f"正在处理 0 / {n}…")
        self._task_counter.config(text=f"已完成 0 / {n}")
        self._status_bar.set(f"批量解析中（共 {n} 个文件）…", "running")
        for k in self._stats:
            self._stats[k].config(text="—")

        self._log.append(
            f"开始批量解析，共 {n} 个文件（最多 3 个并行）", "INFO")

        from core.pipeline import run_batch_async
        self._cancel_event = run_batch_async(
            self._current_files,
            file_progress_cb=self._on_file_progress,
            file_done_cb=self._on_file_done,
            all_done_cb=self._on_all_done,
            max_workers=3,
        )

    def _cancel_parse(self):
        if self._cancel_event:
            self._cancel_event.set()
        self._cancel_btn.config(state="disabled")
        self._start_btn.config(state="normal")
        self._status_bar.set("已取消", "normal")
        self._progress.reset("已取消")
        self._log.append("用户取消了批量解析", "WARNING")

    # ── 进度 / 完成回调（worker 线程中调用，需 after 转到主线程）──

    def _on_file_progress(self, idx, cur, total, msg):
        self.after(0, lambda: self._update_file_row(idx, cur, total, msg))

    def _update_file_row(self, idx, cur, total, msg):
        pct = (cur / max(total, 1)) * 100
        if 0 <= idx < len(self._file_row_frames):
            self._file_row_frames[idx].set_status("running", pct, msg)
        fname = Path(self._current_files[idx]).name
        self._log.append(f"[{fname}] {msg}", "INFO")
        self._status_bar.set(f"[{fname}] {msg[:60]}", "running")

    def _on_file_done(self, idx, result):
        self.after(0, lambda: self._handle_file_done(idx, result))

    def _handle_file_done(self, idx, result):
        self._results[idx] = result
        self._done_count   += 1
        n     = len(self._current_files)
        fname = Path(self._current_files[idx]).name

        if 0 <= idx < len(self._file_row_frames):
            row = self._file_row_frames[idx]
            if result.success:
                blocks = result.stats.get("blocks", 0)
                row.set_status("done", 100, f"完成 ✓  {blocks} 个内容块")
            else:
                row.set_status("error", 0,
                               f"失败：{result.error[:50]}")

        overall_pct = (self._done_count / n) * 100
        self._progress.update(overall_pct,
                               f"已完成 {self._done_count} / {n}")
        self._task_counter.config(text=f"已完成 {self._done_count} / {n}")

        if result.success:
            self._log.append(
                f"[{fname}] 解析完成 ✓  输出: {Path(result.md_path).parent}",
                "SUCCESS")
        else:
            self._log.append(
                f"[{fname}] 失败：{result.error[:120]}", "ERROR")

    def _on_all_done(self, results):
        self.after(0, lambda: self._handle_all_done(results))

    def _handle_all_done(self, results):
        self._start_btn.config(state="normal")
        self._cancel_btn.config(state="disabled")

        successful = [r for r in results if r and r.success]
        failed     = [r for r in results if r and not r.success]

        # 聚合统计
        self._stats["files"].config(
            text=f"{len(successful)} / {len(results)}")
        self._stats["blocks"].config(
            text=str(sum(r.stats.get("blocks", 0) for r in successful)))
        self._stats["figures"].config(
            text=str(sum(r.stats.get("figures", 0) for r in successful)))
        self._stats["tables"].config(
            text=str(sum(r.stats.get("tables", 0) for r in successful)))
        self._stats["formulas"].config(
            text=str(sum(r.stats.get("formulas", 0) for r in successful)))

        if successful:
            self._dl_zip_btn.config(state="normal")
            self._pass_btn.config(state="normal")
            # 将最后一个成功结果传入共享状态
            last = successful[-1]
            self._shared["last_md_path"]   = last.md_path
            self._shared["last_images_dir"] = last.images_dir

        self._progress.update(100,
            f"全部完成 ✓  （{len(successful)} / {len(results)} 成功）")

        if failed:
            self._status_bar.set(
                f"完成：{len(successful)} 成功，{len(failed)} 失败", "error")
            self._log.append(
                f"批量解析完成：{len(successful)} 成功，{len(failed)} 失败",
                "WARNING")
        else:
            self._status_bar.set(
                f"全部解析完成 ✓  共 {len(successful)} 个文件", "success")
            self._log.append(
                f"批量解析完成：{len(successful)} 个文件全部成功 ✓",
                "SUCCESS")

    # ── 导出 ──────────────────────────────────────────────────

    def _save_all_zips(self):
        successful = [
            (i, r) for i, r in enumerate(self._results)
            if r and r.success and r.zip_path
        ]
        if not successful:
            messagebox.showinfo("提示", "没有可导出的成功结果", parent=self)
            return

        if len(successful) == 1:
            i, r = successful[0]
            dest = filedialog.asksaveasfilename(
                title="导出 ZIP",
                defaultextension=".zip",
                initialfile=Path(self._current_files[i]).stem + ".zip",
                filetypes=[("ZIP 压缩包", "*.zip"), ("所有文件", "*.*")],
            )
            if dest:
                import shutil
                shutil.copy2(r.zip_path, dest)
                self._status_bar.set(f"已导出：{dest}", "success")
        else:
            folder = filedialog.askdirectory(title="选择导出文件夹（所有 ZIP 将保存到此处）")
            if folder:
                import shutil
                for i, r in successful:
                    shutil.copy2(r.zip_path,
                                 Path(folder) / Path(r.zip_path).name)
                self._status_bar.set(
                    f"已导出 {len(successful)} 个 ZIP 到：{folder}", "success")
                messagebox.showinfo(
                    "导出完成",
                    f"已将 {len(successful)} 个 ZIP 文件导出到：\n{folder}",
                    parent=self,
                )

    def _pass_to_next(self):
        last = next(
            (r for r in reversed(self._results) if r and r.success), None)
        if last:
            self._shared["last_md_path"]   = last.md_path
            self._shared["last_images_dir"] = last.images_dir
            messagebox.showinfo(
                "已传递",
                f"已将最新 Markdown 路径传递给其他工具：\n{last.md_path}\n\n"
                "请切换到「MD→Word」或「数据标注」标签直接使用。",
                parent=self,
            )

    def open_settings(self):
        try:
            from gui.settings_window import SettingsWindow
            SettingsWindow(self.winfo_toplevel())
        except Exception as e:
            messagebox.showerror("错误", str(e), parent=self)
