#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""md2docx Tkinter GUI — 마크다운을 Word(.docx) / 한글(.hwpx)로 변환한다."""

from __future__ import annotations

import importlib.util
import os
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, font as tkfont, messagebox, ttk

import md2docx


def _load_hwpx_cli():
    """로컬 md2hwpx.py를 패키지와 다른 이름으로 불러온다.

    파일명과 패키지명이 같아서 일반 import 하면 패키지 대신 스크립트가
    캐시되어 한글 변환이 실패한다.
    """
    path = Path(__file__).resolve().parent / "md2hwpx.py"
    spec = importlib.util.spec_from_file_location("md2hwpx_cli", path)
    if spec is None or spec.loader is None:
        raise ImportError(f"md2hwpx.py를 불러올 수 없습니다 → {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules["md2hwpx_cli"] = module
    spec.loader.exec_module(module)
    return module


md2hwpx = _load_hwpx_cli()

MD_FILETYPES = [("마크다운", "*.md *.markdown"), ("모든 파일", "*.*")]
DOCX_FILETYPES = [("Word 문서", "*.docx"), ("모든 파일", "*.*")]
HWPX_FILETYPES = [("한글 문서", "*.hwpx"), ("모든 파일", "*.*")]


def _enable_dpi_awareness() -> None:
    if sys.platform != "win32":
        return
    try:
        from ctypes import windll

        windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def _open_in_explorer(path: Path) -> None:
    target = path if path.is_dir() else path.parent
    if sys.platform == "win32":
        os.startfile(target)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.run(["open", str(target)], check=False)
    else:
        subprocess.run(["xdg-open", str(target)], check=False)


def _scan_markdown(folder: Path, recursive: bool) -> list[Path]:
    patterns = ("*.md", "*.markdown")
    found: list[Path] = []
    walker = folder.rglob if recursive else folder.glob
    for pattern in patterns:
        found.extend(p for p in walker(pattern) if p.is_file())
    return sorted({p.resolve() for p in found})


class ConverterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("md2docx — 마크다운 변환기")
        self.root.minsize(760, 640)
        self.root.geometry("840x760")

        self._busy = False
        self._paths: dict[str, Path] = {}
        self._file_seq = 0

        self.do_docx = tk.BooleanVar(value=True)
        self.do_hwpx = tk.BooleanVar(value=False)
        self.save_beside_source = tk.BooleanVar(value=True)
        self.out_dir = tk.StringVar(value="")
        self.docx_template = tk.StringVar(value="")
        self.hwpx_template = tk.StringVar(value="")
        self.toc = tk.BooleanVar(value=False)
        self.toc_depth = tk.IntVar(value=3)
        self.highlight = tk.StringVar(value=md2docx.DEFAULT_HIGHLIGHT_STYLE)
        self.include_subfolders = tk.BooleanVar(value=False)
        self.open_after = tk.BooleanVar(value=True)

        self._build_style()
        self._build_menu()
        self._build_ui()
        self._sync_option_states()
        self._log("마크다운 파일을 추가한 뒤 변환을 누르세요.", "info")

    def _build_style(self) -> None:
        style = ttk.Style(self.root)
        if sys.platform == "win32" and "vista" in style.theme_names():
            style.theme_use("vista")
        default = tkfont.nametofont("TkDefaultFont")
        default.configure(size=10)
        self.root.option_add("*Font", default)
        style.configure("TLabelframe.Label", font=(default.actual("family"), 10, "bold"))
        style.configure("Convert.TButton", font=(default.actual("family"), 11, "bold"), padding=(12, 8))

    def _build_menu(self) -> None:
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="파일 추가…", command=self.add_files, accelerator="Ctrl+O")
        file_menu.add_command(label="폴더 추가…", command=self.add_folder)
        file_menu.add_separator()
        file_menu.add_command(label="변환", command=self.start_convert, accelerator="Ctrl+Return")
        file_menu.add_separator()
        file_menu.add_command(label="종료", command=self.root.destroy)
        menubar.add_cascade(label="파일", menu=file_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="정보", command=self._show_about)
        menubar.add_cascade(label="도움말", menu=help_menu)
        self.root.config(menu=menubar)

        self.root.bind("<Control-o>", lambda _e: self.add_files())
        self.root.bind("<Control-Return>", lambda _e: self.start_convert())
        self.root.bind("<Delete>", lambda _e: self.remove_selected())

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=14)
        main.pack(fill=tk.BOTH, expand=True)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(0, weight=1)
        main.rowconfigure(5, weight=1)

        self._build_files_section(main)
        self._build_output_section(main)
        self._build_template_section(main)
        self._build_options_section(main)

        convert_row = ttk.Frame(main)
        convert_row.grid(row=4, column=0, sticky="ew", pady=(10, 8))
        convert_row.columnconfigure(0, weight=1)
        self.convert_btn = ttk.Button(
            convert_row,
            text="변환 시작",
            style="Convert.TButton",
            command=self.start_convert,
        )
        self.convert_btn.grid(row=0, column=0, sticky="ew")

        self._build_log_section(main)

        status = ttk.Frame(self.root, padding=(14, 4, 14, 8))
        status.pack(fill=tk.X, side=tk.BOTTOM)
        self.status_var = tk.StringVar(value="준비")
        ttk.Label(status, textvariable=self.status_var, foreground="#555").pack(anchor="w")

    def _build_files_section(self, parent: ttk.Frame) -> None:
        frame = ttk.Labelframe(parent, text="변환할 파일", padding=8)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        tree_wrap = ttk.Frame(frame)
        tree_wrap.grid(row=0, column=0, sticky="nsew")
        tree_wrap.columnconfigure(0, weight=1)
        tree_wrap.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            tree_wrap,
            columns=("name", "folder"),
            show="headings",
            selectmode="extended",
            height=6,
        )
        self.tree.heading("name", text="파일")
        self.tree.heading("folder", text="위치")
        self.tree.column("name", width=220, stretch=False)
        self.tree.column("folder", width=480, stretch=True)
        vsb = ttk.Scrollbar(tree_wrap, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        btns = ttk.Frame(frame)
        btns.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(btns, text="파일 추가…", command=self.add_files).pack(side=tk.LEFT)
        ttk.Button(btns, text="폴더 추가…", command=self.add_folder).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(btns, text="선택 제거", command=self.remove_selected).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(btns, text="목록 비우기", command=self.clear_files).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Checkbutton(
            btns,
            text="하위 폴더 포함",
            variable=self.include_subfolders,
        ).pack(side=tk.RIGHT)

    def _build_output_section(self, parent: ttk.Frame) -> None:
        frame = ttk.Labelframe(parent, text="출력", padding=8)
        frame.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        frame.columnconfigure(1, weight=1)

        fmt = ttk.Frame(frame)
        fmt.grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(fmt, text="형식:").pack(side=tk.LEFT)
        ttk.Checkbutton(
            fmt, text="Word (.docx)", variable=self.do_docx, command=self._sync_option_states
        ).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Checkbutton(
            fmt, text="한글 (.hwpx)", variable=self.do_hwpx, command=self._sync_option_states
        ).pack(side=tk.LEFT, padx=(8, 0))

        ttk.Checkbutton(
            frame,
            text="원본과 같은 폴더에 저장",
            variable=self.save_beside_source,
            command=self._sync_option_states,
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 4))

        ttk.Label(frame, text="출력 폴더").grid(row=2, column=0, sticky="w")
        self.out_dir_entry = ttk.Entry(frame, textvariable=self.out_dir)
        self.out_dir_entry.grid(row=2, column=1, sticky="ew", padx=(8, 6))
        self.out_dir_btn = ttk.Button(frame, text="찾아보기…", command=self.browse_out_dir)
        self.out_dir_btn.grid(row=2, column=2, sticky="e")

    def _build_template_section(self, parent: ttk.Frame) -> None:
        frame = ttk.Labelframe(parent, text="템플릿 (참조 문서)", padding=8)
        frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Word (.docx)").grid(row=0, column=0, sticky="w")
        self.docx_tpl_entry = ttk.Entry(frame, textvariable=self.docx_template)
        self.docx_tpl_entry.grid(row=0, column=1, sticky="ew", padx=(8, 6), pady=2)
        docx_btns = ttk.Frame(frame)
        docx_btns.grid(row=0, column=2, sticky="e")
        self.docx_tpl_browse = ttk.Button(docx_btns, text="불러오기…", command=self.browse_docx_template)
        self.docx_tpl_browse.pack(side=tk.LEFT)
        self.docx_tpl_make = ttk.Button(docx_btns, text="기본 생성", command=self.make_docx_template)
        self.docx_tpl_make.pack(side=tk.LEFT, padx=(4, 0))
        self.docx_tpl_clear = ttk.Button(docx_btns, text="지우기", command=lambda: self.docx_template.set(""))
        self.docx_tpl_clear.pack(side=tk.LEFT, padx=(4, 0))

        ttk.Label(frame, text="한글 (.hwpx)").grid(row=1, column=0, sticky="w")
        self.hwpx_tpl_entry = ttk.Entry(frame, textvariable=self.hwpx_template)
        self.hwpx_tpl_entry.grid(row=1, column=1, sticky="ew", padx=(8, 6), pady=2)
        hwpx_btns = ttk.Frame(frame)
        hwpx_btns.grid(row=1, column=2, sticky="e")
        self.hwpx_tpl_browse = ttk.Button(hwpx_btns, text="불러오기…", command=self.browse_hwpx_template)
        self.hwpx_tpl_browse.pack(side=tk.LEFT)
        self.hwpx_tpl_make = ttk.Button(hwpx_btns, text="기본 생성", command=self.make_hwpx_template)
        self.hwpx_tpl_make.pack(side=tk.LEFT, padx=(4, 0))
        self.hwpx_tpl_clear = ttk.Button(hwpx_btns, text="지우기", command=lambda: self.hwpx_template.set(""))
        self.hwpx_tpl_clear.pack(side=tk.LEFT, padx=(4, 0))

        ttk.Label(
            frame,
            text="비워 두면 기본 스타일을 사용합니다. 워드/한글에서 스타일을 편집한 뒤 다시 불러오세요.",
            foreground="#555",
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 0))

    def _build_options_section(self, parent: ttk.Frame) -> None:
        frame = ttk.Labelframe(parent, text="옵션", padding=8)
        frame.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        frame.columnconfigure(1, weight=1)

        word_opts = ttk.Frame(frame)
        word_opts.grid(row=0, column=0, columnspan=3, sticky="w")
        self.toc_chk = ttk.Checkbutton(word_opts, text="목차 생성 (Word)", variable=self.toc)
        self.toc_chk.pack(side=tk.LEFT)
        ttk.Label(word_opts, text="목차 깊이").pack(side=tk.LEFT, padx=(16, 4))
        self.toc_depth_spin = ttk.Spinbox(
            word_opts, from_=1, to=6, width=4, textvariable=self.toc_depth, state="readonly"
        )
        self.toc_depth_spin.pack(side=tk.LEFT)
        ttk.Label(word_opts, text="코드 하이라이트").pack(side=tk.LEFT, padx=(16, 4))
        self.highlight_combo = ttk.Combobox(
            word_opts,
            textvariable=self.highlight,
            values=md2docx.VALID_HIGHLIGHT_STYLES,
            state="readonly",
            width=14,
        )
        self.highlight_combo.pack(side=tk.LEFT)

        ttk.Label(frame, text="추가 리소스 폴더").grid(row=1, column=0, sticky="nw", pady=(10, 0))
        res_wrap = ttk.Frame(frame)
        res_wrap.grid(row=1, column=1, sticky="nsew", padx=(8, 6), pady=(10, 0))
        res_wrap.columnconfigure(0, weight=1)
        self.res_list = tk.Listbox(res_wrap, height=3, exportselection=False)
        self.res_list.grid(row=0, column=0, sticky="ew")
        res_btns = ttk.Frame(frame)
        res_btns.grid(row=1, column=2, sticky="ne", pady=(10, 0))
        ttk.Button(res_btns, text="추가…", command=self.add_resource_dir).pack(fill=tk.X)
        ttk.Button(res_btns, text="제거", command=self.remove_resource_dir).pack(fill=tk.X, pady=(4, 0))

        ttk.Checkbutton(
            frame,
            text="변환 후 결과 폴더 열기",
            variable=self.open_after,
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(10, 0))

    def _build_log_section(self, parent: ttk.Frame) -> None:
        frame = ttk.Labelframe(parent, text="진행 상황", padding=8)
        frame.grid(row=5, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.log = tk.Text(frame, height=8, wrap=tk.WORD, state=tk.DISABLED, relief=tk.FLAT)
        vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.log.yview)
        self.log.configure(yscrollcommand=vsb.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        self.log.tag_configure("ok", foreground="#0a7a2f")
        self.log.tag_configure("err", foreground="#b00020")
        self.log.tag_configure("info", foreground="#333333")

    def _sync_option_states(self) -> None:
        docx_state = tk.NORMAL if self.do_docx.get() else tk.DISABLED
        hwpx_state = tk.NORMAL if self.do_hwpx.get() else tk.DISABLED
        for widget in (
            self.docx_tpl_entry,
            self.docx_tpl_browse,
            self.docx_tpl_make,
            self.docx_tpl_clear,
            self.toc_chk,
            self.highlight_combo,
        ):
            widget.configure(state=docx_state)
        self.toc_depth_spin.configure(state="readonly" if self.do_docx.get() else tk.DISABLED)
        for widget in (
            self.hwpx_tpl_entry,
            self.hwpx_tpl_browse,
            self.hwpx_tpl_make,
            self.hwpx_tpl_clear,
        ):
            widget.configure(state=hwpx_state)

        save_custom = not self.save_beside_source.get()
        self.out_dir_entry.configure(state=tk.NORMAL if save_custom else tk.DISABLED)
        self.out_dir_btn.configure(state=tk.NORMAL if save_custom else tk.DISABLED)

    def _show_about(self) -> None:
        messagebox.showinfo(
            "md2docx",
            "마크다운을 Word(.docx)와 한글(.hwpx)로 변환합니다.\n\n"
            "Word: pandoc    한글: md2hwpx\n"
            "템플릿(참조 문서)으로 폰트·헤딩 스타일을 맞출 수 있습니다.",
            parent=self.root,
        )

    def _log(self, message: str, tag: str = "info") -> None:
        self.log.configure(state=tk.NORMAL)
        self.log.insert(tk.END, message.rstrip() + "\n", tag)
        self.log.see(tk.END)
        self.log.configure(state=tk.DISABLED)

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        self.convert_btn.configure(state=tk.DISABLED if busy else tk.NORMAL)
        self.status_var.set("변환 중…" if busy else "준비")

    def add_files(self) -> None:
        picked = filedialog.askopenfilenames(
            parent=self.root,
            title="변환할 마크다운 파일 선택",
            filetypes=MD_FILETYPES,
        )
        if picked:
            self._add_paths([Path(p) for p in picked])

    def add_folder(self) -> None:
        folder = filedialog.askdirectory(parent=self.root, title="마크다운이 있는 폴더 선택")
        if not folder:
            return
        files = _scan_markdown(Path(folder), recursive=self.include_subfolders.get())
        if not files:
            messagebox.showinfo("폴더 추가", "이 폴더에서 마크다운 파일을 찾지 못했습니다.", parent=self.root)
            return
        self._add_paths(files)

    def _add_paths(self, paths: list[Path]) -> None:
        added = 0
        existing = {p.resolve() for p in self._paths.values()}
        for raw in paths:
            path = raw.expanduser().resolve()
            if path in existing:
                continue
            self._file_seq += 1
            iid = f"f{self._file_seq}"
            self._paths[iid] = path
            existing.add(path)
            self.tree.insert("", tk.END, iid=iid, values=(path.name, str(path.parent)))
            added += 1
        if added:
            self.status_var.set(f"파일 {len(self._paths)}개")

    def remove_selected(self) -> None:
        for iid in self.tree.selection():
            self.tree.delete(iid)
            self._paths.pop(iid, None)
        self.status_var.set(f"파일 {len(self._paths)}개" if self._paths else "준비")

    def clear_files(self) -> None:
        self.tree.delete(*self.tree.get_children())
        self._paths.clear()
        self.status_var.set("준비")

    def browse_out_dir(self) -> None:
        folder = filedialog.askdirectory(parent=self.root, title="출력 폴더 선택")
        if folder:
            self.out_dir.set(folder)

    def browse_docx_template(self) -> None:
        picked = filedialog.askopenfilename(
            parent=self.root,
            title="Word 템플릿(참조 문서) 선택",
            filetypes=DOCX_FILETYPES,
        )
        if picked:
            self.docx_template.set(picked)

    def browse_hwpx_template(self) -> None:
        picked = filedialog.askopenfilename(
            parent=self.root,
            title="한글 템플릿(참조 문서) 선택",
            filetypes=HWPX_FILETYPES,
        )
        if picked:
            self.hwpx_template.set(picked)

    def make_docx_template(self) -> None:
        dest = filedialog.asksaveasfilename(
            parent=self.root,
            title="Word 기본 템플릿 저장",
            defaultextension=".docx",
            filetypes=DOCX_FILETYPES,
            initialfile="reference.docx",
        )
        if not dest:
            return
        try:
            pandoc = md2docx.find_pandoc()
            created = md2docx.make_reference_doc(pandoc, Path(dest))
        except md2docx.ConversionError as e:
            messagebox.showerror("템플릿 생성 실패", str(e), parent=self.root)
            return
        self.docx_template.set(str(created))
        self._log(f"Word 템플릿 생성: {created}", "ok")
        messagebox.showinfo(
            "템플릿 생성",
            "기본 Word 템플릿을 만들었습니다.\n워드에서 스타일(본문, Heading, Source Code)을 편집한 뒤 변환에 사용하세요.",
            parent=self.root,
        )

    def make_hwpx_template(self) -> None:
        dest = filedialog.asksaveasfilename(
            parent=self.root,
            title="한글 기본 템플릿 저장",
            defaultextension=".hwpx",
            filetypes=HWPX_FILETYPES,
            initialfile="reference.hwpx",
        )
        if not dest:
            return
        try:
            created = md2hwpx.make_reference_doc(Path(dest))
        except md2hwpx.ConversionError as e:
            messagebox.showerror("템플릿 생성 실패", str(e), parent=self.root)
            return
        self.hwpx_template.set(str(created))
        self._log(f"한글 템플릿 생성: {created}", "ok")
        messagebox.showinfo(
            "템플릿 생성",
            "기본 한글 템플릿을 만들었습니다.\n한글에서 제목·본문·표 스타일을 편집한 뒤 변환에 사용하세요.",
            parent=self.root,
        )

    def add_resource_dir(self) -> None:
        folder = filedialog.askdirectory(parent=self.root, title="이미지 등 추가 탐색 폴더")
        if not folder:
            return
        existing = set(self.res_list.get(0, tk.END))
        if folder not in existing:
            self.res_list.insert(tk.END, folder)

    def remove_resource_dir(self) -> None:
        selection = list(self.res_list.curselection())
        for index in reversed(selection):
            self.res_list.delete(index)

    def _validate(self) -> bool:
        if self._busy:
            return False
        if not self._paths:
            messagebox.showwarning("파일 없음", "변환할 마크다운 파일을 추가하세요.", parent=self.root)
            return False
        if not self.do_docx.get() and not self.do_hwpx.get():
            messagebox.showwarning("형식 없음", "Word 또는 한글 중 하나 이상을 선택하세요.", parent=self.root)
            return False
        if not self.save_beside_source.get() and not self.out_dir.get().strip():
            messagebox.showwarning("출력 폴더", "출력 폴더를 지정하거나 ‘원본과 같은 폴더에 저장’을 선택하세요.", parent=self.root)
            return False
        docx_tpl = self.docx_template.get().strip()
        if self.do_docx.get() and docx_tpl and not Path(docx_tpl).is_file():
            messagebox.showerror("템플릿", f"Word 템플릿을 찾을 수 없습니다.\n{docx_tpl}", parent=self.root)
            return False
        hwpx_tpl = self.hwpx_template.get().strip()
        if self.do_hwpx.get() and hwpx_tpl and not Path(hwpx_tpl).is_file():
            messagebox.showerror("템플릿", f"한글 템플릿을 찾을 수 없습니다.\n{hwpx_tpl}", parent=self.root)
            return False
        return True

    def start_convert(self) -> None:
        if not self._validate():
            return

        files = [self._paths[iid] for iid in self.tree.get_children()]
        out_dir = None if self.save_beside_source.get() else Path(self.out_dir.get().strip())
        resources = list(self.res_list.get(0, tk.END))
        job = {
            "files": files,
            "do_docx": self.do_docx.get(),
            "do_hwpx": self.do_hwpx.get(),
            "out_dir": out_dir,
            "docx_ref": Path(self.docx_template.get().strip()) if self.docx_template.get().strip() else None,
            "hwpx_ref": Path(self.hwpx_template.get().strip()) if self.hwpx_template.get().strip() else None,
            "toc": self.toc.get(),
            "toc_depth": int(self.toc_depth.get()),
            "highlight": self.highlight.get(),
            "resources": resources,
            "open_after": self.open_after.get(),
        }

        self._set_busy(True)
        self._log("—— 변환 시작 ——", "info")
        threading.Thread(target=self._convert_worker, args=(job,), daemon=True).start()

    def _convert_worker(self, job: dict) -> None:
        outputs: list[Path] = []
        errors: list[str] = []
        files: list[Path] = job["files"]
        total = len(files)
        pandoc = None

        try:
            if job["do_docx"]:
                try:
                    pandoc = md2docx.find_pandoc()
                except md2docx.ConversionError as e:
                    self._ui(lambda msg=str(e): self._finish([], [msg], job))
                    return

            for index, src in enumerate(files, start=1):
                self._ui(
                    lambda i=index, n=total, name=src.name: self.status_var.set(
                        f"변환 중… {i}/{n}  {name}"
                    )
                )

                if job["do_docx"] and pandoc:
                    out = (job["out_dir"] / f"{src.stem}.docx") if job["out_dir"] else src.with_suffix(".docx")
                    try:
                        result = md2docx.convert(
                            pandoc=pandoc,
                            src=src,
                            out=out,
                            highlight_style=job["highlight"],
                            reference_doc=job["docx_ref"],
                            toc=job["toc"],
                            extra_resource_paths=job["resources"],
                            toc_depth=job["toc_depth"],
                        )
                        outputs.append(result)
                        self._ui(lambda r=result: self._log(f"완료 (docx): {r}", "ok"))
                    except Exception as e:
                        errors.append(f"{src.name} (docx): {e}")
                        self._ui(lambda msg=str(e): self._log(msg, "err"))

                if job["do_hwpx"]:
                    out = (job["out_dir"] / f"{src.stem}.hwpx") if job["out_dir"] else src.with_suffix(".hwpx")
                    try:
                        result = md2hwpx.convert(
                            src=src,
                            out=out,
                            reference_doc=job["hwpx_ref"],
                            extra_resource_paths=job["resources"],
                        )
                        outputs.append(result)
                        self._ui(lambda r=result: self._log(f"완료 (hwpx): {r}", "ok"))
                    except Exception as e:
                        errors.append(f"{src.name} (hwpx): {e}")
                        self._ui(lambda msg=str(e): self._log(msg, "err"))
        except Exception as e:
            errors.append(str(e))
            self._ui(lambda msg=str(e): self._log(f"변환 중단: {msg}", "err"))

        self._ui(lambda: self._finish(outputs, errors, job))

    def _ui(self, callback) -> None:
        self.root.after(0, callback)

    def _finish(self, outputs: list[Path], errors: list[str], job: dict) -> None:
        self._set_busy(False)
        self.status_var.set(f"완료 — 성공 {len(outputs)}건, 실패 {len(errors)}건")
        self._log(f"—— 끝: 성공 {len(outputs)} / 실패 {len(errors)} ——", "info")

        if errors and not outputs:
            messagebox.showerror("변환 실패", "\n\n".join(errors[:4]), parent=self.root)
            return
        if errors:
            messagebox.showwarning(
                "일부 실패",
                f"{len(outputs)}건 성공, {len(errors)}건 실패했습니다.\n자세한 내용은 진행 상황을 확인하세요.",
                parent=self.root,
            )
        else:
            messagebox.showinfo("변환 완료", f"{len(outputs)}개 파일을 변환했습니다.", parent=self.root)

        if job["open_after"] and outputs:
            folder = job["out_dir"] if job["out_dir"] else outputs[0].parent
            try:
                _open_in_explorer(Path(folder))
            except Exception:
                pass


def run_app() -> None:
    _enable_dpi_awareness()
    root = tk.Tk()
    ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    run_app()
