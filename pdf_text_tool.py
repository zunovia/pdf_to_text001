"""
PDF テキスト変換ツール - PDFからテキスト/Markdownを抽出

2エンジン構成:
  - pymupdf4llm: デジタルPDFの高速テキスト抽出（GPU不要）
  - marker-pdf:   スキャン/画像PDFの高品質OCR抽出（PyTorch使用）

使い方:
  CLI:  python pdf_text_tool.py input.pdf
  GUI:  python pdf_text_tool.py --gui
"""

import argparse
import os
import re
import sys
import threading
from pathlib import Path


# ---------------------------------------------------------------------------
# PDF種別判定
# ---------------------------------------------------------------------------

def detect_pdf_type(input_path):
    """
    PDFがデジタル（テキスト埋め込み）かスキャン（画像ベース）かを判定。

    Returns:
        "digital" or "scanned"
    """
    try:
        import pymupdf
    except ImportError:
        return "unknown"

    doc = pymupdf.open(str(input_path))
    total_pages = len(doc)
    if total_pages == 0:
        doc.close()
        return "unknown"

    pages_with_text = 0
    sample_pages = min(total_pages, 10)  # 最大10ページサンプリング

    for i in range(sample_pages):
        page = doc[i]
        text = page.get_text().strip()
        # 1ページあたり50文字以上あればテキスト有りと判定
        if len(text) > 50:
            pages_with_text += 1

    doc.close()

    # 半分以上のページにテキストがあればデジタル
    if pages_with_text >= sample_pages * 0.5:
        return "digital"
    return "scanned"


# ---------------------------------------------------------------------------
# テキスト抽出エンジン
# ---------------------------------------------------------------------------

def extract_text_pymupdf(input_path, output_format="txt", progress_callback=None):
    """
    pymupdf4llm を使用してデジタルPDFからテキスト抽出。

    Args:
        input_path: 入力PDFパス
        output_format: "txt" or "md"
        progress_callback: fn(message: str)
    Returns:
        抽出されたテキスト (str)
    """
    def log(msg):
        if progress_callback:
            progress_callback(msg)

    try:
        import pymupdf4llm
    except ImportError:
        raise RuntimeError(
            "pymupdf4llm がインストールされていません。\n"
            "  → pip install pymupdf4llm を実行してください。"
        )

    log("pymupdf4llm でテキスト抽出中...")

    md_text = pymupdf4llm.to_markdown(str(input_path))

    if output_format == "txt":
        # Markdownの装飾を除去してプレーンテキストに変換
        text = _markdown_to_plain_text(md_text)
        log("テキスト変換完了")
        return text

    log("Markdown変換完了")
    return md_text


def extract_text_marker(input_path, output_format="txt", progress_callback=None):
    """
    marker-pdf を使用してスキャン/画像PDFから高品質OCRテキスト抽出。

    Args:
        input_path: 入力PDFパス
        output_format: "txt" or "md"
        progress_callback: fn(message: str)
    Returns:
        抽出されたテキスト (str)
    """
    def log(msg):
        if progress_callback:
            progress_callback(msg)

    try:
        from marker.converters.pdf import PdfConverter
        from marker.models import create_model_dict
        from marker.output import text_from_rendered
    except ImportError:
        raise RuntimeError(
            "marker-pdf がインストールされていません。\n"
            "  → pip install marker-pdf を実行してください。\n"
            "  （PyTorchも必要です）"
        )

    log("marker-pdf でOCRテキスト抽出中（時間がかかる場合があります）...")

    converter = PdfConverter(artifact_dict=create_model_dict())
    rendered = converter(str(input_path))
    text, _, images = text_from_rendered(rendered)

    if output_format == "txt":
        text = _markdown_to_plain_text(text)
        log("テキスト変換完了")
        return text

    log("Markdown変換完了")
    return text


def _markdown_to_plain_text(md_text):
    """Markdownテキストからプレーンテキストに変換"""
    text = md_text
    # 見出しの # を除去
    text = re.sub(r'^#{1,6}\s+', '', text, flags=re.MULTILINE)
    # 太字・斜体の ** / * / __ / _ を除去
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'\*(.+?)\*', r'\1', text)
    text = re.sub(r'__(.+?)__', r'\1', text)
    text = re.sub(r'_(.+?)_', r'\1', text)
    # インラインコード ` を除去
    text = re.sub(r'`(.+?)`', r'\1', text)
    # リンク [text](url) → text
    text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', text)
    # 画像 ![alt](url) を除去
    text = re.sub(r'!\[([^\]]*)\]\([^\)]+\)', r'\1', text)
    # 水平線を除去
    text = re.sub(r'^[-*_]{3,}\s*$', '', text, flags=re.MULTILINE)
    # リストマーカー - / * / + を除去（行頭）
    text = re.sub(r'^\s*[-*+]\s+', '', text, flags=re.MULTILINE)
    # 番号付きリストマーカーを除去
    text = re.sub(r'^\s*\d+\.\s+', '', text, flags=re.MULTILINE)
    # 連続空行を1つに
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


# ---------------------------------------------------------------------------
# 統合抽出関数
# ---------------------------------------------------------------------------

def extract_text(input_path, engine="auto", output_format="txt",
                 progress_callback=None):
    """
    PDFからテキストを抽出する統合関数。

    Args:
        input_path: 入力PDFパス
        engine: "auto" / "pymupdf" / "marker"
        output_format: "txt" / "md"
        progress_callback: fn(message: str)
    Returns:
        抽出されたテキスト (str)
    """
    input_path = Path(input_path)
    if not input_path.exists():
        raise FileNotFoundError(f"ファイルが見つかりません: {input_path}")

    def log(msg):
        if progress_callback:
            progress_callback(msg)

    log(f"処理開始: {input_path.name}")

    # エンジン選択
    if engine == "auto":
        log("PDF種別を判定中...")
        pdf_type = detect_pdf_type(input_path)
        log(f"  判定結果: {pdf_type}")

        if pdf_type == "scanned":
            # marker-pdf が利用可能かチェック
            try:
                import marker  # noqa: F401
                engine = "marker"
                log("  → marker-pdf (高精度OCR) を使用します")
            except ImportError:
                log("  → marker-pdf 未インストール。pymupdf4llm で試行します")
                engine = "pymupdf"
        else:
            engine = "pymupdf"
            log("  → pymupdf4llm (高速) を使用します")

    # テキスト抽出
    if engine == "marker":
        return extract_text_marker(input_path, output_format, progress_callback)
    else:
        return extract_text_pymupdf(input_path, output_format, progress_callback)


def extract_and_save(input_path, output_path=None, engine="auto",
                     output_format="txt", progress_callback=None):
    """
    PDFからテキストを抽出してファイルに保存。

    Args:
        input_path: 入力PDFパス
        output_path: 出力ファイルパス (None なら自動生成)
        engine: "auto" / "pymupdf" / "marker"
        output_format: "txt" / "md"
        progress_callback: fn(message: str)
    Returns:
        output_path (str)
    """
    input_path = Path(input_path)

    if output_path is None:
        ext = ".md" if output_format == "md" else ".txt"
        output_path = input_path.with_suffix(ext)
    output_path = Path(output_path)

    text = extract_text(
        input_path, engine=engine, output_format=output_format,
        progress_callback=progress_callback
    )

    output_path.write_text(text, encoding="utf-8")

    def log(msg):
        if progress_callback:
            progress_callback(msg)
    log(f"保存完了: {output_path.name}")

    return str(output_path)


# ---------------------------------------------------------------------------
# 依存関係チェック
# ---------------------------------------------------------------------------

def check_dependencies():
    """利用可能なエンジンを確認"""
    results = {}

    # pymupdf4llm
    try:
        import pymupdf4llm  # noqa: F401
        results["pymupdf4llm"] = True
    except ImportError:
        results["pymupdf4llm"] = False

    # marker-pdf
    try:
        import marker  # noqa: F401
        results["marker"] = True
    except ImportError:
        results["marker"] = False

    return results


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def run_cli():
    parser = argparse.ArgumentParser(
        description="PDF テキスト変換ツール - PDFからテキスト/Markdownを抽出"
    )
    parser.add_argument(
        "input", nargs="?",
        help="入力PDFファイルパス"
    )
    parser.add_argument(
        "-o", "--output",
        help="出力ファイルパス (省略時: <入力名>.txt or .md)"
    )
    parser.add_argument(
        "-e", "--engine", default="auto",
        choices=["auto", "pymupdf", "marker"],
        help="抽出エンジン (デフォルト: auto)"
    )
    parser.add_argument(
        "-f", "--format", default="txt",
        choices=["txt", "md"],
        help="出力形式 (デフォルト: txt)"
    )
    parser.add_argument(
        "--gui", action="store_true",
        help="GUIモードで起動"
    )
    parser.add_argument(
        "--check", action="store_true",
        help="依存関係を確認して終了"
    )

    args = parser.parse_args()

    # 依存関係チェック
    if args.check:
        deps = check_dependencies()
        print("=== 依存関係チェック ===")
        for name, available in deps.items():
            status = "OK" if available else "未インストール"
            print(f"  {name}: {status}")
        if not any(deps.values()):
            print("\nエラー: 少なくとも1つのエンジンが必要です。")
            print("  → pip install pymupdf4llm")
            sys.exit(1)
        else:
            print("\n少なくとも1つのエンジンが利用可能です。")
            sys.exit(0)

    # GUIモード
    if args.gui or args.input is None:
        run_gui()
        return

    # CLI 実行
    deps = check_dependencies()
    if not any(deps.values()):
        print("エラー: エンジンがインストールされていません。", file=sys.stderr)
        print("  → pip install pymupdf4llm", file=sys.stderr)
        sys.exit(1)

    try:
        output = extract_and_save(
            args.input,
            args.output,
            engine=args.engine,
            output_format=args.format,
            progress_callback=lambda msg: print(f"  {msg}")
        )
        print(f"\n出力: {output}")
    except Exception as e:
        print(f"\nエラー: {e}", file=sys.stderr)
        sys.exit(1)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    has_windnd = False
    try:
        import windnd
        has_windnd = True
    except ImportError:
        pass

    class App:
        def __init__(self, root):
            self.root = root
            self.root.title("PDF テキスト変換ツール")
            self.root.geometry("720x620")
            self.root.resizable(True, True)
            self.root.configure(bg="#f0f0f0")

            self.files = []
            self.processing = False

            self._build_ui()
            self._check_deps()

            if has_windnd:
                windnd.hook_dropfiles(self.root, func=self._on_drop)

        def _build_ui(self):
            # --- タイトル ---
            title_frame = tk.Frame(self.root, bg="#1a5276", pady=10)
            title_frame.pack(fill=tk.X)
            tk.Label(
                title_frame, text="PDF テキスト変換ツール",
                font=("Segoe UI", 16, "bold"), fg="white", bg="#1a5276"
            ).pack()
            tk.Label(
                title_frame,
                text="PDFからテキスト/Markdownを抽出",
                font=("Segoe UI", 9), fg="#aed6f1", bg="#1a5276"
            ).pack()

            # --- ドロップエリア ---
            drop_frame = tk.LabelFrame(
                self.root, text=" PDFファイル ",
                font=("Segoe UI", 10), padx=10, pady=5, bg="#f0f0f0"
            )
            drop_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(10, 5))

            self.drop_label = tk.Label(
                drop_frame,
                text=("ここにPDFをドラッグ&ドロップ\nまたは下のボタンでファイルを選択"
                      if has_windnd else
                      "下のボタンでファイルを選択してください"),
                font=("Segoe UI", 11),
                bg="#eaf2f8", fg="#7f8c8d",
                relief=tk.GROOVE, bd=2,
                width=50, height=3
            )
            self.drop_label.pack(fill=tk.BOTH, expand=True, pady=(5, 5))

            self.file_listbox = tk.Listbox(
                drop_frame, height=4,
                font=("Segoe UI", 9), selectmode=tk.EXTENDED
            )
            self.file_listbox.pack(fill=tk.BOTH, expand=True, pady=(0, 5))

            btn_frame = tk.Frame(drop_frame, bg="#f0f0f0")
            btn_frame.pack(fill=tk.X)
            ttk.Button(
                btn_frame, text="ファイルを追加...",
                command=self._add_files
            ).pack(side=tk.LEFT, padx=(0, 5))
            ttk.Button(
                btn_frame, text="選択を削除",
                command=self._remove_selected
            ).pack(side=tk.LEFT, padx=(0, 5))
            ttk.Button(
                btn_frame, text="すべてクリア",
                command=self._clear_files
            ).pack(side=tk.LEFT)

            # --- 設定 ---
            settings_frame = tk.LabelFrame(
                self.root, text=" 設定 ",
                font=("Segoe UI", 10), padx=10, pady=5, bg="#f0f0f0"
            )
            settings_frame.pack(fill=tk.X, padx=15, pady=5)

            # エンジン
            tk.Label(
                settings_frame, text="エンジン:",
                font=("Segoe UI", 10), bg="#f0f0f0"
            ).grid(row=0, column=0, sticky=tk.W, pady=2)

            self.engine_var = tk.StringVar(value="auto")
            engine_combo = ttk.Combobox(
                settings_frame, textvariable=self.engine_var,
                values=[
                    "auto (自動判定)",
                    "pymupdf (高速・デジタルPDF向け)",
                    "marker (高精度OCR・スキャンPDF向け)"
                ],
                width=35, state="readonly"
            )
            engine_combo.grid(row=0, column=1, sticky=tk.W, padx=10, pady=2)

            # 出力形式
            tk.Label(
                settings_frame, text="出力形式:",
                font=("Segoe UI", 10), bg="#f0f0f0"
            ).grid(row=1, column=0, sticky=tk.W, pady=2)

            self.format_var = tk.StringVar(value="txt")
            format_combo = ttk.Combobox(
                settings_frame, textvariable=self.format_var,
                values=["txt (プレーンテキスト)", "md (Markdown)"],
                width=35, state="readonly"
            )
            format_combo.grid(row=1, column=1, sticky=tk.W, padx=10, pady=2)

            # 出力先
            tk.Label(
                settings_frame, text="出力先:",
                font=("Segoe UI", 10), bg="#f0f0f0"
            ).grid(row=2, column=0, sticky=tk.W, pady=2)

            output_row = tk.Frame(settings_frame, bg="#f0f0f0")
            output_row.grid(row=2, column=1, sticky=tk.W, padx=10, pady=2)

            self.output_var = tk.StringVar(value="same")
            ttk.Radiobutton(
                output_row, text="入力と同じフォルダ",
                variable=self.output_var, value="same",
                command=self._on_output_mode_change
            ).pack(anchor=tk.W)

            custom_row = tk.Frame(output_row, bg="#f0f0f0")
            custom_row.pack(anchor=tk.W, fill=tk.X)
            ttk.Radiobutton(
                custom_row, text="フォルダを指定:",
                variable=self.output_var, value="custom",
                command=self._on_output_mode_change
            ).pack(side=tk.LEFT)

            self.output_dir_var = tk.StringVar(value="")
            self.output_dir_entry = ttk.Entry(
                custom_row, textvariable=self.output_dir_var, width=25,
                state=tk.DISABLED
            )
            self.output_dir_entry.pack(side=tk.LEFT, padx=(5, 3))

            self.output_dir_btn = ttk.Button(
                custom_row, text="参照...",
                command=self._browse_output_dir, state=tk.DISABLED
            )
            self.output_dir_btn.pack(side=tk.LEFT)

            # --- 実行ボタン ---
            exec_frame = tk.Frame(self.root, bg="#f0f0f0", pady=5)
            exec_frame.pack(fill=tk.X, padx=15)

            self.run_btn = ttk.Button(
                exec_frame, text="テキスト抽出 実行",
                command=self._start_extraction
            )
            self.run_btn.pack(fill=tk.X, ipady=5)

            # --- プログレス ---
            self.progress = ttk.Progressbar(
                self.root, mode="indeterminate"
            )
            self.progress.pack(fill=tk.X, padx=15, pady=(5, 0))

            # --- ログ ---
            log_frame = tk.LabelFrame(
                self.root, text=" ログ ",
                font=("Segoe UI", 10), padx=5, pady=5, bg="#f0f0f0"
            )
            log_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(5, 10))

            self.log_text = tk.Text(
                log_frame, height=6,
                font=("Consolas", 9), state=tk.DISABLED, wrap=tk.WORD
            )
            scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
            self.log_text.configure(yscrollcommand=scrollbar.set)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            self.log_text.pack(fill=tk.BOTH, expand=True)

            # --- ステータスバー ---
            self.status_var = tk.StringVar(value="準備完了")
            tk.Label(
                self.root, textvariable=self.status_var,
                font=("Segoe UI", 9), fg="#7f8c8d", bg="#f0f0f0",
                anchor=tk.W
            ).pack(fill=tk.X, padx=15, pady=(0, 5))

        def _check_deps(self):
            deps = check_dependencies()
            available = []
            missing = []
            for name, ok in deps.items():
                if ok:
                    available.append(name)
                else:
                    missing.append(name)

            if available:
                self._log(f"利用可能なエンジン: {', '.join(available)}")
            if missing:
                self._log(f"未インストール: {', '.join(missing)}")
            if not available:
                self._log("エラー: エンジンが1つもインストールされていません！")
                self._log("  → pip install pymupdf4llm")
                self.status_var.set("エンジン未インストール")

        def _log(self, msg):
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)
            self.log_text.configure(state=tk.DISABLED)

        def _on_output_mode_change(self):
            if self.output_var.get() == "custom":
                self.output_dir_entry.configure(state=tk.NORMAL)
                self.output_dir_btn.configure(state=tk.NORMAL)
                if not self.output_dir_var.get():
                    self._browse_output_dir()
            else:
                self.output_dir_entry.configure(state=tk.DISABLED)
                self.output_dir_btn.configure(state=tk.DISABLED)

        def _browse_output_dir(self):
            d = filedialog.askdirectory(title="出力先フォルダを選択")
            if d:
                self.output_dir_var.set(d)

        def _on_drop(self, files):
            for f in files:
                if isinstance(f, bytes):
                    for enc in ("utf-8", "cp932", "shift_jis"):
                        try:
                            f = f.decode(enc)
                            break
                        except (UnicodeDecodeError, LookupError):
                            continue
                    else:
                        f = f.decode("utf-8", errors="replace")
                f = str(f).strip().strip('"')
                if f.lower().endswith(".pdf") and f not in self.files:
                    self.files.append(f)
                    self.file_listbox.insert(tk.END, os.path.basename(f))
                    self._log(f"追加: {os.path.basename(f)}")
            self._update_drop_label()

        def _add_files(self):
            paths = filedialog.askopenfilenames(
                title="PDFファイルを選択",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
            )
            for p in paths:
                if p not in self.files:
                    self.files.append(p)
                    self.file_listbox.insert(tk.END, os.path.basename(p))
            self._update_drop_label()

        def _remove_selected(self):
            selected = list(self.file_listbox.curselection())
            for i in reversed(selected):
                self.file_listbox.delete(i)
                del self.files[i]
            self._update_drop_label()

        def _clear_files(self):
            self.files.clear()
            self.file_listbox.delete(0, tk.END)
            self._update_drop_label()

        def _update_drop_label(self):
            count = len(self.files)
            if count > 0:
                self.drop_label.configure(
                    text=f"{count} 個のPDFファイルが選択されています",
                    fg="#27ae60"
                )
            else:
                self.drop_label.configure(
                    text=("ここにPDFをドラッグ&ドロップ\n"
                          "または下のボタンでファイルを選択"
                          if has_windnd else
                          "下のボタンでファイルを選択してください"),
                    fg="#7f8c8d"
                )

        def _get_engine(self):
            val = self.engine_var.get()
            if val.startswith("pymupdf"):
                return "pymupdf"
            elif val.startswith("marker"):
                return "marker"
            return "auto"

        def _get_format(self):
            val = self.format_var.get()
            if val.startswith("md"):
                return "md"
            return "txt"

        def _start_extraction(self):
            if self.processing:
                return
            if not self.files:
                messagebox.showwarning("警告", "PDFファイルを選択してください。")
                return

            deps = check_dependencies()
            if not any(deps.values()):
                messagebox.showerror(
                    "エラー",
                    "エンジンがインストールされていません。\n"
                    "pip install pymupdf4llm を実行してください。"
                )
                return

            if self.output_var.get() == "custom":
                out_dir = self.output_dir_var.get()
                if not out_dir or not os.path.isdir(out_dir):
                    messagebox.showwarning(
                        "警告", "出力先フォルダを選択してください。"
                    )
                    return

            self.processing = True
            self.run_btn.configure(state=tk.DISABLED)
            self.progress.start(10)

            thread = threading.Thread(target=self._extraction_worker, daemon=True)
            thread.start()

        def _extraction_worker(self):
            engine = self._get_engine()
            out_format = self._get_format()
            output_mode = self.output_var.get()
            output_dir = (self.output_dir_var.get()
                          if output_mode == "custom" else None)
            total = len(self.files)
            success = 0
            failed = 0

            for i, filepath in enumerate(self.files):
                self.root.after(0, self.status_var.set,
                                f"処理中: {i+1}/{total}")

                # 出力パスを決定
                ext = ".md" if out_format == "md" else ".txt"
                if output_dir:
                    fname = Path(filepath).stem + ext
                    out_path = str(Path(output_dir) / fname)
                else:
                    out_path = str(Path(filepath).with_suffix(ext))

                try:
                    extract_and_save(
                        filepath,
                        output_path=out_path,
                        engine=engine,
                        output_format=out_format,
                        progress_callback=lambda msg: self.root.after(
                            0, self._log, msg
                        )
                    )
                    success += 1
                except Exception as e:
                    self.root.after(0, self._log, f"エラー: {e}")
                    failed += 1

            def finish():
                self.progress.stop()
                self.run_btn.configure(state=tk.NORMAL)
                self.processing = False
                self.status_var.set(
                    f"完了: 成功 {success} / 失敗 {failed} / 合計 {total}"
                )
                messagebox.showinfo(
                    "完了",
                    f"テキスト抽出が完了しました。\n"
                    f"成功: {success}\n失敗: {failed}\n合計: {total}"
                )

            self.root.after(0, finish)

    root = tk.Tk()
    app = App(root)  # noqa: F841
    root.mainloop()


# ---------------------------------------------------------------------------
# エントリポイント
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    run_cli()
