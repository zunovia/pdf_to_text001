"""
PDF OCR Tool - スキャンPDFを検索可能PDFに変換
書籍スキャンPDFにテキストレイヤーを追加し、検索可能にします。
画像はそのまま保持されます。

使い方:
  CLI:  python pdf_ocr_tool.py input.pdf
  GUI:  python pdf_ocr_tool.py --gui
"""

import argparse
import os
import subprocess
import sys
import threading
from pathlib import Path


# ---------------------------------------------------------------------------
# コア OCR 処理
# ---------------------------------------------------------------------------

def check_dependencies():
    """Tesseract がインストール済みか確認"""
    errors = []

    # Tesseract
    tesseract_path = find_tesseract()
    if tesseract_path is None:
        errors.append(
            "Tesseract OCR が見つかりません。\n"
            "  → setup.bat を実行するか、手動でインストールしてください。\n"
            "  → https://github.com/UB-Mannheim/tesseract/wiki"
        )

    return errors


def find_tesseract():
    """Tesseract の実行パスを探す"""
    # PATH にあるか
    try:
        result = subprocess.run(
            ["tesseract", "--version"],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0:
            return "tesseract"
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass

    # よくあるインストール先を確認
    common_paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    for p in common_paths:
        if os.path.isfile(p):
            return p

    return None


def find_ghostscript():
    """Ghostscript の実行パスを探す"""
    try:
        result = subprocess.run(
            ["gswin64c", "--version"],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0:
            return "gswin64c"
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass

    try:
        result = subprocess.run(
            ["gswin32c", "--version"],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0:
            return "gswin32c"
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass

    # よくあるインストール先
    import glob
    for pattern in [
        r"C:\Program Files\gs\gs*\bin\gswin64c.exe",
        r"C:\Program Files (x86)\gs\gs*\bin\gswin32c.exe",
    ]:
        matches = glob.glob(pattern)
        if matches:
            return matches[0]

    return None


def check_tesseract_languages():
    """インストール済みの Tesseract 言語データを確認"""
    tesseract = find_tesseract()
    if tesseract is None:
        return []
    try:
        result = subprocess.run(
            [tesseract, "--list-langs"],
            capture_output=True, text=True, timeout=10
        )
        langs = result.stdout.strip().split("\n")[1:]  # 1行目はヘッダ
        return [lang.strip() for lang in langs if lang.strip()]
    except Exception:
        return []


def run_ocr(input_path, output_path=None, languages="jpn+eng",
            progress_callback=None):
    """
    OCR を実行して検索可能 PDF を生成する。

    Args:
        input_path: 入力PDFパス
        output_path: 出力PDFパス (None なら _ocr 付きで自動生成)
        languages: Tesseract 言語コード (例: "jpn+eng")
        progress_callback: 進捗コールバック fn(message: str)
    Returns:
        output_path (str)
    Raises:
        RuntimeError: OCR 失敗時
    """
    input_path = Path(input_path)
    if not input_path.exists():
        raise FileNotFoundError(f"ファイルが見つかりません: {input_path}")

    if output_path is None:
        output_path = input_path.with_stem(input_path.stem + "_ocr")
    output_path = Path(output_path)

    def log(msg):
        if progress_callback:
            progress_callback(msg)

    log(f"処理開始: {input_path.name}")

    # ocrmypdf をインポート
    try:
        import ocrmypdf
    except ImportError:
        raise RuntimeError(
            "ocrmypdf がインストールされていません。\n"
            "  → pip install ocrmypdf を実行してください。"
        )

    # Tesseract パスを設定
    tesseract = find_tesseract()
    if tesseract and tesseract != "tesseract":
        os.environ["TESSERACT"] = tesseract
        tess_dir = os.path.dirname(tesseract)
        if tess_dir not in os.environ.get("PATH", ""):
            os.environ["PATH"] = tess_dir + os.pathsep + os.environ["PATH"]

    log(f"OCR 実行中 (言語: {languages}) ...")

    try:
        result = ocrmypdf.ocr(
            input_file=str(input_path),
            output_file=str(output_path),
            language=languages,
            skip_text=True,          # 既にテキストがあるページはスキップ
            optimize=0,              # Ghostscript不要にするため最適化オフ
            progress_bar=False,      # 自前で管理
            jobs=os.cpu_count() or 2,
        )
    except ocrmypdf.exceptions.PriorOcrFoundError:
        log("このPDFは既にテキストレイヤーを持っています。スキップしました。")
        # 入力をそのままコピー
        import shutil
        shutil.copy2(str(input_path), str(output_path))
    except Exception as e:
        raise RuntimeError(f"OCR に失敗しました: {e}")

    log(f"完了: {output_path.name}")
    return str(output_path)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def run_cli():
    parser = argparse.ArgumentParser(
        description="PDF OCR Tool - スキャンPDFを検索可能PDFに変換"
    )
    parser.add_argument(
        "input", nargs="?",
        help="入力PDFファイルパス"
    )
    parser.add_argument(
        "-o", "--output",
        help="出力PDFファイルパス (省略時: <入力名>_ocr.pdf)"
    )
    parser.add_argument(
        "-l", "--lang", default="jpn+eng",
        help="OCR言語 (デフォルト: jpn+eng)"
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
        errors = check_dependencies()
        if errors:
            print("=== 問題が見つかりました ===")
            for e in errors:
                print(f"\n{e}")
            sys.exit(1)
        else:
            langs = check_tesseract_languages()
            print("OK: すべての依存関係が揃っています。")
            print(f"  Tesseract 言語: {', '.join(langs)}")
            sys.exit(0)

    # GUIモード
    if args.gui or args.input is None:
        run_gui()
        return

    # CLI 実行
    errors = check_dependencies()
    if errors:
        print("エラー: 依存関係が不足しています。", file=sys.stderr)
        for e in errors:
            print(f"\n{e}", file=sys.stderr)
        print("\n→ python pdf_ocr_tool.py --check で詳細を確認できます。")
        sys.exit(1)

    try:
        output = run_ocr(
            args.input,
            args.output,
            args.lang,
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

    # ドラッグ&ドロップ対応（windnd が使えれば）
    has_windnd = False
    try:
        import windnd
        has_windnd = True
    except ImportError:
        pass

    class App:
        def __init__(self, root):
            self.root = root
            self.root.title("PDF OCR Tool")
            self.root.geometry("700x580")
            self.root.resizable(True, True)
            self.root.configure(bg="#f0f0f0")

            self.files = []
            self.processing = False

            self._build_ui()
            self._check_deps()

            # ドラッグ&ドロップ
            if has_windnd:
                windnd.hook_dropfiles(self.root, func=self._on_drop)

        def _build_ui(self):
            # --- タイトル ---
            title_frame = tk.Frame(self.root, bg="#2c3e50", pady=10)
            title_frame.pack(fill=tk.X)
            tk.Label(
                title_frame, text="PDF OCR Tool",
                font=("Segoe UI", 16, "bold"), fg="white", bg="#2c3e50"
            ).pack()
            tk.Label(
                title_frame,
                text="スキャンPDFを検索可能PDFに変換",
                font=("Segoe UI", 9), fg="#bdc3c7", bg="#2c3e50"
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
                bg="#ecf0f1", fg="#7f8c8d",
                relief=tk.GROOVE, bd=2,
                width=50, height=4
            )
            self.drop_label.pack(fill=tk.BOTH, expand=True, pady=(5, 5))

            # ファイルリスト
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

            tk.Label(
                settings_frame, text="OCR言語:",
                font=("Segoe UI", 10), bg="#f0f0f0"
            ).grid(row=0, column=0, sticky=tk.W, pady=2)

            self.lang_var = tk.StringVar(value="jpn+eng")
            lang_combo = ttk.Combobox(
                settings_frame, textvariable=self.lang_var,
                values=["jpn+eng", "jpn", "eng", "jpn+eng+chi_sim"],
                width=20, state="readonly"
            )
            lang_combo.grid(row=0, column=1, sticky=tk.W, padx=10, pady=2)

            tk.Label(
                settings_frame, text="出力先:",
                font=("Segoe UI", 10), bg="#f0f0f0"
            ).grid(row=1, column=0, sticky=tk.W, pady=2)

            self.output_var = tk.StringVar(value="same")
            output_row = tk.Frame(settings_frame, bg="#f0f0f0")
            output_row.grid(row=1, column=1, sticky=tk.W, padx=10, pady=2)

            ttk.Radiobutton(
                output_row, text="入力と同じフォルダ (_ocr付き)",
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
                custom_row, textvariable=self.output_dir_var, width=30,
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
                exec_frame, text="OCR 実行",
                command=self._start_ocr
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
                log_frame, height=5,
                font=("Consolas", 9), state=tk.DISABLED, wrap=tk.WORD
            )
            self.log_text.pack(fill=tk.BOTH, expand=True)

            # --- ステータスバー ---
            self.status_var = tk.StringVar(value="準備完了")
            tk.Label(
                self.root, textvariable=self.status_var,
                font=("Segoe UI", 9), fg="#7f8c8d", bg="#f0f0f0",
                anchor=tk.W
            ).pack(fill=tk.X, padx=15, pady=(0, 5))

        def _check_deps(self):
            """起動時に依存関係を確認"""
            errors = check_dependencies()
            if errors:
                self._log("警告: 依存関係に問題があります。")
                for e in errors:
                    self._log(e)
                self._log("")
                self._log("setup.bat を実行してセットアップしてください。")
                self.status_var.set("依存関係に問題があります")
            else:
                langs = check_tesseract_languages()
                self._log(f"Tesseract 言語: {', '.join(langs)}")
                if "jpn" not in langs:
                    self._log("警告: 日本語データ(jpn)がありません。")
                    self._log("  → setup.bat を実行してインストールしてください。")

        def _log(self, msg):
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)
            self.log_text.configure(state=tk.DISABLED)

        def _on_output_mode_change(self):
            """出力先モード切替"""
            if self.output_var.get() == "custom":
                self.output_dir_entry.configure(state=tk.NORMAL)
                self.output_dir_btn.configure(state=tk.NORMAL)
                if not self.output_dir_var.get():
                    self._browse_output_dir()
            else:
                self.output_dir_entry.configure(state=tk.DISABLED)
                self.output_dir_btn.configure(state=tk.DISABLED)

        def _browse_output_dir(self):
            """出力先フォルダを選択"""
            d = filedialog.askdirectory(title="出力先フォルダを選択")
            if d:
                self.output_dir_var.set(d)

        def _on_drop(self, files):
            """ドラッグ&ドロップで受け取ったファイルを追加"""
            for f in files:
                if isinstance(f, bytes):
                    # Windows日本語環境ではcp932でエンコードされる場合がある
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

        def _start_ocr(self):
            if self.processing:
                return
            if not self.files:
                messagebox.showwarning("警告", "PDFファイルを選択してください。")
                return

            errors = check_dependencies()
            if errors:
                messagebox.showerror(
                    "エラー",
                    "依存関係が不足しています。\n"
                    "setup.bat を実行してセットアップしてください。"
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

            thread = threading.Thread(target=self._ocr_worker, daemon=True)
            thread.start()

        def _ocr_worker(self):
            lang = self.lang_var.get()
            output_mode = self.output_var.get()
            output_dir = self.output_dir_var.get() if output_mode == "custom" else None
            total = len(self.files)
            success = 0
            failed = 0

            for i, filepath in enumerate(self.files):
                self.root.after(0, self.status_var.set,
                                f"処理中: {i+1}/{total}")

                # 出力パスを決定
                out_path = None
                if output_dir:
                    fname = Path(filepath).stem + "_ocr.pdf"
                    out_path = str(Path(output_dir) / fname)

                try:
                    run_ocr(
                        filepath,
                        output_path=out_path,
                        languages=lang,
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
                    f"OCR処理が完了しました。\n"
                    f"成功: {success}\n失敗: {failed}\n合計: {total}"
                )

            self.root.after(0, finish)

    root = tk.Tk()
    app = App(root)
    root.mainloop()


# ---------------------------------------------------------------------------
# エントリポイント
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    run_cli()
