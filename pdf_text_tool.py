"""
PDF テキスト変換ツール - PDFからWord/テキスト/Markdownに変換

エンジン構成:
  - pdf2docx:    PDF→Word直接変換（レイアウト・表・書式を保持）
  - pymupdf4llm: デジタルPDFの高速テキスト抽出（txt/md出力）
  - marker-pdf:  スキャン/画像PDFのOCR抽出（オプション、PyTorch使用）

出力形式:
  - docx: Word文書（デフォルト、高品質レイアウト保持）
  - txt:  プレーンテキスト
  - md:   Markdown

使い方:
  CLI:  python pdf_text_tool.py input.pdf
  CLI:  python pdf_text_tool.py input.pdf -f txt
  GUI:  python pdf_text_tool.py --gui
"""

import argparse
import os
import re
import sys
import threading
import time
from pathlib import Path


# ---------------------------------------------------------------------------
# バリデーション
# ---------------------------------------------------------------------------

def validate_pdf_path(input_path):
    """入力ファイルのバリデーション。問題があればエラーメッセージを返す。"""
    input_path = Path(input_path)
    if not input_path.exists():
        return f"ファイルが見つかりません: {input_path}"
    if input_path.suffix.lower() != ".pdf":
        return f"PDFファイルではありません: {input_path.name} (拡張子: {input_path.suffix})"
    return None


def is_pdf_encrypted(input_path):
    """PDFがパスワード保護されているか確認"""
    try:
        import pymupdf
        doc = pymupdf.open(str(input_path))
        encrypted = doc.is_encrypted
        doc.close()
        return encrypted
    except Exception:
        return False


# ---------------------------------------------------------------------------
# PDF種別判定
# ---------------------------------------------------------------------------

def detect_pdf_type(input_path):
    """PDFがデジタル（テキスト埋め込み）かスキャン（画像ベース）かを判定。"""
    try:
        import pymupdf
    except ImportError:
        return "unknown"

    try:
        doc = pymupdf.open(str(input_path))
    except Exception:
        return "unknown"

    total_pages = len(doc)
    if total_pages == 0:
        doc.close()
        return "unknown"

    pages_with_text = 0
    total_text_len = 0
    sample_pages = min(total_pages, 10)

    for i in range(sample_pages):
        page = doc[i]
        text = page.get_text().strip()
        text_len = len(text)
        total_text_len += text_len
        if text_len > 10:
            pages_with_text += 1

    doc.close()

    if pages_with_text > 0 or total_text_len > 20:
        return "digital"
    return "scanned"


# ---------------------------------------------------------------------------
# スキャンPDF用 OCR テキスト抽出 (Tesseract)
# ---------------------------------------------------------------------------

def _find_tesseract():
    """Tesseractのパスを探す"""
    import shutil
    t = shutil.which("tesseract")
    if t:
        return t
    candidates = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None


def _find_tessdata_best():
    """高品質 tessdata (tessdata_best) のパスを探す。なければダウンロード。"""
    local_dir = Path(__file__).parent
    jpn_best = local_dir / "jpn.traineddata"
    if jpn_best.exists() and jpn_best.stat().st_size > 10_000_000:
        return str(local_dir)
    return None


def ocr_scanned_pdf(input_path, progress_callback=None):
    """
    スキャンPDFの各ページを画像化 → 前処理 → Tesseract OCR でテキスト抽出。

    見開きスキャンの場合は左右に分割して処理する。
    """
    import subprocess
    import tempfile

    def log(msg):
        if progress_callback:
            progress_callback(msg)

    try:
        import pymupdf
    except ImportError:
        raise RuntimeError("pymupdf がインストールされていません。")

    tesseract_path = _find_tesseract()
    if not tesseract_path:
        raise RuntimeError(
            "Tesseract がインストールされていません。\n"
            "  → https://github.com/tesseract-ocr/tesseract からインストールしてください。"
        )

    tessdata_dir = _find_tessdata_best()
    tessdata_args = ["--tessdata-dir", tessdata_dir] if tessdata_dir else []

    try:
        from PIL import Image, ImageEnhance
    except ImportError:
        raise RuntimeError("Pillow がインストールされていません。\n  → pip install Pillow")

    log("スキャンPDFをOCR処理中（Tesseract）...")

    doc = pymupdf.open(str(input_path))
    total = len(doc)
    all_text = []

    for i in range(total):
        if (i + 1) % 10 == 0 or i == 0:
            log(f"  ページ {i+1}/{total} 処理中...")

        page = doc[i]
        images = page.get_images()

        if images:
            # PDF埋め込み画像を直接抽出（高品質）
            xref = images[0][0]
            img_data = doc.extract_image(xref)
            img_bytes = img_data["image"]
            img_w = img_data["width"]
            img_h = img_data["height"]

            from io import BytesIO
            img = Image.open(BytesIO(img_bytes))
        else:
            # 画像がない場合はページをレンダリング
            pix = page.get_pixmap(dpi=300)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            img_w, img_h = img.size

        # 見開き判定: 横長なら左右分割
        is_spread = img_w > img_h * 1.3
        if is_spread:
            mid = img_w // 2
            pages_to_ocr = [
                img.crop((0, 0, mid, img_h)),       # 左ページ
                img.crop((mid, 0, img_w, img_h)),    # 右ページ
            ]
        else:
            pages_to_ocr = [img]

        for sub_img in pages_to_ocr:
            # 前処理: グレースケール → コントラスト強調 → 二値化
            gray = sub_img.convert("L")
            gray = ImageEnhance.Contrast(gray).enhance(2.0)
            binary = gray.point(lambda x: 255 if x > 128 else 0, "L")

            # 一時ファイルに保存してTesseractを呼び出す
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                tmp_path = tmp.name
                binary.save(tmp_path)

            try:
                result = subprocess.run(
                    [tesseract_path, tmp_path, "stdout",
                     "-l", "jpn", "--psm", "5"] + tessdata_args,
                    capture_output=True, text=True, encoding="utf-8",
                    timeout=60
                )
                page_text = result.stdout.strip()
                if page_text:
                    all_text.append(page_text)
            except (subprocess.TimeoutExpired, Exception) as e:
                log(f"  ページ {i+1} OCRエラー: {e}")
            finally:
                try:
                    os.unlink(tmp_path)
                except OSError:
                    pass

    doc.close()

    full_text = "\n\n".join(all_text)
    log(f"OCR完了: {len(full_text):,}文字抽出")
    return full_text


# ---------------------------------------------------------------------------
# Word (.docx) 変換 — pdf2docx ライブラリ使用
# ---------------------------------------------------------------------------

def convert_to_docx(input_path, output_path, progress_callback=None):
    """
    pdf2docx を使用してPDFをWord文書に直接変換。
    PDFのレイアウト・表・書式をそのまま保持する。
    """
    def log(msg):
        if progress_callback:
            progress_callback(msg)

    try:
        from pdf2docx import Converter
    except ImportError:
        raise RuntimeError(
            "pdf2docx がインストールされていません。\n"
            "  → pip install pdf2docx を実行してください。"
        )

    log("pdf2docx でWord文書に変換中...")

    cv = Converter(str(input_path))
    cv.convert(str(output_path))
    cv.close()

    log(f"Word文書保存完了: {Path(output_path).name}")


def _save_text_as_docx(text, output_path, source_name="", progress_callback=None):
    """OCR抽出テキストをシンプルなWord文書として保存。"""
    def log(msg):
        if progress_callback:
            progress_callback(msg)

    try:
        from docx import Document
        from docx.shared import Pt
    except ImportError:
        raise RuntimeError("python-docx がインストールされていません。\n  → pip install python-docx")

    log("Word文書を生成中...")

    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Yu Gothic'
    style.font.size = Pt(10.5)

    if source_name:
        doc.add_heading(f"OCR変換: {source_name}", level=1)

    for para_text in text.split("\n\n"):
        para_text = para_text.strip()
        if para_text:
            doc.add_paragraph(para_text)

    doc.save(str(output_path))
    log(f"Word文書保存完了: {Path(output_path).name}")


# ---------------------------------------------------------------------------
# テキスト抽出エンジン
# ---------------------------------------------------------------------------

def extract_text_pymupdf(input_path, output_format="txt", progress_callback=None):
    """pymupdf4llm を使用してデジタルPDFからテキスト抽出。"""
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

    if output_format == "txt":
        log("pymupdf4llm.to_text() でテキスト抽出中...")
        text = pymupdf4llm.to_text(str(input_path), show_progress=False)
        log("テキスト変換完了")
        return text.strip()

    log("pymupdf4llm.to_markdown() でMarkdown抽出中...")
    md_text = pymupdf4llm.to_markdown(
        str(input_path), ignore_images=True, show_progress=False
    )
    log("Markdown変換完了")
    return md_text.strip()


def extract_text_marker(input_path, output_format="txt", progress_callback=None):
    """marker-pdf を使用してスキャン/画像PDFから高品質OCRテキスト抽出。"""
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
    """Markdownテキストからプレーンテキストに変換（marker-pdf出力用）"""
    text = md_text
    text = re.sub(r'!\[([^\]]*)\]\([^\)]+\)', r'\1', text)
    text = re.sub(r'```[^\n]*\n(.*?)```', r'\1', text, flags=re.DOTALL)
    text = re.sub(r'^\|[-:| ]+\|\s*$', '', text, flags=re.MULTILINE)
    text = re.sub(
        r'^\|(.+)\|\s*$',
        lambda m: '\t'.join(cell.strip() for cell in m.group(1).split('|')),
        text, flags=re.MULTILINE
    )
    text = re.sub(r'^#{1,6}\s+', '', text, flags=re.MULTILINE)
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)', r'\1', text)
    text = re.sub(r'__(.+?)__', r'\1', text)
    text = re.sub(r'`(.+?)`', r'\1', text)
    text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', text)
    text = re.sub(r'^[-*_]{3,}\s*$', '', text, flags=re.MULTILINE)
    text = re.sub(r'^\s*[-*+]\s+', '', text, flags=re.MULTILINE)
    text = re.sub(r'^\s*\d+\.\s+', '', text, flags=re.MULTILINE)
    text = re.sub(r'^>\s?', '', text, flags=re.MULTILINE)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


# ---------------------------------------------------------------------------
# 統合抽出関数（txt/md用）
# ---------------------------------------------------------------------------

def extract_text(input_path, engine="auto", output_format="txt",
                 progress_callback=None):
    """PDFからテキストを抽出する統合関数（txt/md出力用）。"""
    input_path = Path(input_path)

    def log(msg):
        if progress_callback:
            progress_callback(msg)

    # エンジン選択
    if engine == "auto":
        log("PDF種別を判定中...")
        pdf_type = detect_pdf_type(input_path)
        log(f"  判定結果: {pdf_type}")

        if pdf_type == "scanned":
            # スキャンPDF: Tesseract OCR → marker-pdf → pymupdf4llm の順で試す
            tesseract = _find_tesseract()
            if tesseract:
                engine = "tesseract"
                log("  → Tesseract OCR を使用します")
            else:
                try:
                    import marker  # noqa: F401
                    engine = "marker"
                    log("  → marker-pdf (高精度OCR) を使用します")
                except ImportError:
                    log("  → OCRエンジン未インストール。pymupdf4llm で試行します")
                    engine = "pymupdf"
        else:
            engine = "pymupdf"
            log("  → pymupdf4llm (高速) を使用します")

    if engine == "tesseract":
        return ocr_scanned_pdf(input_path, progress_callback)
    elif engine == "marker":
        return extract_text_marker(input_path, output_format, progress_callback)
    else:
        return extract_text_pymupdf(input_path, output_format, progress_callback)


# ---------------------------------------------------------------------------
# メイン変換関数
# ---------------------------------------------------------------------------

def extract_and_save(input_path, output_path=None, engine="auto",
                     output_format="docx", progress_callback=None):
    """PDFを変換してファイルに保存。"""
    input_path = Path(input_path)

    # バリデーション
    err = validate_pdf_path(input_path)
    if err:
        raise ValueError(err)

    if is_pdf_encrypted(input_path):
        raise ValueError(f"パスワード保護されたPDFです: {input_path.name}\n"
                         "  パスワードを解除してから再度お試しください。")

    # 出力パス決定
    ext_map = {"txt": ".txt", "md": ".md", "docx": ".docx"}
    ext = ext_map.get(output_format, ".docx")

    if output_path is None:
        output_path = input_path.with_suffix(ext)
    output_path = Path(output_path)

    # 入力と出力が同一ファイルになる場合の防止
    try:
        if input_path.resolve() == output_path.resolve():
            output_path = input_path.parent / (input_path.stem + "_converted" + ext)
    except (OSError, ValueError):
        pass

    def log(msg):
        if progress_callback:
            progress_callback(msg)

    log(f"処理開始: {input_path.name}")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    start_time = time.time()

    if output_format == "docx":
        pdf_type = detect_pdf_type(input_path)
        if pdf_type == "scanned":
            # スキャンPDF: OCR→テキスト→Word文書として保存
            log("スキャンPDFを検出。OCRでテキスト抽出してWord出力します...")
            text = ocr_scanned_pdf(input_path, progress_callback)
            _save_text_as_docx(text, output_path, input_path.name, progress_callback)
        else:
            # デジタルPDF: pdf2docxで直接変換
            convert_to_docx(input_path, output_path, progress_callback)
    else:
        text = extract_text(
            input_path, engine=engine, output_format=output_format,
            progress_callback=progress_callback
        )
        output_path.write_text(text, encoding="utf-8")

        char_count = len(text)
        line_count = text.count('\n') + 1 if text else 0
        log(f"保存完了: {output_path.name} ({char_count:,}文字, {line_count:,}行)")

        if not text or not text.strip():
            log("  警告: テキストが抽出できませんでした（空の結果）")

    elapsed = time.time() - start_time
    log(f"  処理時間: {elapsed:.1f}秒")

    return str(output_path)


# ---------------------------------------------------------------------------
# 依存関係チェック
# ---------------------------------------------------------------------------

def check_dependencies():
    """利用可能なエンジンを確認"""
    results = {}

    try:
        import pdf2docx  # noqa: F401
        results["pdf2docx"] = True
    except ImportError:
        results["pdf2docx"] = False

    try:
        import pymupdf4llm  # noqa: F401
        results["pymupdf4llm"] = True
    except ImportError:
        results["pymupdf4llm"] = False

    try:
        import marker  # noqa: F401
        results["marker"] = True
    except ImportError:
        results["marker"] = False

    results["tesseract"] = _find_tesseract() is not None

    return results


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def run_cli():
    parser = argparse.ArgumentParser(
        description="PDF テキスト変換ツール - PDFからWord/テキスト/Markdownに変換"
    )
    parser.add_argument("input", nargs="?", help="入力PDFファイルパス")
    parser.add_argument("-o", "--output", help="出力ファイルパス")
    parser.add_argument(
        "-e", "--engine", default="auto",
        choices=["auto", "pymupdf", "marker"],
        help="抽出エンジン (デフォルト: auto)"
    )
    parser.add_argument(
        "-f", "--format", default="docx",
        choices=["docx", "txt", "md"],
        help="出力形式 (デフォルト: docx)"
    )
    parser.add_argument("--gui", action="store_true", help="GUIモードで起動")
    parser.add_argument("--check", action="store_true", help="依存関係を確認して終了")

    args = parser.parse_args()

    if args.check:
        deps = check_dependencies()
        print("=== 依存関係チェック ===")
        for name, available in deps.items():
            status = "OK" if available else "未インストール"
            print(f"  {name}: {status}")
        if not any(deps.values()):
            print("\nエラー: 少なくとも1つのエンジンが必要です。")
            print("  → pip install pdf2docx pymupdf4llm")
            sys.exit(1)
        sys.exit(0)

    if args.gui or args.input is None:
        run_gui()
        return

    err = validate_pdf_path(args.input)
    if err:
        print(f"エラー: {err}", file=sys.stderr)
        sys.exit(1)

    deps = check_dependencies()
    if args.format == "docx" and not deps.get("pdf2docx"):
        print("エラー: Word出力にはpdf2docxが必要です。", file=sys.stderr)
        print("  → pip install pdf2docx", file=sys.stderr)
        sys.exit(1)
    if args.format in ("txt", "md") and not deps.get("pymupdf4llm"):
        print("エラー: テキスト出力にはpymupdf4llmが必要です。", file=sys.stderr)
        print("  → pip install pymupdf4llm", file=sys.stderr)
        sys.exit(1)

    try:
        output = extract_and_save(
            args.input, args.output,
            engine=args.engine, output_format=args.format,
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
            self.root.title("PDF 変換ツール")
            self.root.geometry("720x680")
            self.root.resizable(True, True)
            self.root.configure(bg="#f0f0f0")

            self.files = []
            self.processing = False

            self._build_ui()
            self._check_deps()

            if has_windnd:
                windnd.hook_dropfiles(self.root, func=self._on_drop)

        def _build_ui(self):
            title_frame = tk.Frame(self.root, bg="#1a5276", pady=10)
            title_frame.pack(fill=tk.X)
            tk.Label(
                title_frame, text="PDF 変換ツール",
                font=("Segoe UI", 16, "bold"), fg="white", bg="#1a5276"
            ).pack()
            tk.Label(
                title_frame,
                text="PDFからWord/テキスト/Markdownに変換",
                font=("Segoe UI", 9), fg="#aed6f1", bg="#1a5276"
            ).pack()

            # ドロップエリア
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
            ttk.Button(btn_frame, text="ファイルを追加...",
                       command=self._add_files).pack(side=tk.LEFT, padx=(0, 5))
            ttk.Button(btn_frame, text="選択を削除",
                       command=self._remove_selected).pack(side=tk.LEFT, padx=(0, 5))
            ttk.Button(btn_frame, text="すべてクリア",
                       command=self._clear_files).pack(side=tk.LEFT)

            # 設定
            settings_frame = tk.LabelFrame(
                self.root, text=" 設定 ",
                font=("Segoe UI", 10), padx=10, pady=5, bg="#f0f0f0"
            )
            settings_frame.pack(fill=tk.X, padx=15, pady=5)

            tk.Label(settings_frame, text="出力形式:",
                     font=("Segoe UI", 10), bg="#f0f0f0"
                     ).grid(row=0, column=0, sticky=tk.W, pady=2)

            self.format_var = tk.StringVar(value="docx (Word文書)")
            ttk.Combobox(
                settings_frame, textvariable=self.format_var,
                values=[
                    "docx (Word文書)",
                    "txt (プレーンテキスト)",
                    "md (Markdown)",
                ],
                width=35, state="readonly"
            ).grid(row=0, column=1, sticky=tk.W, padx=10, pady=2)

            tk.Label(settings_frame, text="出力先:",
                     font=("Segoe UI", 10), bg="#f0f0f0"
                     ).grid(row=1, column=0, sticky=tk.W, pady=2)

            output_row = tk.Frame(settings_frame, bg="#f0f0f0")
            output_row.grid(row=1, column=1, sticky=tk.W, padx=10, pady=2)

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

            # 実行ボタン
            exec_frame = tk.Frame(self.root, bg="#f0f0f0", pady=5)
            exec_frame.pack(fill=tk.X, padx=15)

            self.run_btn = ttk.Button(
                exec_frame, text="変換実行",
                command=self._start_extraction
            )
            self.run_btn.pack(fill=tk.X, ipady=5)

            self.progress = ttk.Progressbar(self.root, mode="indeterminate")
            self.progress.pack(fill=tk.X, padx=15, pady=(5, 0))

            # ログ
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

            self.status_var = tk.StringVar(value="準備完了")
            tk.Label(
                self.root, textvariable=self.status_var,
                font=("Segoe UI", 9), fg="#7f8c8d", bg="#f0f0f0",
                anchor=tk.W
            ).pack(fill=tk.X, padx=15, pady=(0, 5))

        def _check_deps(self):
            deps = check_dependencies()
            available = [n for n, ok in deps.items() if ok]
            missing = [n for n, ok in deps.items() if not ok]
            if available:
                self._log(f"利用可能: {', '.join(available)}")
            if missing:
                self._log(f"未インストール: {', '.join(missing)}")
            if not any(deps.values()):
                self._log("エラー: エンジンが1つもありません！")
                self._log("  → pip install pdf2docx pymupdf4llm")
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
            for i in reversed(list(self.file_listbox.curselection())):
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

        def _get_format(self):
            val = self.format_var.get()
            if val.startswith("txt"):
                return "txt"
            elif val.startswith("md"):
                return "md"
            return "docx"

        def _start_extraction(self):
            if self.processing:
                return
            if not self.files:
                messagebox.showwarning("警告", "PDFファイルを選択してください。")
                return

            deps = check_dependencies()
            fmt = self._get_format()

            if fmt == "docx" and not deps.get("pdf2docx"):
                messagebox.showerror("エラー",
                                     "Word出力にはpdf2docxが必要です。\n"
                                     "pip install pdf2docx を実行してください。")
                return
            if fmt in ("txt", "md") and not deps.get("pymupdf4llm"):
                messagebox.showerror("エラー",
                                     "テキスト出力にはpymupdf4llmが必要です。\n"
                                     "pip install pymupdf4llm を実行してください。")
                return

            if self.output_var.get() == "custom":
                out_dir = self.output_dir_var.get()
                if not out_dir or not os.path.isdir(out_dir):
                    messagebox.showwarning("警告", "出力先フォルダを選択してください。")
                    return

            self.processing = True
            self.run_btn.configure(state=tk.DISABLED)
            self.progress.start(10)

            thread = threading.Thread(target=self._worker, daemon=True)
            thread.start()

        def _worker(self):
            out_format = self._get_format()
            output_mode = self.output_var.get()
            output_dir = (self.output_dir_var.get()
                          if output_mode == "custom" else None)
            total = len(self.files)
            success = 0
            failed = 0
            overall_start = time.time()

            ext_map = {"txt": ".txt", "md": ".md", "docx": ".docx"}
            ext = ext_map.get(out_format, ".docx")

            for i, filepath in enumerate(self.files):
                self.root.after(0, self.status_var.set,
                                f"処理中: {i+1}/{total}")

                if output_dir:
                    fname = Path(filepath).stem + ext
                    out_path = str(Path(output_dir) / fname)
                else:
                    out_path = str(Path(filepath).with_suffix(ext))

                try:
                    extract_and_save(
                        filepath, output_path=out_path,
                        output_format=out_format,
                        progress_callback=lambda msg: self.root.after(
                            0, self._log, msg
                        )
                    )
                    success += 1
                except Exception as e:
                    self.root.after(0, self._log, f"エラー: {e}")
                    failed += 1

            overall_elapsed = time.time() - overall_start

            def finish():
                self.progress.stop()
                self.run_btn.configure(state=tk.NORMAL)
                self.processing = False
                self.status_var.set(
                    f"完了: 成功 {success} / 失敗 {failed} / "
                    f"合計 {total} ({overall_elapsed:.1f}秒)"
                )
                self._log(f"全体処理時間: {overall_elapsed:.1f}秒")
                messagebox.showinfo(
                    "完了",
                    f"変換が完了しました。\n"
                    f"成功: {success}\n失敗: {failed}\n合計: {total}\n"
                    f"処理時間: {overall_elapsed:.1f}秒"
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
