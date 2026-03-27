@echo off
chcp 65001 >nul
echo ============================================
echo   PDF テキスト変換ツール セットアップ
echo ============================================
echo.

REM --- Python パッケージ（メインエンジン） ---
echo [1/3] pymupdf4llm をインストール中（高速テキスト抽出）...
pip install pymupdf4llm windnd
echo.

REM --- Tesseract インストール（pymupdfのOCRフォールバック用） ---
echo [2/3] Tesseract OCR をインストール中...
winget install --id "UB-Mannheim.TesseractOCR" --accept-package-agreements --accept-source-agreements
echo.

REM --- 日本語データのインストール ---
set TESSDATA=C:\Program Files\Tesseract-OCR\tessdata

if exist "%TESSDATA%\jpn.traineddata" (
    echo   日本語データ: 既にインストール済み
) else (
    echo   日本語データをダウンロード中...
    curl -L -o "%TEMP%\jpn.traineddata" "https://github.com/tesseract-ocr/tessdata_best/raw/main/jpn.traineddata"
    echo   コピー中（管理者権限が必要な場合があります）...
    copy "%TEMP%\jpn.traineddata" "%TESSDATA%\jpn.traineddata" >nul 2>&1
    if errorlevel 1 (
        echo   管理者権限でコピーします...
        powershell -Command "Start-Process cmd -ArgumentList '/c copy %TEMP%\jpn.traineddata \"%TESSDATA%\jpn.traineddata\"' -Verb RunAs -Wait"
    )
    echo   日本語データ: OK
)
echo.

REM --- marker-pdf（オプション・高精度OCR） ---
echo [3/3] marker-pdf のインストール（オプション）
echo   ※ スキャンPDFの高精度OCRが必要な場合のみ
echo   ※ PyTorch が必要です（GPU推奨、CPUでも動作可）
echo.
set /p INSTALL_MARKER="marker-pdf をインストールしますか？ (y/N): "
if /i "%INSTALL_MARKER%"=="y" (
    echo   marker-pdf をインストール中...
    pip install marker-pdf
    echo   marker-pdf: OK
) else (
    echo   marker-pdf: スキップ（後から pip install marker-pdf で追加可能）
)
echo.

echo ============================================
echo   セットアップ完了
echo ============================================
echo.
echo 依存関係を確認するには:
echo   python pdf_text_tool.py --check
echo.
echo GUI起動:
echo   run_gui.bat または python pdf_text_tool.py --gui
echo.
pause
