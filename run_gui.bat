@echo off
chcp 65001 >nul
cd /d "%~dp0"
python pdf_text_tool.py --gui
