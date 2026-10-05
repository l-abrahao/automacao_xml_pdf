@echo off
title Analise XML e PDF - NF-e
cd /d "%~dp0"

echo.
echo ============================================================
echo   ANALISE DE XML + PDF - NF-e
echo ============================================================
echo.

python processar_xml_pdf_nf.py

echo.
pause
