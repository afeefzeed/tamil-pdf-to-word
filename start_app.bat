@echo off
title Tamil PDF to Word Converter
echo ========================================================
echo   Tamil PDF to Word Converter
echo ========================================================
echo.
echo Starting local web server...
start "" http://localhost:5000
python "%~dp0app.py"
pause
