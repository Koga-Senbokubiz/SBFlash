@echo off
setlocal EnableExtensions

REM バッチファイル名（拡張子なし）取得
set SCRIPT_NAME=%~n0

REM Pythonコマンド（環境によってpyの方が通る場合あり）
set PYTHON_CMD=python

REM 実行
%PYTHON_CMD% "%SCRIPT_NAME%.py"

endlocal
rem pause
exit
