@echo off
REM ============================================================
REM Limpeza do repositorio depurador - gerado pelo Claude
REM Remove: arquivos-lixo de terminal, codigo duplicado,
REM caches Python, venv versionado e backups.
REM Execute a partir da pasta depurador-main.
REM ============================================================
cd /d "%~dp0"

echo.
echo === Removendo arquivos-lixo (fragmentos de comandos salvos por acidente) ===
del /q ".append(str(ex))" 2>nul
del /q "cript..._)" 2>nul
del /q "et_value('dummy')" 2>nul
del /q "f_" 2>nul
del /q "qlite3; conn = sqlite3.connect('orcamentos.db'); cursor = conn.cursor(); cursor.execute(_SELECT name FROM sqlite_master WHERE type='table';_); print(cursor.fetchall())_" 2>nul
del /q "t 120 lines) ===_" 2>nul
del /q "t execution__)" 2>nul
del /q "tatus --short" 2>nul
del /q "tatus..._)" 2>nul
del /q "treamlit.testing.v1 import AppTest; at = AppTest.from_file(_app.py_); at.run(); fu = at.file_uploader[0]; import inspect; print(inspect.signature(fu.upload)); print(fu.upload.__doc__)'" 2>nul
del /q "_app" 2>nul
del /q "t.py" 2>nul
del /q "t_script.py" 2>nul
del /q "app.py.backup" 2>nul

echo === Removendo copia duplicada do codigo (orcamentos_app/) ===
rmdir /s /q "orcamentos_app" 2>nul

echo === Removendo caches Python ===
rmdir /s /q "__pycache__" 2>nul

echo === Removendo ambiente virtual versionado (.venv) ===
rmdir /s /q ".venv" 2>nul

echo.
echo Limpeza concluida. Recomendado: recriar o venv com
echo    python -m venv .venv ^&^& .venv\Scripts\activate ^&^& pip install -r requirements.txt
echo.
pause
