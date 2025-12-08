@echo off
echo === Starting PPTX Analyzer Servers ===
echo.

start "Hybrid Server (3001)" cmd /k "node hybrid_server.js"
timeout /t 1 /nobreak >nul

start "Format Converter Server (3002)" cmd /k "node convert_server.js"
timeout /t 1 /nobreak >nul

echo.
echo All servers started in separate windows!
echo.
echo Using Live Server at http://127.0.0.1:5500/
echo.
echo Press any key to exit this window (servers will keep running)...
pause >nul

