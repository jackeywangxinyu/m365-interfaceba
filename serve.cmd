@echo off
cd /d "%~dp0"
node src/cli.js serve --port 4141
pause
