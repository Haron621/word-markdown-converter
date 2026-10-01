@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist "dist\ru\index.html" (
  echo Сборка приложения не найдена. Выполняю локальную сборку...
  set ASTRO_TELEMETRY_DISABLED=1
  call npm run build
  if errorlevel 1 (
    echo Ошибка сборки.
    pause
    exit /b 1
  )
)

start "" "http://127.0.0.1:3000/ru/"
node local-server.mjs

