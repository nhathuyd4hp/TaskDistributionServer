@echo off
set N=%1
if "%N%"=="" set N=8

echo ===== GIT PULL =====
git pull origin main
if errorlevel 1 (
    pause
    exit /b 1
)

echo ===== SYNC ENV =====
uv sync
if errorlevel 1 (
    pause
    exit /b 1
)

echo ===== START SERVER =====
start "API Server" ./.venv/Scripts/python -m uvicorn src.main:app --host 0.0.0.0 --port 8000

echo ===== START WORKER =====
for /L %%i in (1,1,%N%) do (
    start "Worker-%%i" ./.venv/Scripts/python -m celery -A src.worker.Worker worker ^
        --hostname=worker%%i@%COMPUTERNAME% ^
        --pool=solo ^
        --concurrency=1 ^
        --prefetch-multiplier=1 ^
        --max-tasks-per-child=1
)
