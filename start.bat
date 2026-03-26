@echo off
echo ============================
echo  IndiaScreen Stock Screener
echo ============================

echo.
echo [1/3] Setting up Python virtual environment...
cd backend
if not exist venv (
    python -m venv venv
)
call venv\Scripts\activate

echo [2/3] Installing Python dependencies...
pip install -r requirements.txt --quiet

echo [3/3] Starting backend (FastAPI)...
start cmd /k "cd /d %~dp0backend && call venv\Scripts\activate && uvicorn app.main:app --reload --port 8000"

echo.
echo [4/4] Installing and starting frontend...
cd ..\frontend
if not exist node_modules (
    npm install
)
start cmd /k "cd /d %~dp0frontend && npm run dev"

echo.
echo ============================
echo  Backend:  http://localhost:8000
echo  Frontend: http://localhost:5173
echo  API Docs: http://localhost:8000/docs
echo ============================
pause
