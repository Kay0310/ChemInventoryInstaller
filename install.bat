@echo off
chcp 65001 > nul
echo [STEP 1] Python 가상환경 생성 중...
where python > nul
if errorlevel 1 (
    echo [오류] Python이 설치되어 있지 않거나 환경 변수에 등록되지 않았습니다.
    echo https://www.python.org 에서 Python을 설치하고 다시 시도하세요.
    pause
    exit /b
)

python -m venv venv

echo [STEP 2] 가상환경 활성화...
call venv\Scripts\activate.bat

echo [STEP 3] 필수 패키지 설치 중...
pip install --upgrade pip
pip install -r requirements.txt

echo.
echo ✅ 설치가 완료되었습니다!
echo ----------------------------------------
echo [가상환경 활성화 방법]
echo call venv\Scripts\activate.bat
echo [프로그램 실행 방법]
echo python msds_to_excel.py
pause
