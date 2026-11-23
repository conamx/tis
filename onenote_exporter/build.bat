@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo [1/2] 필요한 패키지 설치 중...
python -m pip install -r requirements.txt pyinstaller || goto :error

echo [2/2] exe 만드는 중... (1~3분)
python -m PyInstaller --noconfirm --onefile --noconsole --name OneNoteExporter --icon icon.ico ^
  --add-data "icon.png;." --add-data "icon.ico;." ^
  --add-data "projects;projects" ^
  main.py || goto :error

echo.
echo 완료! dist\OneNoteExporter.exe 를 원하는 폴더로 옮겨서 실행하세요.
echo (처음 실행하면 exe 옆에 projects 폴더와 config.json 이 생깁니다)
pause
exit /b 0

:error
echo.
echo 빌드 실패. 위 오류 메시지를 확인하세요.
pause
exit /b 1
