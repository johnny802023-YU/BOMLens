@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0\.."
title BOMLens Offline

echo BOMLens 離線模式
echo 資料只在本機處理；請勿關閉此視窗。
echo.

where node >nul 2>&1
if errorlevel 1 goto NO_NODE

where npm >nul 2>&1
if errorlevel 1 goto NO_NODE

if not exist "node_modules\" goto NO_MODULES

call npm run build
if errorlevel 1 goto BUILD_FAILED

start "" /B powershell.exe -NoProfile -WindowStyle Hidden -Command "$url='http://127.0.0.1:3784/'; for($i=0; $i -lt 30; $i++){ try { [void](Invoke-WebRequest -UseBasicParsing $url -TimeoutSec 1); Start-Process $url; break } catch { Start-Sleep -Seconds 1 } }"

echo 啟動完成後，瀏覽器會自動開啟：
echo http://127.0.0.1:3784/
echo.
echo 使用完畢請在此視窗按 Ctrl+C 停止服務。
echo.

call npm run offline:serve
set EXIT_CODE=%errorlevel%
if not "%EXIT_CODE%"=="0" goto SERVER_FAILED
exit /b 0

:NO_NODE
echo 找不到 Node.js 或 npm。
echo 請先請 IT 協助安裝 Node.js 22 或更新版本。
goto WAIT_AND_EXIT

:NO_MODULES
echo 缺少程式套件。
echo 請先在可連線環境開啟此資料夾，執行：npm ci
goto WAIT_AND_EXIT

:BUILD_FAILED
echo BOMLens 建置失敗，請把此視窗內容交給 IT 人員。
goto WAIT_AND_EXIT

:SERVER_FAILED
echo BOMLens 服務已停止或啟動失敗。

:WAIT_AND_EXIT
echo.
pause
exit /b 1
