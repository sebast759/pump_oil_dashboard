@echo off
echo ========================================
echo Pushing markets_dashboard to GitHub...
echo ========================================

cd /d "%~dp0"
git add .
git commit -m "Update: %date% %time%"
git pull
git push

echo.
echo ========================================
echo Pushing sebast759.github.io to GitHub...
echo ========================================

