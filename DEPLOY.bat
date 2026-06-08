@echo off
cd /d "C:\Users\whaughey\OneDrive - Refined Technologies Inc\Documents\Claude\Projects\Azzurri Storm"
echo.
echo === Azzurri Storm Deploy ===
echo Pushing updates to GitHub...
git add -A
git commit -m "Update %DATE% %TIME%"
git push origin main
echo.
echo Done! Site will be live in ~60 seconds.
pause
