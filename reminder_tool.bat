@echo off
echo "call reminder tool..."
powershell -NoProfile -ExecutionPolicy Unrestricted .\reminder_tool.ps1
pause > nul
exit

