for /f "tokens=5" %%i in ('netstat -ano ^| findstr :8001') do taskkill /F /PID %%i /T
pause