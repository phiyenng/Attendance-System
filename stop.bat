@echo off
chcp 65001 >nul

echo 🛑 Đang dừng Attendance System...

REM Kiểm tra Docker đã chạy chưa
docker info >nul 2>&1
if errorlevel 1 (
    echo ❌ Docker chưa chạy hoặc đã dừng.
    pause
    exit /b 1
)

REM Dừng và xóa containers
echo 🔄 Đang dừng containers...
docker-compose down

if errorlevel 0 (
    echo ✅ Attendance System đã dừng thành công!
    echo.
    echo 📋 Để khởi động lại:
    echo    ▶️  Chạy: start.bat
    echo    hoặc: docker-compose up -d
) else (
    echo ❌ Có lỗi khi dừng containers
    docker-compose ps
)

echo.
echo Nhấn phím bất kỳ để đóng cửa sổ này...
pause >nul
