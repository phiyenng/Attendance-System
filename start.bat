@echo off
chcp 65001 >nul

echo 🚀 Đang khởi chạy Attendance System...

REM Kiểm tra Docker đã chạy chưa
docker info >nul 2>&1
if errorlevel 1 (
    echo ❌ Docker chưa chạy. Vui lòng khởi động Docker Desktop trước.
    pause
    exit /b 1
)

echo ✅ Docker đã sẵn sàng

REM Kiểm tra file docker-compose.yml
if not exist "docker-compose.yml" (
    echo ❌ Không tìm thấy file docker-compose.yml
    pause
    exit /b 1
)

REM Dừng container cũ nếu có
echo 🔄 Dừng container cũ (nếu có)...
docker-compose down >nul 2>&1

REM Build và khởi chạy
echo 🔨 Đang build và khởi chạy container...
docker-compose up -d --build

if errorlevel 0 (
    echo ✅ Attendance System đã khởi chạy thành công!
    echo.
    echo 📝 Thông tin truy cập:
    echo    🌐 URL: http://localhost:5000
    echo    📁 Dữ liệu: .\uploads\
    echo.
    echo 📋 Các lệnh hữu ích:
    echo    🔍 Xem logs:     docker-compose logs -f
    echo    ⏹️  Dừng app:     docker-compose down
    echo    🔄 Restart:      docker-compose restart
    echo.
    echo ⏳ Đang đợi ứng dụng khởi động hoàn tất...
    
    REM Đợi ứng dụng sẵn sàng
    set /a max_attempts=30
    set /a attempt=0
    
    :wait_loop
    if %attempt% geq %max_attempts% goto timeout
    
    curl -s http://localhost:5000 >nul 2>&1
    if errorlevel 0 (
        echo 🎉 Ứng dụng đã sẵn sàng! Truy cập http://localhost:5000
        goto end
    )
    
    set /a attempt+=1
    echo|set /p="."
    timeout /t 2 /nobreak >nul
    goto wait_loop
    
    :timeout
    echo.
    echo ⚠️  Ứng dụng khởi động chậm hơn dự kiến. Vui lòng kiểm tra logs:
    echo    docker-compose logs
    goto end
    
    :end
    echo.
    echo Nhấn phím bất kỳ để đóng cửa sổ này...
    pause >nul
) else (
    echo ❌ Có lỗi khi khởi chạy. Kiểm tra logs:
    docker-compose logs
    pause
    exit /b 1
)
