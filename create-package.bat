@echo off
chcp 65001 >nul

echo 📦 Đang tạo package để chuyển sang máy khác...

REM Tạo thư mục package
set PACKAGE_NAME=attendance-system-package-%date:~-4,4%%date:~-10,2%%date:~-7,2%
mkdir "%PACKAGE_NAME%" 2>nul

echo 🔨 Đang build Docker image...
docker build -t attendance-system .

if errorlevel 1 (
    echo ❌ Lỗi khi build Docker image
    pause
    exit /b 1
)

echo 💾 Đang export Docker image...
docker save attendance-system -o "%PACKAGE_NAME%\attendance-system.tar"

echo 📁 Đang copy files...
REM Copy essential files
copy "docker-compose.yml" "%PACKAGE_NAME%\"
copy "DOCKER_README.md" "%PACKAGE_NAME%\"
copy "start.bat" "%PACKAGE_NAME%\"
copy "stop.bat" "%PACKAGE_NAME%\"

REM Copy data directories nếu có
if exist "uploads" (
    echo 📂 Copy thư mục uploads...
    xcopy "uploads" "%PACKAGE_NAME%\uploads\" /E /I /Q
)

if exist "rawdata" (
    echo 📂 Copy thư mục rawdata...
    xcopy "rawdata" "%PACKAGE_NAME%\rawdata\" /E /I /Q
)

REM Tạo file hướng dẫn
echo 📝 Tạo file hướng dẫn cài đặt...
(
echo # HƯỚNG DẪN CÀI ĐẶT ATTENDANCE SYSTEM
echo.
echo ## Yêu cầu hệ thống:
echo - Docker Desktop đã được cài đặt và đang chạy
echo.
echo ## Các bước cài đặt:
echo.
echo ### Bước 1: Import Docker image
echo ```bash
echo docker load -i attendance-system.tar
echo ```
echo.
echo ### Bước 2: Khởi chạy ứng dụng
echo ```bash
echo # Cách 1: Sử dụng script tự động
echo start.bat
echo.
echo # Cách 2: Sử dụng Docker Compose
echo docker-compose up -d
echo ```
echo.
echo ### Bước 3: Truy cập ứng dụng
echo Mở trình duyệt và truy cập: http://localhost:5000
echo.
echo ## Quản lý:
echo - Dừng ứng dụng: chạy stop.bat hoặc docker-compose down
echo - Xem logs: docker-compose logs -f
echo - Restart: docker-compose restart
echo.
echo ## Dữ liệu:
echo - Uploads: ./uploads/
echo - Raw data: ./rawdata/
echo.
echo Xem chi tiết trong DOCKER_README.md
) > "%PACKAGE_NAME%\INSTALL.md"

echo 📦 Đang nén package...
powershell -command "Compress-Archive -Path '%PACKAGE_NAME%' -DestinationPath '%PACKAGE_NAME%.zip' -Force"

REM Xóa thư mục tạm
rmdir /s /q "%PACKAGE_NAME%"

echo ✅ Package đã được tạo thành công: %PACKAGE_NAME%.zip
echo.
echo 📋 Nội dung package:
echo    - attendance-system.tar (Docker image)
echo    - docker-compose.yml
echo    - start.bat / stop.bat
echo    - DOCKER_README.md
echo    - INSTALL.md
echo    - uploads/ (dữ liệu)
echo    - rawdata/ (dữ liệu mẫu)
echo.
echo 🚀 Để cài đặt trên máy khác:
echo    1. Copy file %PACKAGE_NAME%.zip sang máy đích
echo    2. Giải nén
echo    3. Đảm bảo Docker Desktop đang chạy
echo    4. Chạy lệnh: docker load -i attendance-system.tar
echo    5. Chạy: start.bat
echo.
echo Nhấn phím bất kỳ để đóng cửa sổ này...
pause >nul
