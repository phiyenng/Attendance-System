#!/bin/bash

# Script khởi chạy nhanh Attendance System với Docker

echo "🚀 Đang khởi chạy Attendance System..."

# Kiểm tra Docker đã chạy chưa
if ! docker info > /dev/null 2>&1; then
    echo "❌ Docker chưa chạy. Vui lòng khởi động Docker Desktop trước."
    exit 1
fi

echo "✅ Docker đã sẵn sàng"

# Kiểm tra file docker-compose.yml
if [ ! -f "docker-compose.yml" ]; then
    echo "❌ Không tìm thấy file docker-compose.yml"
    exit 1
fi

# Dừng container cũ nếu có
echo "🔄 Dừng container cũ (nếu có)..."
docker-compose down > /dev/null 2>&1

# Build và khởi chạy
echo "🔨 Đang build và khởi chạy container..."
docker-compose up -d --build

# Kiểm tra trạng thái
if [ $? -eq 0 ]; then
    echo "✅ Attendance System đã khởi chạy thành công!"
    echo ""
    echo "📝 Thông tin truy cập:"
    echo "   🌐 URL: http://localhost:5000"
    echo "   📁 Dữ liệu: ./uploads/"
    echo ""
    echo "📋 Các lệnh hữu ích:"
    echo "   🔍 Xem logs:     docker-compose logs -f"
    echo "   ⏹️  Dừng app:     docker-compose down"
    echo "   🔄 Restart:      docker-compose restart"
    echo ""
    echo "⏳ Đang đợi ứng dụng khởi động hoàn tất..."
    
    # Đợi ứng dụng sẵn sàng
    max_attempts=30
    attempt=0
    while [ $attempt -lt $max_attempts ]; do
        if curl -s http://localhost:5000 > /dev/null 2>&1; then
            echo "🎉 Ứng dụng đã sẵn sàng! Truy cập http://localhost:5000"
            break
        fi
        attempt=$((attempt + 1))
        echo -n "."
        sleep 2
    done
    
    if [ $attempt -eq $max_attempts ]; then
        echo ""
        echo "⚠️  Ứng dụng khởi động chậm hơn dự kiến. Vui lòng kiểm tra logs:"
        echo "   docker-compose logs"
    fi
else
    echo "❌ Có lỗi khi khởi chạy. Kiểm tra logs:"
    docker-compose logs
    exit 1
fi
