# Dockerize Attendance System

## Prerequisites
- Docker Desktop đã được cài đặt và đang chạy
- Git (tùy chọn, để clone project)

## Hướng dẫn sử dụng Docker

### Cách 1: Sử dụng Docker Compose (Khuyến nghị)

#### Bước 1: Chuẩn bị
```bash
# Clone project (nếu chưa có)
git clone <repository-url>
cd Attendance-System

# Hoặc copy toàn bộ folder project vào máy đích
```

#### Bước 2: Khởi chạy ứng dụng
```bash
# Build và chạy container
docker-compose up -d

# Kiểm tra status
docker-compose ps

# Xem logs
docker-compose logs -f
```

#### Bước 3: Truy cập ứng dụng
- Mở trình duyệt và truy cập: http://localhost:5000
- Ứng dụng sẽ sẵn sàng sử dụng

#### Bước 4: Dừng ứng dụng
```bash
# Dừng containers
docker-compose down

# Dừng và xóa volumes (nếu cần reset hoàn toàn)
docker-compose down -v
```

### Cách 2: Sử dụng Docker commands trực tiếp

#### Build image
```bash
docker build -t attendance-system .
```

#### Chạy container
```bash
docker run -d \
  --name attendance-app \
  -p 5000:5000 \
  -v ${PWD}/uploads:/app/uploads \
  -v ${PWD}/rawdata:/app/rawdata \
  -e SECRET_KEY=your-secret-key \
  attendance-system
```

#### Quản lý container
```bash
# Xem logs
docker logs -f attendance-app

# Dừng container
docker stop attendance-app

# Xóa container
docker rm attendance-app

# Vào trong container (để debug)
docker exec -it attendance-app /bin/bash
```

## Cấu trúc dữ liệu

### Volumes được mount:
- `./uploads` → `/app/uploads`: Lưu trữ file upload và dữ liệu xử lý
- `./rawdata` → `/app/rawdata`: Dữ liệu mẫu và backup

### Files quan trọng:
- `uploads/temp_signinout.xlsx`: Dữ liệu chấm công
- `uploads/temp_apply.xlsx`: Dữ liệu đơn xin nghỉ
- `uploads/temp_otlieu.xlsx`: Dữ liệu OT/Lieu
- `uploads/employee_list.csv`: Danh sách nhân viên
- `uploads/rules.xlsx`: Quy tắc ngày lễ và làm việc

## Troubleshooting

### Lỗi thường gặp:

1. **Port 5000 đã được sử dụng:**
   ```bash
   # Thay đổi port trong docker-compose.yml
   ports:
     - "5001:5000"  # Sử dụng port 5001 thay vì 5000
   ```

2. **Container không start được:**
   ```bash
   # Kiểm tra logs
   docker-compose logs
   
   # Rebuild container
   docker-compose down
   docker-compose up --build
   ```

3. **Mất dữ liệu khi restart:**
   - Đảm bảo volumes được mount đúng cách
   - Dữ liệu sẽ được lưu trong thư mục `uploads/` và `rawdata/` trên host

4. **Lỗi permission:**
   ```bash
   # Cấp quyền cho thư mục uploads
   chmod 755 uploads/
   chmod 755 rawdata/
   ```

## Chuyển sang máy khác

### Cách 1: Copy toàn bộ project folder
1. Copy toàn bộ thư mục project
2. Đảm bảo Docker Desktop đã cài đặt trên máy đích
3. Chạy `docker-compose up -d`

### Cách 2: Sử dụng Docker image
1. Export image từ máy nguồn:
   ```bash
   docker save attendance-system > attendance-system.tar
   ```

2. Copy file .tar sang máy đích

3. Import image trên máy đích:
   ```bash
   docker load < attendance-system.tar
   ```

4. Chạy container với image đã import

### Cách 3: Sử dụng Docker Registry (Advanced)
1. Push image lên Docker Hub hoặc private registry
2. Pull image trên máy đích
3. Chạy container

## Backup và Restore

### Backup dữ liệu:
```bash
# Backup toàn bộ thư mục uploads và rawdata
tar -czf attendance-backup-$(date +%Y%m%d).tar.gz uploads/ rawdata/
```

### Restore dữ liệu:
```bash
# Extract backup
tar -xzf attendance-backup-YYYYMMDD.tar.gz
```

## Environment Variables

Có thể customize các biến môi trường trong `docker-compose.yml`:

```yaml
environment:
  - FLASK_ENV=production
  - SECRET_KEY=your-production-secret-key-here
  - MAX_CONTENT_LENGTH=16777216  # 16MB
```

## Security Notes

1. Thay đổi `SECRET_KEY` trong production
2. Không expose database ports ra ngoài nếu không cần thiết
3. Sử dụng HTTPS trong production environment
4. Regularly update Docker images và dependencies

## Support

Nếu gặp vấn đề, hãy kiểm tra:
1. Docker Desktop đang chạy
2. Ports không bị conflict
3. Đủ dung lượng disk space
4. Logs của container để xem chi tiết lỗi
