# 🐳 ATTENDANCE SYSTEM - DOCKER DEPLOYMENT

## 📋 Tổng quan
Project Attendance System đã được đóng gói sử dụng Docker để dễ dàng triển khai và chuyển giao giữa các máy tính khác nhau.

## 🚀 Cách sử dụng nhanh

### Máy hiện tại (Development):
```bash
# Khởi chạy nhanh
double-click start.bat

# Hoặc sử dụng lệnh
docker-compose up -d
```

### Chuyển sang máy khác:
```bash
# Tạo package để chuyển giao
double-click create-package.bat

# Copy file .zip sang máy đích và làm theo hướng dẫn trong INSTALL.md
```

## 📁 Cấu trúc Files

### Docker Files:
- `Dockerfile` - Định nghĩa container
- `docker-compose.yml` - Cấu hình services
- `.dockerignore` - Loại trừ files không cần thiết

### Scripts:
- `start.bat` - Khởi chạy nhanh (Windows)
- `stop.bat` - Dừng ứng dụng (Windows)
- `start.sh` - Khởi chạy nhanh (Linux/Mac)
- `create-package.bat` - Tạo package để chuyển giao

### Documentation:
- `DOCKER_README.md` - Hướng dẫn chi tiết Docker
- `DOCKER_SETUP.md` - File này

## 🎯 Ưu điểm của Docker

### 1. **Portable (Di động)**
- Chạy được trên bất kỳ máy nào có Docker
- Không cần cài đặt Python, dependencies
- Môi trường giống hệt nhau trên mọi máy

### 2. **Easy Setup (Dễ cài đặt)**
- Chỉ cần Docker Desktop
- 1 lệnh để khởi chạy toàn bộ ứng dụng
- Tự động xử lý dependencies

### 3. **Consistent (Nhất quán)**
- Môi trường production giống development
- Không có lỗi "works on my machine"
- Version control cho infrastructure

### 4. **Scalable (Có thể mở rộng)**
- Dễ dàng thêm database, Redis, etc.
- Load balancing với multiple containers
- Monitoring và logging tập trung

## 🔧 Cấu hình

### Ports:
- **5000**: Web application (http://localhost:5000)

### Volumes:
- `./uploads` → `/app/uploads`: Dữ liệu upload và xử lý
- `./rawdata` → `/app/rawdata`: Dữ liệu mẫu

### Environment Variables:
```yaml
FLASK_ENV: production
SECRET_KEY: your-production-secret-key
```

## 📊 Monitoring

### Health Check:
```bash
# Kiểm tra trạng thái container
docker-compose ps

# Kiểm tra health status
docker inspect attendance-system-app | grep Health
```

### Logs:
```bash
# Xem logs real-time
docker-compose logs -f

# Xem logs của service cụ thể
docker-compose logs attendance-system
```

### Resource Usage:
```bash
# Xem resource usage
docker stats attendance-system-app
```

## 🛠 Troubleshooting

### Common Issues:

#### 1. Port conflict:
```yaml
# Thay đổi port trong docker-compose.yml
ports:
  - "5001:5000"  # Dùng port 5001 thay vì 5000
```

#### 2. Permission issues:
```bash
# Windows
icacls uploads /grant Users:F /T
icacls rawdata /grant Users:F /T

# Linux/Mac
chmod -R 755 uploads/ rawdata/
```

#### 3. Memory issues:
```yaml
# Thêm memory limit trong docker-compose.yml
deploy:
  resources:
    limits:
      memory: 1G
```

#### 4. Container won't start:
```bash
# Xem logs chi tiết
docker-compose logs

# Rebuild container
docker-compose down
docker-compose up --build -d
```

## 🔄 Updates & Maintenance

### Update application:
```bash
# Pull new code
git pull

# Rebuild containers
docker-compose down
docker-compose up --build -d
```

### Backup data:
```bash
# Backup uploads directory
tar -czf attendance-backup-$(date +%Y%m%d).tar.gz uploads/ rawdata/
```

### Clean up:
```bash
# Remove unused images
docker image prune

# Remove unused volumes
docker volume prune

# Remove unused containers
docker container prune
```

## 🌐 Production Deployment

### Security checklist:
- [ ] Change default SECRET_KEY
- [ ] Use HTTPS (SSL/TLS)
- [ ] Set up firewall rules
- [ ] Regular security updates
- [ ] Backup strategy
- [ ] Monitoring setup

### Recommended additions:
- **Nginx**: Reverse proxy và SSL termination
- **Let's Encrypt**: Free SSL certificates
- **Monitoring**: Prometheus + Grafana
- **Logging**: ELK Stack hoặc centralized logging

## 📞 Support

### Quick commands:
```bash
# Start application
start.bat

# Stop application  
stop.bat

# View logs
docker-compose logs -f

# Restart application
docker-compose restart

# Access container shell
docker exec -it attendance-system-app /bin/bash
```

### Files location:
- **Application**: Inside container `/app/`
- **Data**: Host `./uploads/` and `./rawdata/`
- **Logs**: `docker-compose logs`

---

## 📝 Notes

1. **Data persistence**: Dữ liệu được lưu trong volumes, sẽ không mất khi restart container
2. **Auto-restart**: Container sẽ tự động khởi động lại nếu crash
3. **Health checks**: Tự động kiểm tra tình trạng ứng dụng
4. **Resource limits**: Có thể set giới hạn CPU/Memory nếu cần

Để biết thêm chi tiết, xem `DOCKER_README.md`.
