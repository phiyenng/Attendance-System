# 🐳 CÀI ĐẶT DOCKER DESKTOP CHO ATTENDANCE SYSTEM

## ✅ Trạng thái hiện tại
Docker đã được cài đặt thành công trên máy của bạn (Version 28.3.0).

## 🔧 Cài đặt tối ưu Docker Desktop

### 1. **Cài đặt Resources (Tài nguyên)**

Mở Docker Desktop → Settings → Resources:

#### **Memory (RAM):**
- **Minimum**: 2GB 
- **Recommended**: 4GB
- **Optimal**: 6-8GB (nếu máy có đủ RAM)

#### **CPU:**
- **Minimum**: 2 cores
- **Recommended**: 4 cores

#### **Disk Space:**
- **Minimum**: 10GB free space
- **Recommended**: 20GB+ 

### 2. **File Sharing (Chia sẻ file)**

Đảm bảo drive C: được share:
```
Settings → Resources → File Sharing
☑️ C: (hoặc drive chứa project)
```

### 3. **WSL 2 Integration (Nếu dùng Windows)**

Nếu bạn dùng Windows 10/11:
```
Settings → General
☑️ Use the WSL 2 based engine
☑️ Use Docker Compose V2
```

### 4. **Network Settings**

```
Settings → Resources → Network
Port range: 5000-5010 (để chạy nhiều ứng dụng)
```

## 🚀 Kiểm tra cài đặt

### Test Docker hoạt động:
```bash
# Test container đơn giản
docker run hello-world

# Kiểm tra Docker Compose
docker-compose --version
```

### Kiểm tra tài nguyên:
```bash
# Xem thông tin Docker
docker info

# Xem containers đang chạy
docker ps
```

## ⚙️ Cài đặt không bắt buộc nhưng hữu ích

### 1. **Docker Desktop Extensions**
- **Logs Explorer**: Xem logs dễ dàng hơn
- **Resource Usage**: Monitor tài nguyên
- **Volume Browser**: Quản lý volumes

### 2. **Startup Settings**
```
Settings → General
☑️ Start Docker Desktop when you log in
☑️ Open Docker Dashboard at startup (tùy chọn)
```

### 3. **Update Settings**
```
Settings → Software Updates
☑️ Automatically check for updates
```

## 🛠 Troubleshooting Settings

### Nếu gặp lỗi memory:
```yaml
# Trong docker-compose.yml, thêm:
services:
  attendance-system:
    deploy:
      resources:
        limits:
          memory: 1G
        reservations:
          memory: 512M
```

### Nếu gặp lỗi port conflict:
```yaml
# Thay đổi port mapping:
ports:
  - "5001:5000"  # Dùng port 5001 thay vì 5000
```

### Nếu gặp lỗi file permission:
```bash
# Windows PowerShell (Run as Administrator)
icacls "C:\Users\PC\Documents\GitHub\Attendance-System" /grant Users:F /T
```

## 📋 Checklist trước khi chạy

- [ ] Docker Desktop đang chạy (icon màu xanh)
- [ ] Drive C: được share trong File Sharing
- [ ] Memory allocation ≥ 2GB
- [ ] Port 5000 không bị chiếm bởi app khác
- [ ] Windows Defender/Antivirus không block Docker

## 🎯 Test ứng dụng

Sau khi setup xong:

```bash
# Chạy test nhanh
cd "C:\Users\PC\Documents\GitHub\Attendance-System"
docker-compose up -d

# Kiểm tra logs
docker-compose logs -f

# Test truy cập
# Mở browser: http://localhost:5000
```

## 🔍 Monitoring Docker

### Docker Desktop Dashboard:
- **Containers**: Xem containers đang chạy
- **Images**: Quản lý Docker images  
- **Volumes**: Xem data volumes
- **Dev Environments**: Development tools

### Command line monitoring:
```bash
# Xem resource usage
docker stats

# Xem disk usage
docker system df

# Cleanup (khi cần)
docker system prune
```

## ⚠️ Lưu ý quan trọng

### 1. **Windows Defender**
Có thể cần thêm Docker vào whitelist:
```
Windows Security → Virus & threat protection → 
Add or remove exclusions → Add folder:
C:\Program Files\Docker
```

### 2. **Hyper-V** (Windows Pro/Enterprise)
Docker có thể conflict với VirtualBox. Nếu cần dùng cả hai:
```bash
# Tắt Hyper-V tạm thời
bcdedit /set hypervisorlaunchtype off

# Bật lại Hyper-V
bcdedit /set hypervisorlaunchtype auto
```

### 3. **WSL 2** (Windows 10/11)
Đảm bảo WSL 2 được cài đặt và update:
```bash
# Trong PowerShell (Admin)
wsl --update
wsl --set-default-version 2
```

## 📞 Nếu gặp vấn đề

### Khởi động lại Docker:
1. Right-click Docker icon → Restart Docker Desktop
2. Hoặc: Task Manager → End Docker tasks → Start Docker Desktop

### Reset Docker về default:
```
Docker Desktop → Settings → Reset and cleanup → 
Reset to factory defaults
```

### Logs troubleshooting:
```bash
# Xem logs ứng dụng
docker-compose logs attendance-system

# Xem logs Docker Desktop
C:\Users\%USERNAME%\AppData\Roaming\Docker\log.txt
```

---

## ✅ Tóm tắt

**Cài đặt tối thiểu cần thiết:**
1. ✅ Docker Desktop đã cài đặt (Version 28.3.0)
2. ✅ Đảm bảo Docker đang chạy
3. ✅ File sharing cho drive C:
4. ⚙️ Memory allocation ≥ 2GB

**Không cần cài đặt thêm gì khác!** 

Bạn có thể chạy ngay lệnh `start.bat` hoặc `docker-compose up -d` để khởi chạy ứng dụng.
