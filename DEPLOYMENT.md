# Hướng dẫn Deploy lên Web

## 1. Deploy lên Heroku (Miễn phí - Khuyên dùng)

### Bước 1: Cài đặt Heroku CLI
- Tải và cài đặt từ: https://devcenter.heroku.com/articles/heroku-cli
- Đăng ký tài khoản Heroku tại: https://signup.heroku.com/

### Bước 2: Đăng nhập và tạo app
```bash
heroku login
heroku create your-attendance-system
```

### Bước 3: Deploy
```bash
git add .
git commit -m "Prepare for deployment"
git push heroku main
```

### Bước 4: Mở app
```bash
heroku open
```

## 2. Deploy lên Railway (Dễ dàng)

### Bước 1: 
- Đăng nhập Railway.app với GitHub
- Connect repository này

### Bước 2:
- Railway sẽ tự động detect Flask app
- Deploy sẽ chạy tự động

## 3. Deploy lên Render (Miễn phí)

### Bước 1:
- Đăng ký tại render.com
- Connect GitHub repository

### Bước 2: Tạo Web Service
- Build Command: `pip install -r requirements.txt`
- Start Command: `gunicorn app:app`

## 4. Deploy lên PythonAnywhere (Miễn phí có giới hạn)

### Bước 1:
- Đăng ký tại pythonanywhere.com
- Upload code lên

### Bước 2: Cấu hình Web App
- Chọn Flask
- Trỏ đến app.py

## 5. Deploy lên VPS (DigitalOcean, Linode, etc.)

### Bước 1: Cài đặt server
```bash
sudo apt update
sudo apt install python3 python3-pip nginx
```

### Bước 2: Clone code
```bash
git clone https://github.com/your-username/Attendance-System.git
cd Attendance-System
pip3 install -r requirements.txt
```

### Bước 3: Cấu hình Nginx
```nginx
server {
    listen 80;
    server_name your-domain.com;
    
    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
```

### Bước 4: Chạy với Gunicorn
```bash
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

## Lưu ý quan trọng:

1. **Bảo mật**: Đổi SECRET_KEY trong production
2. **Database**: Cân nhắc dùng PostgreSQL thay vì file Excel
3. **File uploads**: Cấu hình storage (AWS S3, etc.)
4. **HTTPS**: Bắt buộc cho production
5. **Environment variables**: Sử dụng cho config nhạy cảm

## Khuyến nghị:

- **Người mới**: Dùng Heroku hoặc Railway
- **Miễn phí**: Render, PythonAnywhere  
- **Chuyên nghiệp**: VPS với Docker
- **Doanh nghiệp**: AWS, GCP, Azure
