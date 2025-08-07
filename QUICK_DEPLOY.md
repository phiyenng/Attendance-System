# 🚀 QUICK DEPLOY GUIDE

## Cách nhanh nhất để deploy lên web:

### Option 1: Heroku (Khuyên dùng - Miễn phí)

1. **Cài Heroku CLI**: https://devcenter.heroku.com/articles/heroku-cli
2. **Chạy script tự động**:
   ```bash
   # Windows
   deploy_heroku.bat
   
   # Linux/Mac  
   chmod +x deploy_heroku.sh
   ./deploy_heroku.sh
   ```

### Option 2: Railway (Dễ nhất - 1 click)

1. Vào https://railway.app
2. Login với GitHub
3. Click "Deploy from GitHub repo"
4. Chọn repository này
5. Deploy tự động!

### Option 3: Render (Miễn phí)

1. Vào https://render.com
2. Connect GitHub
3. Create "Web Service"
4. Build: `pip install -r requirements.txt`
5. Start: `gunicorn app:app`

## 🔧 Sau khi deploy:

1. **Test các chức năng chính**:
   - Upload employee list
   - Import sign-in/out data  
   - Export báo cáo

2. **Cấu hình domain** (tùy chọn):
   - Heroku: Settings → Domains
   - Railway: Settings → Custom Domain
   - Render: Settings → Custom Domain

3. **Bảo mật** (quan trọng):
   - Đổi SECRET_KEY trong environment variables
   - Cấu hình HTTPS (tự động trên hầu hết platforms)

## 🆘 Troubleshooting:

- **Build fails**: Kiểm tra requirements.txt
- **App crashes**: Xem logs với `heroku logs --tail`
- **Upload không work**: Kiểm tra file size limits
- **Export lỗi**: Đảm bảo openpyxl install đúng

## 💡 Tips:

- Heroku sleep sau 30 phút không dùng (free tier)
- Railway có 500 giờ/tháng miễn phí
- Render rebuild mỗi khi push code
- Backup dữ liệu định kỳ

**🎉 Chúc bạn deploy thành công!**
