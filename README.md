<div align="center">

# 📄 VietDoc Assistant - Hệ Thống Chuẩn Hóa Tài Liệu Thông Minh

[![Java](https://img.shields.io/badge/Java-17-orange?style=flat-square&logo=openjdk)](https://www.oracle.com/java/)
[![Spring Boot](https://img.shields.io/badge/Spring%20Boot-3.2.1-brightgreen?style=flat-square&logo=spring-boot)](https://spring.io/projects/spring-boot)
[![Maven](https://img.shields.io/badge/Maven-3.8+-red?style=flat-square&logo=apache-maven)](https://maven.apache.org/)
[![Docker](https://img.shields.io/badge/Docker-Ready-blue?style=flat-square&logo=docker)](https://www.docker.com/)

**Giải pháp tự động hóa định dạng tài liệu Word theo chuẩn Bộ Giáo dục và Nghị định 30**

[🌐 Website](https://formatpro.id.vn) • [📖 Tài liệu](#-tài-liệu-api) • [🚀 Bắt đầu](#-hướng-dẫn-cài-đặt--chạy-local)

</div>

---

## 📋 Giới thiệu

VietDoc Assistant là nền tảng SaaS chuyên nghiệp giúp tự động hóa quá trình định dạng tài liệu Word (`.docx`) theo các quy chuẩn của Bộ Giáo dục và Nghị định 30/2020/NĐ-CP. Hệ thống được thiết kế đặc biệt cho:

- 🎓 **Sinh viên** chuẩn hóa đồ án, khóa luận tốt nghiệp
- 👨‍💼 **Nhân viên văn phòng** định dạng báo cáo, văn bản
- 🏫 **Giảng viên** thẩm định và kiểm tra định dạng tài liệu

## ✨ Tính năng nổi bật

### 🎯 Định dạng thông minh
- **Cấu hình linh hoạt:** Tùy chỉnh Căn lề, Cỡ chữ, Giãn dòng, Thụt lề
- **Font chuẩn hóa:** Tự động chuyển về Times New Roman
- **Dọn dẹp văn bản:** Xóa ký tự rác, khắc phục lỗi ngắt dòng

### 📊 Xử lý bảng biểu & hình ảnh
- **Tự động căn giữa:** Bo khung và đồng bộ font trong bảng
- **Caption thông minh:** In nghiêng tự động chú thích Bảng/Ảnh
- **Đồng bộ kích thước:** Chuẩn hóa font và size trong bảng

### 📄 Đánh số trang chuyên nghiệp
- **Bỏ qua trang bìa:** Tự động chèn số trang từ trang nội dung
- **Định dạng linh hoạt:** Hỗ trợ nhiều kiểu đánh số trang

### 🔒 Bảo mật & Tối ưu
- **Xử lý in-memory:** File được xử lý trong bộ nhớ
- **Tự động dọn dẹp:** Xóa file tạm ngay sau khi hoàn tất
- **An toàn dữ liệu:** Không lưu trữ file người dùng

---

## 🛠 Kiến trúc & Công nghệ

### Backend Architecture
```
┌─────────────────────────────────────────────────────────────┐
│                    Spring Boot 3.2.1                        │
├─────────────────────────────────────────────────────────────┤
│  REST API  │  Security  │  Validation  │  Redis Cache      │
├─────────────────────────────────────────────────────────────┤
│  Apache POI │  Docx4j   │  VnCoreNLP   │  PostgreSQL       │
└─────────────────────────────────────────────────────────────┘
```

### Tech Stack

#### 🎨 Frontend
- **[React](https://reactjs.org/)** + **[Vite](https://vitejs.dev/)** - Tối ưu build và render
- **[Tailwind CSS](https://tailwindcss.com/)** - Giao diện hiện đại, responsive

#### ⚙️ Backend
- **[Java 17](https://www.oracle.com/java/)** - Core language
- **[Spring Boot 3](https://spring.io/projects/spring-boot)** - Main framework
- **[Apache POI 5.2.3](https://poi.apache.org/)** - Xử lý file Word
- **[Docx4j 11.4.9](https://www.docx4java.org/)** - Can thiệp cấu trúc XML
- **[VnCoreNLP](https://github.com/vncorenlp/VnCoreNLP)** - Xử lý ngôn ngữ tiếng Việt
- **[PostgreSQL](https://www.postgresql.org/)** - Database
- **[Redis](https://redis.io/)** - Caching layer

#### 🐳 Deployment
- **[Docker](https://www.docker.com/)** - Containerization
- **[Docker Compose](https://docs.docker.com/compose/)** - Multi-container deployment
- **[Nginx](https://nginx.org/)** - Reverse proxy & SSL termination

---

## 🚀 Hướng dẫn Cài đặt & Chạy Local

### 📋 Yêu cầu hệ thống

| Công cụ | Phiên bản tối thiểu | Bắt buộc |
|---------|-------------------|----------|
| **Java JDK** | 17+ | ✅ |
| **Node.js** | 18+ | ✅ |
| **Maven** | 3.8+ | ✅ |
| **Docker** | 20.10+ | ⚠️ (Tùy chọn) |
| **PostgreSQL** | 13+ | ⚠️ (Tùy chọn) |

### 🔧 Cài đặt Backend

```bash
# 1. Clone repository
git clone https://github.com/your-username/vietdoc-assistant.git
cd vietdoc-assistant

# 2. Build project với Maven
mvn clean package -DskipTests

# 3. Chạy ứng dụng
java -jar target/vietdoc-assistant-0.0.1-SNAPSHOT.jar
```

🎉 **Backend sẽ chạy tại:** `http://localhost:8080`

### 🐳 Triển khai với Docker (Production)

```bash
# Build và khởi chạy tất cả services
docker compose up -d --build

# Kiểm tra trạng thái services
docker compose ps

# Xem log của backend
docker compose logs -f backend
```

### 🗄️ Cấu hình Database

```yaml
# application.yml
spring:
  datasource:
    url: jdbc:postgresql://localhost:5432/vietdoc_db
    username: ${DB_USERNAME:vietdoc}
    password: ${DB_PASSWORD:your_password}
  
  redis:
    host: localhost
    port: 6379
```

---

## 📡 Tài liệu API

### 📤 Upload & Format Document

**Endpoint:** `POST /api/format/upload`

**Content-Type:** `multipart/form-data`

#### Parameters

| Tham số | Kiểu dữ liệu | Bắt buộc | Mặc định | Mô tả |
|---------|-------------|----------|----------|-------|
| `file` | File | ✅ | - | File Word (.docx, Max: 15MB) |
| `marginLeft` | Double | ❌ | 3.5 | Lề trái (cm) |
| `marginRight` | Double | ❌ | 2.0 | Lề phải (cm) |
| `marginTop` | Double | ❌ | 2.0 | Lề trên (cm) |
| `marginBottom` | Double | ❌ | 2.0 | Lề dưới (cm) |
| `lineSpacing` | Double | ❌ | 1.5 | Giãn dòng |
| `fontSizeBody` | Integer | ❌ | 13 | Cỡ chữ nội dung |
| `italicCaption` | Boolean | ❌ | true | In nghiêng caption |

#### Response

**Success (200 OK)**
```http
Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document
Content-Disposition: attachment; filename="formatted_document.docx"

[Binary file data]
```

**Error Responses**
```json
// 400 Bad Request
{
  "error": "INVALID_FILE_FORMAT",
  "message": "File phải có định dạng .docx"
}

// 413 Payload Too Large
{
  "error": "FILE_TOO_LARGE",
  "message": "Kích thước file không được vượt quá 15MB"
}

// 500 Internal Server Error
{
  "error": "PROCESSING_ERROR",
  "message": "Lỗi xử lý file, vui lòng thử lại"
}
```

### 🔍 Health Check

**Endpoint:** `GET /api/health`

```json
{
  "status": "UP",
  "timestamp": "2024-01-15T10:30:00Z",
  "version": "0.0.1-SNAPSHOT"
}
```

---

## 📁 Cấu trúc Project

```
vietdoc-assistant/
├── 📁 src/
│   ├── 📁 main/
│   │   ├── 📁 java/vn/vietdoc/
│   │   │   ├── 📁 controller/     # REST API endpoints
│   │   │   ├── 📁 service/        # Business logic
│   │   │   ├── 📁 model/          # Data models
│   │   │   ├── 📁 config/         # Configuration
│   │   │   └── 📁 util/           # Utility classes
│   │   └── 📁 resources/
│   │       ├── 📄 application.yml # App configuration
│   │       └── 📁 static/         # Static files
│   └── 📁 test/                   # Unit tests
├── 📁 libs/                       # External libraries
│   └── 📄 VnCoreNLP-1.2.jar      # Vietnamese NLP library
├── 📄 pom.xml                     # Maven configuration
├── 📄 Dockerfile                  # Docker build file
├── 📄 docker-compose.yml          # Multi-container setup
└── 📄 README.md                   # This file
```

---

## 🧪 Testing

```bash
# Chạy tất cả tests
mvn test

# Chạy tests với coverage
mvn jacoco:report

# Chạy tests cho class cụ thể
mvn test -Dtest=DocumentFormattingServiceTest
```

---

## 🤝 Đóng góp

Chúng tôi chào đón mọi đóng góp từ cộng đồng!

### 📝 Quy trình đóng góp

1. **Fork** repository
2. Tạo **branch** mới (`git checkout -b feature/amazing-feature`)
3. **Commit** changes (`git commit -m 'Add amazing feature'`)
4. **Push** lên branch (`git push origin feature/amazing-feature`)
5. Tạo **Pull Request**

### 📋 Code style

- Tuân thủ **Google Java Style Guide**
- Sử dụng **Spring Boot conventions**
- Viết **unit tests** cho mọi functions mới
- Thêm **JavaDoc** cho public APIs

---

## 📄 License

Dự án được phân phối dưới **MIT License** - xem file [LICENSE](LICENSE) để biết chi tiết.

---

## 📞 Liên hệ

- **📧 Email:** support@formatpro.id.vn
- **🌐 Website:** [formatpro.id.vn](https://formatpro.id.vn)
- **🐛 Issues:** [GitHub Issues](https://github.com/your-username/vietdoc-assistant/issues)

---

<div align="center">

**⭐ Nếu dự án hữu ích, hãy cho chúng tôi một star!**

Made with ❤️ by VietDoc Team

</div>
