# BƯỚC 1: MƯỢN MÁY KHOAN MAVEN ĐỂ DỊCH CODE (Build)
FROM maven:3.9.5-eclipse-temurin-17 AS builder
WORKDIR /app

# Copy file cấu hình maven và mã nguồn vào
COPY pom.xml .
COPY src ./src

# ========================================================
# [DÒNG THÊM MỚI] Copy thư mục chứa file .jar local vào Docker
# ========================================================
COPY libs ./libs

# Chạy lệnh build ra file .jar (Bỏ qua test cho nhanh)
RUN mvn clean package -DskipTests

# BƯỚC 2: GÓI FILE .JAR VÀO VIÊN NHỘNG SIÊU NHẸ ĐỂ CHẠY (Runtime)
FROM eclipse-temurin:17-jre-alpine
WORKDIR /app

# Copy file .jar thành phẩm từ BƯỚC 1 sang BƯỚC 2
COPY --from=builder /app/target/*.jar app.jar

# Khai báo cổng ứng dụng (Spring Boot mặc định 8080)
EXPOSE 8080

# Lệnh khởi động ứng dụng
ENTRYPOINT ["java", "-jar", "app.jar"]