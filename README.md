# RPA Control Center

## 🛠 Công nghệ sử dụng

| Thành phần | Công nghệ |
| :--- | :--- |
| **Core** | [Python 3.10+](https://www.python.org/) & [FastAPI](https://fastapi.tiangolo.com/)
| **Task Queue** | [Celery](https://docs.celeryq.dev/) & [Redis](https://redis.io/)
| **Database** | [MySQL](https://www.mysql.com/)
| **Migration** | [Alembic](https://alembic.sqlalchemy.org/)
| **Real-time** | [Socket.IO](https://socket.io/)
| **Package Manager** | [uv](https://github.com/astral-sh/uv)
| **Plugin** | [C++](https://cplusplus.com)

<a title="Python" href="https://www.python.org/">
  <img
    src="https://img.shields.io/badge/python-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54"
  />
</a>
<a title="FastAPI" href="https://fastapi.tiangolo.com/">
  <img
    src="https://img.shields.io/badge/FastAPI-005571?style=for-the-badge&logo=fastapi"
  />
</a>

## 🚀 Cài đặt & Chạy dự án

### 1. Yêu cầu tiên quyết

Đảm bảo máy tính của bạn đã cài đặt:

*   [Python 3.10+](https://www.python.org/)
*   [Docker](https://www.docker.com/) & Docker Compose
*   [uv](https://github.com/astral-sh/uv)

### 2. Thiết lập môi trường

**Bước 1: Clone dự án**

```bash
git clone <repository_url>
cd TaskDistribution
```

**Bước 2: Cấu hình biến môi trường**

Copy file cấu hình mẫu và cập nhật thông tin kết nối (Database, Redis, v.v.):

```bash
cp .env.example .env
```

**Bước 3: Cài đặt thư viện**

Sử dụng `uv` để cài đặt các dependencies nhanh chóng:

```bash
uv sync
```

### 3. Khởi chạy Database & Services

Sử dụng Docker để khởi chạy Redis và MySQL (nếu chưa có sẵn):

```bash
docker-compose up -d
```

Chạy migration để khởi tạo cấu trúc database:

```bash
alembic upgrade head
```

### 4. Chạy ứng dụng

Khởi chạy API Server:

```bash
uv run uvicorn main:app --reload
```

Khởi chạy Celery Worker (trên terminal khác):

```bash
uv run celery -A worker.celery_app worker --loglevel=info
```

## 📚 Tài liệu API

Sau khi server khởi chạy thành công, bạn có thể truy cập:

*   **Documentaion:** `https://nhathuyd4hp.github.io/RPAControlCenter/`
