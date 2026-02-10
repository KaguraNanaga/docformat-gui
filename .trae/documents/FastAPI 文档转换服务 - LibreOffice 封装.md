## 项目结构
```
backend/
├── app/
│   ├── __init__.py
│   ├── main.py              # FastAPI 主入口
│   ├── api/
│   │   ├── __init__.py
│   │   └── convert.py      # 转换 API 端点
│   ├── core/
│   │   ├── __init__.py
│   │   ├── config.py        # 配置
│   │   └── converter.py     # LibreOffice 转换核心
│   ├── models/
│   │   ├── __init__.py
│   │   └── schemas.py       # Pydantic 模型
│   └── utils/
│       ├── __init__.py
│       └── cleanup.py        # 临时文件清理工具
├── Dockerfile
├── requirements.txt
└── .env.example
```

## 核心功能模块

### 1. 配置模块 (config.py)
- LibreOffice 容器连接配置
- 临时文件目录配置
- 清理策略配置（定时清理/阈值清理）

### 2. 转换核心 (converter.py)
#### 格式映射表
```python
FORMAT_MAP = {
    # 文档互转
    'doc': {'writer_pdf_Export': 'pdf', 'MS Word 97': 'doc', 'Office Open XML Text': 'docx'},
    'docx': {'writer_pdf_Export': 'pdf', 'MS Word 97': 'doc'},
    'pdf': {'writer8': 'odt', 'MS Word 97': 'doc', 'Office Open XML Text': 'docx'},
    'wps': {'writer_pdf_Export': 'pdf', 'Office Open XML Text': 'docx'},
    
    # 表格互转
    'xls': {'calc_pdf_Export': 'pdf', 'MS Excel 97': 'xls', 'calc8': 'ods', 'Text CSV': 'csv'},
    'xlsx': {'calc_pdf_Export': 'pdf', 'MS Excel 97': 'xls', 'calc8': 'ods', 'Text CSV': 'csv'},
    'csv': {'MS Excel 97': 'xls', 'Office Open XML Spreadsheet': 'xlsx', 'calc8': 'ods'},
}
```

#### 转换方法
1. **直接命令行转换**：`libreoffice --headless --convert-to <format>`
2. **Docker 容器调用**：`docker exec libreoffice ...`
3. **错误重试机制**：最多重试 3 次

### 3. 临时文件清理 (cleanup.py)
#### 清理策略
- **策略1：任务级清理**：每个请求完成后立即清理
- **策略2：定时清理**：每小时清理过期文件（>1小时）
- **策略3：阈值清理**：当目录大小超过 500MB 时清理最旧文件
- **策略4：启动清理**：服务启动时清理遗留文件

#### 清理实现
```python
async def cleanup_task(temp_dir: Path, max_age_hours: int = 1, max_size_mb: int = 500)
async def cleanup_on_request(temp_dir: Path, task_id: str)
async def cleanup_on_startup(temp_dir: Path)
```

### 4. API 端点 (convert.py)
```
POST /api/v1/convert/文档          # 文档格式互转
POST /api/v1/convert/spreadsheet   # 表格格式互转
GET  /api/v1/health                # 健康检查
GET  /api/v1/formats               # 支持的格式列表
```

## Docker Compose 配置更新

```yaml
services:
  libreoffice:
    image: linuxserver/libreoffice:25.8.1
    container_name: libreoffice
    volumes:
      - ./data/temp:/temp           # 临时文件目录
      - ./data/output:/output
    ports:
      - "3000:3000"
      - "8100:8100"                # UNO API 端口
    restart: unless-stopped

  backend:
    build: ./backend
    container_name: doc-converter-api
    volumes:
      - ./backend:/app
      - ./data/temp:/temp
    ports:
      - "8000:8000"
    depends_on:
      - libreoffice
    environment:
      - LIBREOFFICE_HOST=libreoffice
      - TEMP_DIR=/temp
      - MAX_TEMP_SIZE_MB=500
      - CLEANUP_INTERVAL_HOURS=1
```

## 临时文件目录结构
```
/temp/
├── uploads/                        # 上传的原始文件
│   └── {task_id}/
│       └── {filename}
├── converted/                      # 转换后的文件
│   └── {task_id}/
│       └── {filename}
└── .cleanup.log                    # 清理日志
```

## 依赖文件
```
fastapi==0.109.0
uvicorn[standard]==0.27.0
python-multipart==0.0.6
aiofiles==23.2.1
python-dotenv==1.0.0
```

## 实施步骤
1. 创建项目目录结构
2. 编写配置和转换核心模块
3. 实现临时文件清理工具
4. 创建 FastAPI 端点
5. 编写 Dockerfile 和 docker-compose.yml
6. 测试转换功能
7. 测试清理机制