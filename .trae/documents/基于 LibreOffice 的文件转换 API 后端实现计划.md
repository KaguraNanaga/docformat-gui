# 基于 LibreOffice Docker 的文件转换 API 后端实现计划

## 项目结构
```
backend/
├── app/
│   ├── __init__.py
│   ├── main.py                 # FastAPI 主入口
│   ├── config.py               # 配置管理
│   ├── dependencies.py         # 依赖注入
│   ├── api/
│   │   ├── __init__.py
│   │   └── convert.py         # 转换 API 端点
│   ├── core/
│   │   ├── __init__.py
│   │   ├── converter.py       # LibreOffice 转换核心逻辑
│   │   └── file_handler.py    # 文件上传/下载处理
│   └── schemas.py            # Pydantic 模型
├── requirements.txt
├── Dockerfile
└── docker-compose.yml
```

## 实现步骤

### 1. 后端容器配置
- 使用 FastAPI + Uvicorn 构建后端服务
- 配置与 LibreOffice 容器的网络通信
- 设置文件上传/下载的临时存储目录

### 2. 核心转换模块
- `libreoffice --headless --convert-to <format>` 命令封装
- 支持的输入/输出格式映射
- 异步转换处理
- 错误处理和日志记录

### 3. API 端点设计
```
POST /api/v1/convert/document
- 接收: 文档文件 (docx/doc/wps/pdf)
- 参数: target_format (目标格式)
- 返回: 转换后的文件下载链接

POST /api/v1/convert/spreadsheet
- 接收: 表格文件 (xlsx/xls/csv)
- 参数: target_format (目标格式)
- 返回: 转换后的文件下载链接

GET /api/v1/formats
- 返回: 支持的所有格式列表
```

### 4. Docker Compose 配置更新
- 添加 FastAPI 后端服务
- 配置后端与 LibreOffice 的共享卷
- 设置健康检查和自动重启

### 5. 安全性考虑
- 文件类型验证
- 文件大小限制
- 临时文件清理机制
- 转换超时控制

## 支持的转换格式

### 文档转换
- doc ↔ docx ↔ pdf ↔ wps

### 表格转换  
- xlsx ↔ xls ↔ csv