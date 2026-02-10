# 文档格式转换 API 使用文档

## 简介

本文档介绍如何使用文档格式转换服务进行文档和表格格式互转。该服务基于 LibreOffice Docker 容器，支持多种文档格式的相互转换。

## 基本信息

- **API 基础地址**: `http://localhost:8000`
- **API 版本**: `/api/v1`
- **认证方式**: 无需认证
- **响应格式**: JSON（文件下载除外）

## 支持的格式

### 文档格式
| 格式 | 说明 | MIME 类型 |
|------|------|-----------|
| doc | Microsoft Word 97-2003 | `application/msword` |
| docx | Microsoft Word 2007+ | `application/vnd.openxmlformats-officedocument.wordprocessingml.document` |
| pdf | 便携式文档格式 | `application/pdf` |
| wps | WPS 文字文档 | - |
| odt | OpenDocument 文本 | `application/vnd.oasis.opendocument.text` |
| rtf | 富文本格式 | `application/rtf` |
| txt | 纯文本 | `text/plain` |
| html | HTML 网页 | `text/html` |

### 表格格式
| 格式 | 说明 | MIME 类型 |
|------|------|-----------|
| xls | Microsoft Excel 97-2003 | `application/vnd.ms-excel` |
| xlsx | Microsoft Excel 2007+ | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` |
| csv | 逗号分隔值 | `text/csv` |
| ods | OpenDocument 电子表格 | `application/vnd.oasis.opendocument.spreadsheet` |

## API 端点

### 1. 健康检查

检查服务是否正常运行。

**请求**
```
GET /api/v1/convert/health
```

**响应示例**
```json
{
  "status": "healthy",
  "timestamp": "2026-02-10T08:24:08.307789",
  "libreoffice_available": true,
  "temp_dir_size": "39.83 KB"
}
```

**字段说明**
- `status`: 服务状态 (`healthy` | `degraded`)
- `timestamp`: 检查时间
- `libreoffice_available`: LibreOffice 容器是否可用
- `temp_dir_size`: 临时目录占用空间

---

### 2. 获取支持的格式

获取所有支持的格式及其转换路径。

**请求**
```
GET /api/v1/convert/formats
```

**响应示例**
```json
{
  "document_formats": ["doc", "docx", "pdf", "wps", "odt", "rtf", "txt", "html"],
  "spreadsheet_formats": ["xls", "xlsx", "csv", "ods"],
  "all_conversions": [
    {"source": "doc", "target": "docx", "supported": true},
    {"source": "doc", "target": "pdf", "supported": true},
    {"source": "wps", "target": "docx", "supported": true},
    ...
  ]
}
```

**字段说明**
- `document_formats`: 支持的文档格式列表
- `spreadsheet_formats`: 支持的表格格式列表
- `all_conversions`: 所有支持的转换路径

---

### 3. 文档格式转换

转换文档文件格式。

**请求**
```
POST /api/v1/convert/document
Content-Type: multipart/form-data

file: <文件>
target_format: <目标格式>
cleanup_after: <是否清理临时文件>
```

**参数说明**
| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| file | File | 是 | 上传的文档文件 |
| target_format | String | 是 | 目标格式（docx, doc, pdf, wps, odt, rtf, txt, html） |
| cleanup_after | Boolean | 否 | 是否在转换后清理临时文件，默认 `true` |

**响应**
- 成功：返回转换后的文件
- 失败：返回错误信息

**状态码**
- `200`: 转换成功
- `400`: 请求参数错误
- `500`: 服务器内部错误

---

### 4. 表格格式转换

转换表格文件格式。

**请求**
```
POST /api/v1/convert/spreadsheet
Content-Type: multipart/form-data

file: <文件>
target_format: <目标格式>
cleanup_after: <是否清理临时文件>
```

**参数说明**
| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| file | File | 是 | 上传的表格文件 |
| target_format | String | 是 | 目标格式（xlsx, xls, csv, ods） |
| cleanup_after | Boolean | 否 | 是否在转换后清理临时文件，默认 `true` |

**响应**
- 成功：返回转换后的文件
- 失败：返回错误信息

---

### 5. 清理临时文件

手动触发临时文件清理。

**请求**
```
POST /api/v1/convert/cleanup
```

**响应示例**
```json
{
  "cleaned_by_age": 0,
  "cleaned_by_size": 0,
  "current_size": "39.83 KB",
  "timestamp": "2026-02-10T08:24:08.307789"
}
```

**字段说明**
- `cleaned_by_age`: 按年龄清理的文件数量
- `cleaned_by_size`: 按大小清理的文件数量
- `current_size`: 当前临时目录大小
- `timestamp`: 清理时间

---

## 使用示例

### cURL 示例

#### 1. 将 WPS 文档转换为 DOCX

```bash
curl -X POST http://localhost:8000/api/v1/convert/document \
  -F "file=@document.wps" \
  -F "target_format=docx" \
  -o output.docx
```

#### 2. 将 DOC 文档转换为 PDF

```bash
curl -X POST http://localhost:8000/api/v1/convert/document \
  -F "file=@report.doc" \
  -F "target_format=pdf" \
  -o report.pdf
```

#### 3. 将 XLSX 表格转换为 CSV

```bash
curl -X POST http://localhost:8000/api/v1/convert/spreadsheet \
  -F "file=@data.xlsx" \
  -F "target_format=csv" \
  -o data.csv
```

#### 4. 不清理临时文件

```bash
curl -X POST http://localhost:8000/api/v1/convert/document \
  -F "file=@document.docx" \
  -F "target_format=pdf" \
  -F "cleanup_after=false" \
  -o document.pdf
```

### Python 示例

#### 基本用法

```python
import requests

API_URL = "http://localhost:8000"

def convert_document(file_path, target_format):
    url = f"{API_URL}/api/v1/convert/document"

    with open(file_path, "rb") as f:
        files = {"file": (file_path, f)}
        data = {"target_format": target_format}

        response = requests.post(url, files=files, data=data)

    if response.status_code == 200:
        output_path = file_path.rsplit(".", 1)[0] + f".{target_format}"
        with open(output_path, "wb") as f:
            f.write(response.content)
        print(f"转换成功: {output_path}")
        return True
    else:
        print(f"转换失败: {response.status_code}")
        print(response.text)
        return False

# 使用示例
convert_document("report.wps", "docx")
```

#### 高级用法（带错误处理和重试）

```python
import requests
import time

API_URL = "http://localhost:8000"

def convert_with_retry(file_path, target_format, max_retries=3):
    url = f"{API_URL}/api/v1/convert/document"

    for attempt in range(max_retries):
        try:
            with open(file_path, "rb") as f:
                files = {"file": (file_path, f)}
                data = {"target_format": target_format, "cleanup_after": "true"}

                response = requests.post(
                    url,
                    files=files,
                    data=data,
                    timeout=60
                )

            if response.status_code == 200:
                output_path = file_path.rsplit(".", 1)[0] + f".{target_format}"
                with open(output_path, "wb") as f:
                    f.write(response.content)
                print(f"转换成功: {output_path}")
                return True
            else:
                print(f"转换失败 (尝试 {attempt + 1}/{max_retries}): {response.text}")
                if attempt < max_retries - 1:
                    time.sleep(2)
        except requests.exceptions.RequestException as e:
            print(f"请求异常 (尝试 {attempt + 1}/{max_retries}): {e}")
            if attempt < max_retries - 1:
                time.sleep(2)

    return False
```

### JavaScript/Node.js 示例

```javascript
const FormData = require('form-data');
const fs = require('fs');
const axios = require('axios');

const API_URL = 'http://localhost:8000';

async function convertDocument(filePath, targetFormat) {
    const url = `${API_URL}/api/v1/convert/document`;

    const form = new FormData();
    form.append('file', fs.createReadStream(filePath));
    form.append('target_format', targetFormat);

    try {
        const response = await axios.post(url, form, {
            responseType: 'arraybuffer',
            headers: form.getHeaders()
        });

        const outputPath = filePath.replace(/\.[^.]+$/, `.${targetFormat}`);
        fs.writeFileSync(outputPath, response.data);

        console.log(`转换成功: ${outputPath}`);
        return true;
    } catch (error) {
        console.error('转换失败:', error.response?.data || error.message);
        return false;
    }
}

// 使用示例
convertDocument('document.wps', 'docx');
```

### PowerShell 示例

```powershell
$API_URL = "http://localhost:8000"

function Convert-Document {
    param(
        [string]$FilePath,
        [string]$TargetFormat
    )

    $uri = "$API_URL/api/v1/convert/document"
    $outputPath = [System.IO.Path]::ChangeExtension($FilePath, $TargetFormat)

    try {
        $form = @{
            file = Get-Item -Path $FilePath
            target_format = $TargetFormat
        }

        Invoke-RestMethod -Uri $uri -Method POST -Form $form -OutFile $outputPath
        Write-Host "转换成功: $outputPath"
    } catch {
        Write-Host "转换失败: $_"
    }
}

# 使用示例
Convert-Document -FilePath "document.wps" -TargetFormat "docx"
```

---

## 常见错误处理

### 错误代码

| 状态码 | 说明 | 解决方案 |
|--------|------|----------|
| 400 | 不支持的格式或转换 | 检查文件扩展名和目标格式是否正确 |
| 500 | 服务器内部错误 | 查看服务器日志，检查 LibreOffice 容器状态 |

### 错误响应示例

```json
{
  "error": "Internal Server Error",
  "message": "服务器内部错误",
  "detail": "LibreOffice 转换失败: Error: source file could not be loaded"
}
```

---

## 性能建议

1. **文件大小限制**: 建议单个文件不超过 100MB
2. **并发限制**: 建议同时处理的文件数量不超过 10 个
3. **超时设置**: 建议设置 60 秒超时
4. **临时文件清理**: 建议启用自动清理（`cleanup_after=true`）

---

## 故障排查

### 服务无法访问

```bash
# 检查容器状态
docker ps

# 检查服务健康
curl http://localhost:8000/api/v1/convert/health
```

### 转换失败

```bash
# 查看后端日志
docker logs doc-converter-api --tail 50

# 查看 LibreOffice 日志
docker logs libreoffice --tail 50
```

### 手动清理临时文件

```bash
curl -X POST http://localhost:8000/api/v1/convert/cleanup
```

---

## 配置说明

环境变量配置 (`.env` 文件):

| 变量 | 默认值 | 说明 |
|------|--------|------|
| LIBREOFFICE_HOST | libreoffice | LibreOffice 容器主机名 |
| LIBREOFFICE_CONTAINER_NAME | libreoffice | LibreOffice 容器名称 |
| TEMP_DIR | /temp | 临时文件目录 |
| MAX_TEMP_SIZE_MB | 1024 | 最大临时文件大小（MB） |
| TEMP_FILE_AGE_HOURS | 24 | 临时文件保留时间（小时） |

---

## 技术架构

```
┌─────────────┐         ┌─────────────┐         ┌──────────────┐
│  客户端     │  HTTP   │  FastAPI    │  Docker   │  LibreOffice │
│  (Browser)  │ <-----> │  Backend    │ <----->  │  Container   │
└─────────────┘         └─────────────┘         └──────────────┘
                              │
                              ▼
                        ┌─────────────┐
                        │  Temp Files │
                        │  (/temp)    │
                        └─────────────┘
```

---

## 版本历史

- **v1.0.0** (2026-02-10)
  - 初始版本
  - 支持文档和表格格式互转
  - 支持临时文件自动清理
  - 支持中文文件名

---

## 支持

如有问题，请提交 Issue 或联系项目维护者。
