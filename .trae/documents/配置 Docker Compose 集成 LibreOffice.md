## 计划：配置 Docker Compose 集成 LibreOffice

### 目标
创建 `docker-compose.yml` 文件，使用 `linuxserver/libreoffice:25.8.1` 镜像部署 LibreOffice 服务。

### 实施步骤

1. **创建 docker-compose.yml 文件**
   - 使用 `linuxserver/libreoffice:25.8.1` 镜像
   - 配置服务名称为 `libreoffice`
   - 设置端口映射（默认使用 3000:3000 用于 Web UI 访问）
   - 配置必要的环境变量：
     - `PUID` / `PGID`：用户权限设置
     - `TZ`：时区设置
   - 配置数据卷持久化：
     - `/config`：存储 LibreOffice 配置
     - `/output`：存储输出文件
   - 设置重启策略为 `unless-stopped`
   - 添加容器名称和健康检查

### Docker Compose 配置结构
```yaml
services:
  libreoffice:
    image: linuxserver/libreoffice:25.8.1
    container_name: libreoffice
    environment:
      - PUID=1000
      - PGID=1000
      - TZ=Asia/Shanghai
    volumes:
      - ./data/config:/config
      - ./data/output:/output
    ports:
      - "3000:3000"
    restart: unless-stopped
```

### 说明
- Web UI 访问地址：`http://localhost:3000`
- 配置文件将持久化到 `./data/config` 目录
- 输出文件将保存到 `./data/output` 目录
- 此配置为后续 Web 服务集成 LibreOffice API 做准备