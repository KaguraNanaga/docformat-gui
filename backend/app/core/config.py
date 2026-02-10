"""配置模块"""
from pathlib import Path
from typing import Optional
from pydantic_settings import BaseSettings
import os


class Settings(BaseSettings):
    """应用配置类"""

    # LibreOffice 配置
    libreoffice_host: str = "libreoffice"
    libreoffice_container_name: str = "libreoffice"
    libreoffice_port: int = 3000
    uno_port: int = 8100

    # 临时文件配置
    temp_dir: str = "/temp"
    uploads_subdir: str = "uploads"
    converted_subdir: str = "converted"
    cleanup_log_file: str = ".cleanup.log"

    # 清理策略配置
    max_temp_size_mb: int = 500
    max_file_age_hours: int = 1
    cleanup_interval_hours: int = 1
    enable_startup_cleanup: bool = True
    enable_on_request_cleanup: bool = True

    # API 配置
    api_prefix: str = "/api/v1"
    max_upload_size_mb: int = 50
    allowed_origins: list[str] = ["*"]

    # 转换配置
    conversion_timeout_seconds: int = 120
    max_retry_attempts: int = 3

    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"
        case_sensitive = False

    @property
    def uploads_dir(self) -> Path:
        """上传文件目录"""
        return Path(self.temp_dir) / self.uploads_subdir

    @property
    def converted_dir(self) -> Path:
        """转换后文件目录"""
        return Path(self.temp_dir) / self.converted_subdir

    @property
    def cleanup_log_path(self) -> Path:
        """清理日志路径"""
        return Path(self.temp_dir) / self.cleanup_log_file

    @property
    def max_temp_size_bytes(self) -> int:
        """临时目录最大大小（字节）"""
        return self.max_temp_size_mb * 1024 * 1024

    @property
    def max_file_age_seconds(self) -> int:
        """文件最大存活时间（秒）"""
        return self.max_file_age_hours * 3600


# 全局配置实例
settings = Settings()


def get_settings() -> Settings:
    """获取配置实例"""
    return settings
