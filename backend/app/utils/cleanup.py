"""临时文件清理工具模块"""
import asyncio
import logging
import shutil
import time
from pathlib import Path
from typing import Optional
from datetime import datetime

from app.core.config import settings

logger = logging.getLogger(__name__)


class TempFileCleaner:
    """临时文件清理器"""

    def __init__(self, temp_dir: Optional[Path] = None):
        """初始化清理器"""
        self.temp_dir = temp_dir or Path(settings.temp_dir)
        self.uploads_dir = self.temp_dir / settings.uploads_subdir
        self.converted_dir = self.temp_dir / settings.converted_subdir
        self.log_path = self.temp_dir / settings.cleanup_log_file

        # 确保目录存在
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        self.uploads_dir.mkdir(parents=True, exist_ok=True)
        self.converted_dir.mkdir(parents=True, exist_ok=True)

    def _log_cleanup(self, message: str):
        """记录清理日志"""
        try:
            with open(self.log_path, "a", encoding="utf-8") as f:
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                f.write(f"[{timestamp}] {message}\n")
        except Exception as e:
            logger.error(f"写入清理日志失败: {e}")

    def _get_dir_size(self, directory: Path) -> int:
        """获取目录总大小（字节）"""
        total_size = 0
        try:
            for item in directory.rglob("*"):
                if item.is_file():
                    total_size += item.stat().st_size
        except Exception as e:
            logger.error(f"计算目录大小失败 {directory}: {e}")
        return total_size

    def _format_size(self, size_bytes: int) -> str:
        """格式化文件大小显示"""
        for unit in ["B", "KB", "MB", "GB"]:
            if size_bytes < 1024.0:
                return f"{size_bytes:.2f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.2f} TB"

    async def cleanup_by_age(self, max_age_seconds: Optional[int] = None) -> int:
        """根据文件年龄清理过期文件

        Args:
            max_age_seconds: 最大存活时间（秒），默认使用配置值

        Returns:
            int: 清理的文件数量
        """
        max_age = max_age_seconds or settings.max_file_age_seconds
        current_time = time.time()
        cleaned_count = 0
        total_size_freed = 0

        for directory in [self.uploads_dir, self.converted_dir]:
            if not directory.exists():
                continue

            for file_path in directory.rglob("*"):
                if not file_path.is_file():
                    continue

                try:
                    file_age = current_time - file_path.stat().st_mtime
                    if file_age > max_age:
                        file_size = file_path.stat().st_size
                        file_path.unlink()
                        cleaned_count += 1
                        total_size_freed += file_size
                except Exception as e:
                    logger.error(f"删除文件失败 {file_path}: {e}")

        if cleaned_count > 0:
            message = f"按年龄清理: 删除 {cleaned_count} 个文件, 释放空间 {self._format_size(total_size_freed)}"
            self._log_cleanup(message)
            logger.info(message)

        return cleaned_count

    async def cleanup_by_size(self, max_size_bytes: Optional[int] = None) -> int:
        """根据目录大小清理最旧文件

        Args:
            max_size_bytes: 最大目录大小（字节），默认使用配置值

        Returns:
            int: 清理的文件数量
        """
        max_size = max_size_bytes or settings.max_temp_size_bytes
        total_size = self._get_dir_size(self.temp_dir)

        if total_size <= max_size:
            return 0

        # 收集所有文件及其修改时间
        files_with_time = []
        for directory in [self.uploads_dir, self.converted_dir]:
            if not directory.exists():
                continue

            for file_path in directory.rglob("*"):
                if not file_path.is_file():
                    continue

                try:
                    files_with_time.append((
                        file_path,
                        file_path.stat().st_mtime,
                        file_path.stat().st_size
                    ))
                except Exception:
                    continue

        # 按修改时间排序（最旧的在前）
        files_with_time.sort(key=lambda x: x[1])

        # 删除最旧的文件直到大小低于阈值
        cleaned_count = 0
        total_size_freed = 0

        for file_path, mtime, file_size in files_with_time:
            if total_size <= max_size:
                break

            try:
                file_path.unlink()
                total_size -= file_size
                total_size_freed += file_size
                cleaned_count += 1
            except Exception as e:
                logger.error(f"删除文件失败 {file_path}: {e}")

        if cleaned_count > 0:
            message = f"按大小清理: 删除 {cleaned_count} 个文件, 释放空间 {self._format_size(total_size_freed)}"
            self._log_cleanup(message)
            logger.info(message)

        return cleaned_count

    async def cleanup_task_directory(self, task_id: str) -> int:
        """清理指定任务的所有文件

        Args:
            task_id: 任务ID

        Returns:
            int: 清理的文件数量
        """
        cleaned_count = 0

        for directory in [self.uploads_dir, self.converted_dir]:
            task_dir = directory / task_id
            if task_dir.exists():
                try:
                    shutil.rmtree(task_dir)
                    cleaned_count += 1
                except Exception as e:
                    logger.error(f"删除任务目录失败 {task_dir}: {e}")

        if cleaned_count > 0:
            message = f"清理任务 {task_id}: 删除 {cleaned_count} 个目录"
            self._log_cleanup(message)
            logger.info(message)

        return cleaned_count

    async def cleanup_all(self) -> dict:
        """执行所有清理策略

        Returns:
            dict: 清理结果统计
        """
        results = {
            "by_age": await self.cleanup_by_age(),
            "by_size": await self.cleanup_by_size(),
            "current_size": self._format_size(self._get_dir_size(self.temp_dir)),
            "timestamp": datetime.now().isoformat()
        }

        return results

    async def startup_cleanup(self) -> dict:
        """服务启动时执行清理"""
        if not settings.enable_startup_cleanup:
            return {"status": "disabled"}

        logger.info("执行启动时清理...")
        results = await self.cleanup_all()
        logger.info(f"启动清理完成: {results}")
        return results

    async def periodic_cleanup(self):
        """定期清理任务（作为后台任务运行）"""
        while True:
            try:
                await asyncio.sleep(settings.cleanup_interval_hours * 3600)
                logger.info("执行定期清理...")
                await self.cleanup_all()
            except asyncio.CancelledError:
                logger.info("定期清理任务已取消")
                break
            except Exception as e:
                logger.error(f"定期清理失败: {e}")


# 全局清理器实例
cleaner = TempFileCleaner()


def get_cleaner() -> TempFileCleaner:
    """获取清理器实例"""
    return cleaner
