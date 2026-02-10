"""LibreOffice 文档转换核心模块"""
import asyncio
import logging
import shutil
import subprocess
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from enum import Enum

from app.core.config import settings

logger = logging.getLogger(__name__)


class FileType(Enum):
    """文件类型枚举"""
    DOCUMENT = "document"
    SPREADSHEET = "spreadsheet"
    PRESENTATION = "presentation"
    UNKNOWN = "unknown"


class LibreOfficeConverter:
    """LibreOffice 转换器"""

    # 格式映射表：源格式 -> {目标格式 -> LibreOffice 导出过滤器}
    FORMAT_MAP: Dict[str, Dict[str, str]] = {
        # 文档格式互转
        "doc": {
            "docx": "Office Open XML Text",
            "pdf": "writer_pdf_Export",
            "txt": "Text",
            "odt": "writer8",
            "html": "HTML (StarWriter)",
            "rtf": "Rich Text Format",
        },
        "docx": {
            "doc": "MS Word 97",
            "pdf": "writer_pdf_Export",
            "txt": "Text",
            "odt": "writer8",
            "html": "HTML (StarWriter)",
            "rtf": "Rich Text Format",
        },
        "pdf": {
            "doc": "MS Word 97",
            "docx": "Office Open XML Text",
            "odt": "writer8",
            "txt": "Text",
        },
        "wps": {
            "doc": "MS Word 97",
            "docx": "Office Open XML Text",
            "pdf": "writer_pdf_Export",
            "odt": "writer8",
        },
        "odt": {
            "doc": "MS Word 97",
            "docx": "Office Open XML Text",
            "pdf": "writer_pdf_Export",
            "txt": "Text",
            "html": "HTML (StarWriter)",
        },
        "rtf": {
            "doc": "MS Word 97",
            "docx": "Office Open XML Text",
            "pdf": "writer_pdf_Export",
            "txt": "Text",
        },

        # 表格格式互转
        "xls": {
            "xlsx": "Office Open XML Spreadsheet",
            "csv": "Text CSV",
            "pdf": "calc_pdf_Export",
            "ods": "calc8",
        },
        "xlsx": {
            "xls": "MS Excel 97",
            "csv": "Text CSV",
            "pdf": "calc_pdf_Export",
            "ods": "calc8",
        },
        "csv": {
            "xls": "MS Excel 97",
            "xlsx": "Office Open XML Spreadsheet",
            "ods": "calc8",
            "pdf": "calc_pdf_Export",
        },
        "ods": {
            "xls": "MS Excel 97",
            "xlsx": "Office Open XML Spreadsheet",
            "csv": "Text CSV",
            "pdf": "calc_pdf_Export",
        },
    }

    # 支持的文档格式列表
    DOCUMENT_FORMATS = ["doc", "docx", "pdf", "wps", "odt", "rtf", "txt", "html"]

    # 支持的表格格式列表
    SPREADSHEET_FORMATS = ["xls", "xlsx", "csv", "ods"]

    def __init__(self, container_name: Optional[str] = None):
        """初始化转换器

        Args:
            container_name: LibreOffice 容器名称，默认使用配置值
        """
        self.container_name = container_name or settings.libreoffice_container_name

    @classmethod
    def get_file_type(cls, file_path: Path) -> FileType:
        """获取文件类型

        Args:
            file_path: 文件路径

        Returns:
            FileType: 文件类型
        """
        ext = file_path.suffix.lower().lstrip(".")

        if ext in cls.DOCUMENT_FORMATS:
            return FileType.DOCUMENT
        elif ext in cls.SPREADSHEET_FORMATS:
            return FileType.SPREADSHEET
        else:
            return FileType.UNKNOWN

    @classmethod
    def get_supported_conversions(cls, source_format: str) -> List[str]:
        """获取指定源格式支持的目标格式

        Args:
            source_format: 源格式（如 "doc"）

        Returns:
            List[str]: 支持的目标格式列表
        """
        source_format = source_format.lower()
        return list(cls.FORMAT_MAP.get(source_format, {}).keys())

    @classmethod
    def is_conversion_supported(cls, source_format: str, target_format: str) -> bool:
        """检查是否支持指定转换

        Args:
            source_format: 源格式
            target_format: 目标格式

        Returns:
            bool: 是否支持
        """
        source_format = source_format.lower()
        target_format = target_format.lower()
        return target_format in cls.FORMAT_MAP.get(source_format, {})

    def _build_command(
        self,
        input_path: Path,
        output_format: str,
        output_dir: Path,
        export_filter: Optional[str] = None
    ) -> List[str]:
        """构建 LibreOffice 转换命令

        Args:
            input_path: 输入文件路径
            output_format: 输出格式
            output_dir: 输出目录
            export_filter: 导出过滤器（可选）

        Returns:
            List[str]: 命令列表
        """
        # 根据源文件类型选择命令
        # WPS 文件需要使用 soffice 命令
        source_format = input_path.suffix.lower().lstrip(".")
        lo_command = "soffice" if source_format == "wps" else "libreoffice"

        # 基础命令：在容器内执行 LibreOffice
        cmd = [
            "docker", "exec", self.container_name,
            lo_command,
            "--headless",
            "--nologo",
            "--nofirststartwizard",
            "--norestore",
            "--convert-to", output_format,
            "--outdir", str(output_dir),
            str(input_path)
        ]

        # 如果需要特定的导出过滤器
        if export_filter:
            cmd.extend(["--infilter", export_filter])

        return cmd

    def _get_export_filter(self, source_format: str, target_format: str) -> Optional[str]:
        """获取导出过滤器

        Args:
            source_format: 源格式
            target_format: 目标格式

        Returns:
            Optional[str]: 导出过滤器名称
        """
        source_format = source_format.lower()
        target_format = target_format.lower()
        return self.FORMAT_MAP.get(source_format, {}).get(target_format)

    async def convert(
        self,
        input_path: Path,
        target_format: str,
        output_dir: Optional[Path] = None,
        max_retries: Optional[int] = None
    ) -> Tuple[bool, Optional[Path], str]:
        """转换文件

        Args:
            input_path: 输入文件路径
            target_format: 目标格式
            output_dir: 输出目录（默认使用配置的 converted 目录）
            max_retries: 最大重试次数（默认使用配置值）

        Returns:
            Tuple[bool, Optional[Path], str]:
                (是否成功, 输出文件路径, 错误消息)
        """
        source_format = input_path.suffix.lower().lstrip(".")
        target_format = target_format.lower()

        # 验证输入文件
        if not input_path.exists():
            return False, None, f"输入文件不存在: {input_path}"

        # 检查转换是否支持
        if not self.is_conversion_supported(source_format, target_format):
            supported = self.get_supported_conversions(source_format)
            return False, None, f"不支持的转换: {source_format} -> {target_format}。支持的目标格式: {', '.join(supported)}"

        # 设置输出目录
        if output_dir is None:
            output_dir = Path(settings.converted_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        # 生成简单的输出文件名（避免空格等特殊字符）
        import hashlib
        file_hash = hashlib.md5(str(input_path).encode()).hexdigest()[:8]
        simple_filename = f"convert_{file_hash}.{source_format}"
        simple_output = f"convert_{file_hash}.{target_format}"

        # 创建符号链接或复制文件以避免特殊字符
        simple_input_path = input_path.parent / simple_filename
        try:
            import shutil
            shutil.copy(input_path, simple_input_path)
        except Exception as e:
            logger.warning(f"无法复制文件到简单名称: {e}")
            simple_input_path = input_path

        # 获取导出过滤器
        export_filter = self._get_export_filter(source_format, target_format)
        if export_filter:
            pass

        # 构建命令（使用简单文件名）
        cmd = self._build_command(simple_input_path, target_format, output_dir)

        # 重试机制
        max_attempts = max_retries or settings.max_retry_attempts
        timeout = settings.conversion_timeout_seconds

        for attempt in range(1, max_attempts + 1):
            try:
                logger.info(f"转换尝试 {attempt}/{max_attempts}: {input_path.name} -> {target_format}")

                # 异步执行命令
                process = await asyncio.create_subprocess_exec(
                    *cmd,
                    stdout=asyncio.subprocess.PIPE,
                    stderr=asyncio.subprocess.PIPE,
                )

                try:
                    stdout, stderr = await asyncio.wait_for(
                        process.communicate(),
                        timeout=timeout
                    )
                except asyncio.TimeoutError:
                    process.kill()
                    await process.wait()
                    raise TimeoutError(f"转换超时（{timeout}秒）")

                # 检查返回码
                if process.returncode == 0:
                    # 查找输出文件
                    expected_output = output_dir / simple_output

                    if expected_output.exists():
                        # 将输出文件重命名为原始文件名
                        final_output = output_dir / f"{input_path.stem}.{target_format}"
                        shutil.move(expected_output, final_output)
                        return True, final_output, ""
                    else:
                        # 尝试查找任何匹配的文件
                        for file in output_dir.glob(f"convert_{file_hash}*"):
                            if file.suffix.lower() == f".{target_format}":
                                final_output = output_dir / f"{input_path.stem}.{target_format}"
                                shutil.move(file, final_output)
                                return True, final_output, ""
                        return False, None, f"转换成功但未找到输出文件: {expected_output}"
                else:
                    error_msg = stderr.decode("utf-8", errors="ignore")
                    if attempt < max_attempts:
                        logger.warning(f"转换失败（尝试 {attempt}/{max_attempts}）: {error_msg}")
                        await asyncio.sleep(2)
                        continue
                    else:
                        return False, None, f"LibreOffice 转换失败: {error_msg}"

            except TimeoutError as e:
                if attempt < max_attempts:
                    logger.warning(f"{e}（尝试 {attempt}/{max_attempts}）")
                    await asyncio.sleep(2)
                    continue
                else:
                    return False, None, str(e)
            except Exception as e:
                if attempt < max_attempts:
                    logger.warning(f"转换异常（尝试 {attempt}/{max_attempts}）: {e}")
                    await asyncio.sleep(2)
                    continue
                else:
                    return False, None, f"转换异常: {e}"

        return False, None, f"转换失败，已重试 {max_attempts} 次"


# 全局转换器实例
converter = LibreOfficeConverter()


def get_converter() -> LibreOfficeConverter:
    """获取转换器实例"""
    return converter
