"""文件转换 API 端点"""
import asyncio
import logging
import uuid
import tempfile
import os
import urllib.parse
from datetime import datetime
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, Response

from app.core.config import settings
from app.core.converter import LibreOfficeConverter, FileType
from app.models.schemas import (
    ConvertRequest, ConvertResponse, ConvertWithFileResponse,
    FormatsResponse, HealthResponse, CleanupResponse
)
from app.utils.cleanup import TempFileCleaner

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/convert", tags=["转换"])

# 全局实例
converter = LibreOfficeConverter()
cleaner = TempFileCleaner()

# 内容类型映射
CONTENT_TYPE_MAP = {
    "pdf": "application/pdf",
    "doc": "application/msword",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xls": "application/vnd.ms-excel",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "csv": "text/csv",
    "txt": "text/plain",
    "rtf": "application/rtf",
    "odt": "application/vnd.oasis.opendocument.text",
    "ods": "application/vnd.oasis.opendocument.spreadsheet",
    "html": "text/html",
}


@router.post("/document", response_model=ConvertWithFileResponse)
async def convert_document(
    file: UploadFile = File(..., description="文档文件"),
    target_format: str = Form(..., description="目标格式（docx, doc, pdf, wps）"),
    cleanup_after: bool = Form(True, description="转换后是否清理临时文件")
):
    """文档格式互转 API

    支持的格式：doc, docx, pdf, wps, odt, rtf, txt, html

    Args:
        file: 上传的文档文件
        target_format: 目标格式
        cleanup_after: 是否在转换后清理临时文件

    Returns:
        ConvertWithFileResponse: 转换结果（包含下载链接）
    """
    # 生成任务ID
    task_id = str(uuid.uuid4())

    # 验证文件扩展名
    file_ext = Path(file.filename).suffix.lower().lstrip(".")
    if file_ext not in converter.DOCUMENT_FORMATS:
        raise HTTPException(
            status_code=400,
            detail=f"不支持的文档格式: {file_ext}。支持的格式: {', '.join(converter.DOCUMENT_FORMATS)}"
        )

    # 验证目标格式
    target_format = target_format.lower()
    if target_format not in converter.DOCUMENT_FORMATS:
        raise HTTPException(
            status_code=400,
            detail=f"不支持的目标格式: {target_format}。支持的格式: {', '.join(converter.DOCUMENT_FORMATS)}"
        )

    # 检查转换是否支持
    if not converter.is_conversion_supported(file_ext, target_format):
        supported = converter.get_supported_conversions(file_ext)
        raise HTTPException(
            status_code=400,
            detail=f"不支持从 {file_ext} 转换为 {target_format}。支持的目标格式: {', '.join(supported)}"
        )

    try:
        # 创建任务目录
        task_upload_dir = cleaner.uploads_dir / task_id
        task_upload_dir.mkdir(parents=True, exist_ok=True)

        task_convert_dir = cleaner.converted_dir / task_id
        task_convert_dir.mkdir(parents=True, exist_ok=True)

        # 保存上传文件
        input_path = task_upload_dir / file.filename
        with open(input_path, "wb") as f:
            content = await file.read()
            f.write(content)

        logger.info(f"任务 {task_id}: 开始转换 {file.filename} -> {target_format}")

        # 执行转换
        success, output_path, error_msg = await converter.convert(
            input_path=input_path,
            target_format=target_format,
            output_dir=task_convert_dir
        )

        if not success or output_path is None:
            raise HTTPException(
                status_code=500,
                detail=f"转换失败: {error_msg}"
            )

        # 获取内容类型
        content_type = CONTENT_TYPE_MAP.get(target_format, "application/octet-stream")

        # 读取转换后的文件内容
        with open(output_path, "rb") as f:
            file_content = f.read()

        # 定义清理函数
        async def cleanup_task():
            if cleanup_after:
                await cleaner.cleanup_task_directory(task_id)
                logger.info(f"任务 {task_id}: 已清理临时文件")

        # 异步执行清理（不阻塞响应）
        asyncio.create_task(cleanup_task())

        logger.info(f"任务 {task_id}: 转换成功，输出文件: {output_path.name}")

        # 直接返回文件内容
        # 处理文件名中的中文字符
        filename = f"{Path(file.filename).stem}.{target_format}"
        encoded_filename = urllib.parse.quote(filename, safe='')

        return Response(
            content=file_content,
            media_type=content_type,
            headers={
                "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"
            }
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"转换异常: {e}")
        raise HTTPException(
            status_code=500,
            detail=f"转换过程中发生错误: {str(e)}"
        )


@router.post("/spreadsheet", response_model=ConvertWithFileResponse)
async def convert_spreadsheet(
    file: UploadFile = File(..., description="表格文件"),
    target_format: str = Form(..., description="目标格式（xlsx, xls, csv）"),
    cleanup_after: bool = Form(True, description="转换后是否清理临时文件")
):
    """表格格式互转 API

    支持的格式：xls, xlsx, csv, ods

    Args:
        file: 上传的表格文件
        target_format: 目标格式
        cleanup_after: 是否在转换后清理临时文件

    Returns:
        ConvertWithFileResponse: 转换结果（包含下载链接）
    """
    # 生成任务ID
    task_id = str(uuid.uuid4())

    # 验证文件扩展名
    file_ext = Path(file.filename).suffix.lower().lstrip(".")
    if file_ext not in converter.SPREADSHEET_FORMATS:
        raise HTTPException(
            status_code=400,
            detail=f"不支持的表格格式: {file_ext}。支持的格式: {', '.join(converter.SPREADSHEET_FORMATS)}"
        )

    # 验证目标格式
    target_format = target_format.lower()
    if target_format not in converter.SPREADSHEET_FORMATS:
        raise HTTPException(
            status_code=400,
            detail=f"不支持的目标格式: {target_format}。支持的格式: {', '.join(converter.SPREADSHEET_FORMATS)}"
        )

    # 检查转换是否支持
    if not converter.is_conversion_supported(file_ext, target_format):
        supported = converter.get_supported_conversions(file_ext)
        raise HTTPException(
            status_code=400,
            detail=f"不支持从 {file_ext} 转换为 {target_format}。支持的目标格式: {', '.join(supported)}"
        )

    try:
        # 创建任务目录
        task_upload_dir = cleaner.uploads_dir / task_id
        task_upload_dir.mkdir(parents=True, exist_ok=True)

        task_convert_dir = cleaner.converted_dir / task_id
        task_convert_dir.mkdir(parents=True, exist_ok=True)

        # 保存上传文件
        input_path = task_upload_dir / file.filename
        with open(input_path, "wb") as f:
            content = await file.read()
            f.write(content)

        logger.info(f"任务 {task_id}: 开始转换 {file.filename} -> {target_format}")

        # 执行转换
        success, output_path, error_msg = await converter.convert(
            input_path=input_path,
            target_format=target_format,
            output_dir=task_convert_dir
        )

        if not success or output_path is None:
            raise HTTPException(
                status_code=500,
                detail=f"转换失败: {error_msg}"
            )

        # 获取内容类型
        content_type = CONTENT_TYPE_MAP.get(target_format, "application/octet-stream")

        # 读取转换后的文件内容
        with open(output_path, "rb") as f:
            file_content = f.read()

        # 定义清理函数
        async def cleanup_task():
            if cleanup_after:
                await cleaner.cleanup_task_directory(task_id)
                logger.info(f"任务 {task_id}: 已清理临时文件")

        # 异步执行清理（不阻塞响应）
        asyncio.create_task(cleanup_task())

        logger.info(f"任务 {task_id}: 转换成功，输出文件: {output_path.name}")

        # 直接返回文件内容
        # 处理文件名中的中文字符
        filename = f"{Path(file.filename).stem}.{target_format}"
        encoded_filename = urllib.parse.quote(filename, safe='')

        return Response(
            content=file_content,
            media_type=content_type,
            headers={
                "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"
            }
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"转换异常: {e}")
        raise HTTPException(
            status_code=500,
            detail=f"转换过程中发生错误: {str(e)}"
        )


@router.get("/formats", response_model=FormatsResponse)
async def get_supported_formats():
    """获取支持的格式列表

    Returns:
        FormatsResponse: 支持的格式信息
    """
    from app.models.schemas import ConversionPath

    # 生成所有转换路径
    all_conversions = []
    for source_format in converter.DOCUMENT_FORMATS + converter.SPREADSHEET_FORMATS:
        for target_format in converter.get_supported_conversions(source_format):
            all_conversions.append(ConversionPath(
                source=source_format,
                target=target_format,
                supported=True
            ))

    return FormatsResponse(
        document_formats=converter.DOCUMENT_FORMATS,
        spreadsheet_formats=converter.SPREADSHEET_FORMATS,
        all_conversions=all_conversions
    )


@router.post("/cleanup", response_model=CleanupResponse)
async def cleanup_temp_files():
    """手动触发临时文件清理

    Returns:
        CleanupResponse: 清理结果
    """
    results = await cleaner.cleanup_all()

    return CleanupResponse(
        cleaned_by_age=results["by_age"],
        cleaned_by_size=results["by_size"],
        current_size=results["current_size"],
        timestamp=results["timestamp"]
    )


@router.get("/health", response_model=HealthResponse)
async def health_check():
    """健康检查端点

    Returns:
        HealthResponse: 健康状态
    """
    # 检查 LibreOffice 容器是否可用
    try:
        process = await asyncio.create_subprocess_exec(
            "docker", "inspect", settings.libreoffice_container_name,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE
        )
        stdout, stderr = await process.communicate()
        libreoffice_available = process.returncode == 0
    except Exception:
        libreoffice_available = False

    # 获取临时目录大小
    temp_dir_size = None
    try:
        total_size = cleaner._get_dir_size(Path(settings.temp_dir))
        temp_dir_size = cleaner._format_size(total_size)
    except Exception:
        pass

    return HealthResponse(
        status="healthy" if libreoffice_available else "degraded",
        timestamp=datetime.now(),
        libreoffice_available=libreoffice_available,
        temp_dir_size=temp_dir_size
    )
