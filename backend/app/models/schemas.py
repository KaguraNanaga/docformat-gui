"""API 数据模型"""
from typing import Optional, List, Dict, Any
from pydantic import BaseModel, Field
from datetime import datetime


class ConvertRequest(BaseModel):
    """文件转换请求模型"""
    target_format: str = Field(..., description="目标格式", example="pdf")
    cleanup_after: bool = Field(default=True, description="转换后是否清理临时文件")


class ConvertResponse(BaseModel):
    """文件转换响应模型"""
    success: bool = Field(..., description="是否成功")
    message: str = Field(..., description="结果消息")
    task_id: Optional[str] = Field(None, description="任务ID")
    output_filename: Optional[str] = Field(None, description="输出文件名")
    output_url: Optional[str] = Field(None, description="输出文件下载URL")


class ConvertWithFileResponse(BaseModel):
    """带文件的转换响应（用于文件下载）"""
    success: bool = Field(..., description="是否成功")
    message: str = Field(..., description="结果消息")
    task_id: Optional[str] = Field(None, description="任务ID")
    filename: Optional[str] = Field(None, description="文件名")
    content_type: Optional[str] = Field(None, description="内容类型")


class FormatInfo(BaseModel):
    """格式信息模型"""
    name: str = Field(..., description="格式名称")
    extensions: List[str] = Field(..., description="文件扩展名列表")
    category: str = Field(..., description="格式分类：document/spreadsheet/presentation")


class ConversionPath(BaseModel):
    """转换路径模型"""
    source: str = Field(..., description="源格式")
    target: str = Field(..., description="目标格式")
    supported: bool = Field(..., description="是否支持")


class FormatsResponse(BaseModel):
    """支持的格式响应模型"""
    document_formats: List[str] = Field(..., description="支持的文档格式")
    spreadsheet_formats: List[str] = Field(..., description="支持的表格格式")
    all_conversions: List[ConversionPath] = Field(..., description="所有支持的转换路径")


class HealthResponse(BaseModel):
    """健康检查响应模型"""
    status: str = Field(..., description="服务状态")
    timestamp: datetime = Field(default_factory=datetime.now, description="检查时间")
    libreoffice_available: bool = Field(..., description="LibreOffice 是否可用")
    temp_dir_size: Optional[str] = Field(None, description="临时目录大小")


class CleanupResponse(BaseModel):
    """清理操作响应模型"""
    cleaned_by_age: int = Field(..., description="按年龄清理的文件数")
    cleaned_by_size: int = Field(..., description="按大小清理的文件数")
    current_size: str = Field(..., description="当前临时目录大小")
    timestamp: str = Field(..., description="清理时间")


class ErrorResponse(BaseModel):
    """错误响应模型"""
    error: str = Field(..., description="错误类型")
    message: str = Field(..., description="错误详情")
    detail: Optional[Dict[str, Any]] = Field(None, description="额外详情")
