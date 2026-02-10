"""FastAPI 文档转换服务主入口"""
import asyncio
import logging
from contextlib import asynccontextmanager
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from app.core.config import settings
from app.api.convert import router as convert_router
from app.utils.cleanup import TempFileCleaner

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

# 全局清理器实例
cleaner: TempFileCleaner = TempFileCleaner()


@asynccontextmanager
async def lifespan(app: FastAPI):
    """应用生命周期管理"""
    # 启动时
    logger.info("启动文档转换服务...")

    # 执行启动时清理
    try:
        startup_results = await cleaner.startup_cleanup()
        logger.info(f"启动清理完成: {startup_results}")
    except Exception as e:
        logger.error(f"启动清理失败: {e}")

    # 启动定期清理任务
    if settings.cleanup_interval_hours > 0:
        cleanup_task = asyncio.create_task(cleaner.periodic_cleanup())
        app.state.cleanup_task = cleanup_task
        logger.info(f"定期清理任务已启动（间隔：{settings.cleanup_interval_hours}小时）")

    logger.info("文档转换服务启动完成")

    yield

    # 关闭时
    logger.info("正在关闭文档转换服务...")

    # 取消定期清理任务
    if hasattr(app.state, "cleanup_task"):
        app.state.cleanup_task.cancel()
        try:
            await app.state.cleanup_task
        except asyncio.CancelledError:
            pass

    logger.info("文档转换服务已关闭")


# 创建 FastAPI 应用
app = FastAPI(
    title="文档转换服务",
    description="基于 LibreOffice 的文档格式转换 API",
    version="1.0.0",
    lifespan=lifespan,
    docs_url=f"{settings.api_prefix}/docs",
    redoc_url=f"{settings.api_prefix}/redoc",
    openapi_url=f"{settings.api_prefix}/openapi.json"
)

# 配置 CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.allowed_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 注册路由
app.include_router(convert_router, prefix=settings.api_prefix)


@app.get("/")
async def root():
    """根路径"""
    return {
        "service": "文档转换服务",
        "version": "1.0.0",
        "docs": f"{settings.api_prefix}/docs",
        "health": f"{settings.api_prefix}/convert/health",
        "formats": f"{settings.api_prefix}/convert/formats"
    }


@app.get(f"{settings.api_prefix}")
async def api_root():
    """API 根路径"""
    return {
        "service": "文档转换 API",
        "version": "1.0.0",
        "endpoints": {
            "convert_document": f"{settings.api_prefix}/convert/document",
            "convert_spreadsheet": f"{settings.api_prefix}/convert/spreadsheet",
            "supported_formats": f"{settings.api_prefix}/convert/formats",
            "health_check": f"{settings.api_prefix}/convert/health",
            "cleanup": f"{settings.api_prefix}/convert/cleanup"
        }
    }


# 全局异常处理
@app.exception_handler(Exception)
async def global_exception_handler(request, exc):
    """全局异常处理器"""
    logger.error(f"未处理的异常: {exc}", exc_info=True)
    return JSONResponse(
        status_code=500,
        content={
            "error": "Internal Server Error",
            "message": "服务器内部错误",
            "detail": str(exc) if logger.level == logging.DEBUG else None
        }
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "app.main:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
        log_level="info"
    )
