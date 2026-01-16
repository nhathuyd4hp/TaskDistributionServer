import uuid
from datetime import datetime, timedelta
from pathlib import Path

from fastapi import APIRouter, File, UploadFile
from fastapi.responses import StreamingResponse

from src.api.common.response import SuccessResponse
from src.core.config import settings
from src.service import ResultService as minio

router = APIRouter(prefix="/assets", tags=["Upload"])


@router.post(
    path="",
    name="Upload Asset",
    description=f"Upload File | Tồn tại trong {settings.ASSET_RETENTION_DAYS} ngày",
    response_model=SuccessResponse,
)
async def upload_asset(file: UploadFile = File(...)):
    file_extension = Path(file.filename).suffix
    new_object_name = f"{uuid.uuid4()}{file_extension}"
    file.file.seek(0, 2)
    file_size = file.file.tell()
    file.file.seek(0)
    result = minio.put_object(
        bucket_name=settings.TEMP_BUCKET,
        object_name=new_object_name,
        data=file.file,
        length=file_size,
        content_type=file.content_type,
        metadata={"expires_at": (datetime.now() + timedelta(days=settings.ASSET_RETENTION_DAYS)).isoformat()},
    )
    return SuccessResponse(data=f"{settings.TEMP_BUCKET}/{result.object_name}")


@router.get(
    path="/{bucket}/{objectName}",
    response_model=SuccessResponse,
)
async def get_asset(bucket: str, objectName: str):
    obj = minio.get_object(bucket, objectName)
    return StreamingResponse(
        obj,
        media_type=obj.headers.get("Content-Type", "application/octet-stream"),
        headers={"Content-Disposition": f'attachment; filename="{Path(objectName).name}"'},
    )
