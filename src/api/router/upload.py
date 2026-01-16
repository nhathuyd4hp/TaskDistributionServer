import uuid
from pathlib import Path
from fastapi import APIRouter, UploadFile, File
from src.service import ResultService as minio
from src.api.common.response import SuccessResponse

router = APIRouter(prefix="/uploads", tags=["Upload"])

@router.post(
    path="/{bucket_name}",
    name="Upload Asset",
    response_model=SuccessResponse,
)
async def upload_file_to_minio(
    bucket_name: str, 
    file: UploadFile = File(...)
):
    if not minio.bucket_exists(bucket_name):
        minio.make_bucket(bucket_name)
    file_extension = Path(file.filename).suffix 
    new_object_name = f"{uuid.uuid4()}{file_extension}"
    file.file.seek(0, 2)
    file_size = file.file.tell()
    file.file.seek(0)
    result = minio.put_object(
        bucket_name=bucket_name,
        object_name=new_object_name,
        data=file.file,
        length=file_size,
        content_type=file.content_type
    )
    return SuccessResponse(data=result.object_name)