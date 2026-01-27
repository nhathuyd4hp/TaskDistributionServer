from fastapi import Request
from fastapi.responses import JSONResponse
from minio.error import S3Error
from sqlalchemy.exc import ProgrammingError
from starlette.middleware.base import BaseHTTPMiddleware


class GlobalExceptionMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        try:
            return await call_next(request)
        except ProgrammingError as e:
            return JSONResponse(
                status_code=500,
                content={
                    "success": False,
                    "message": str(e.orig),
                },
            )
        except S3Error as e:
            if e.code == "NoSuchKey":
                return JSONResponse(
                    status_code=404,
                    content={
                        "success": False,
                        "message": "asset not found",
                    },
                )
            if e.code == "NoSuchBucket":
                return JSONResponse(
                    status_code=404,
                    content={
                        "success": False,
                        "message": "bucket not found",
                    },
                )
            return JSONResponse(
                status_code=500,
                content={
                    "success": False,
                    "message": e.message,
                },
            )
        except Exception as e:
            print(e)
            return JSONResponse(
                status_code=500,
                content={
                    "success": False,
                    "message": str(e),
                },
            )
