from fastapi import APIRouter

from src.api.router.log import router as LogRouter
from src.api.router.robot import router as RobotRouter
from src.api.router.run import router as RunRouter
from src.api.router.schedule import router as ScheduleRouter
from src.api.router.type import router as TypeRouter
from src.api.router.upload import router as UploadRouter

api = APIRouter()
api.include_router(RobotRouter)
api.include_router(RunRouter)
api.include_router(ScheduleRouter)
api.include_router(TypeRouter)
api.include_router(LogRouter)
api.include_router(UploadRouter)
