import redis
from fastapi import APIRouter, Depends, HTTPException, Query, status
from sqlmodel import Session

from src.api.common.response import SuccessResponse
from src.api.dependency import get_session
from src.core.redis import REDIS_POOL
from src.model.runs import Status
from src.service import ErrorService, RunService

router = APIRouter(prefix="/runs", tags=["Runs"])


@router.get(
    path="",
    name="Lịch sử chạy [All]",
    response_model=SuccessResponse,
)
def get_histories(
    session: Session = Depends(get_session),
    status: Status | None = Query(default=None),
):
    if status:
        histories = RunService(session).findByStatus(status)
    else:
        histories = RunService(session).findMany()
    return SuccessResponse(data=histories)


@router.get(
    path="/{id}",
    name="Lịch sử chạy [1]",
    response_model=SuccessResponse,
)
def get_history(
    id: str,
    session: Session = Depends(get_session),
):
    history = RunService(session).findByID(id)
    if history is None:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="run not found")
    return SuccessResponse(data=history)


@router.get(
    path="/{id}/error",
    name="Lịch sử chạy [Error]",
    response_model=SuccessResponse,
)
def get_error(
    id: str,
    session: Session = Depends(get_session),
):
    error = ErrorService(session).findByRunID(id)
    if error is None:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="error not found")
    return SuccessResponse(data=error)


@router.post(
    path="/{id}/stop",
    name="Hủy/Dừng Robot",
    status_code=200,
    response_model=SuccessResponse,
)
def stop_running(
    id: str,
    session: Session = Depends(get_session),
):
    history = RunService(session).findByID(id)
    if history is None:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="run not found")
    redis.Redis(connection_pool=REDIS_POOL).set(id, "STOP", ex=24 * 60 * 60)
    return SuccessResponse(data=history)
