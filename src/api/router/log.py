import asyncio
import os

import aiofiles
from fastapi import APIRouter
from fastapi.responses import StreamingResponse

router = APIRouter(prefix="/logs", tags=["Log"])


async def streaming(run_id: str):
    file_path = f"logs/{run_id}.log"

    while True:
        if not os.path.exists(file_path):
            await asyncio.sleep(5)
            continue
        break

    async with aiofiles.open(file_path, mode="r", encoding="cp932") as f:
        while True:
            line = await f.readline()
            if line:
                yield f"{line}\n"
            else:
                await asyncio.sleep(0.25)


@router.get("/{run_id}")
async def log_realtime(run_id: str):
    return StreamingResponse(streaming(run_id), media_type="text/event-stream")
