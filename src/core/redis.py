import redis
from src.core.config import settings

REDIS_POOL = redis.ConnectionPool(
    host=settings.REDIS_HOST,
    port=settings.REDIS_PORT,
    db=settings.REDIS_DB,
    max_connections=10,
)
