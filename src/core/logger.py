import json
import logging
import time

import redis


class RedisHandler(logging.Handler):
    def __init__(self, redisClient: redis.Redis, channel: str):
        super().__init__()
        self.client = redisClient
        self.channel = channel

    def emit(self, record: logging.LogRecord):
        message = self.format(record)
        data = {
            "task_id": self.channel,
            "message": message,
        }
        self.client.publish("LOG", json.dumps(data))


class Log:
    _initialized = False
    _loggers = {}

    @classmethod
    def _initialize(cls):
        if cls._initialized:
            return

        logging.Formatter.converter = time.gmtime

        cls.formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(name)s | %(message)s")

        cls._initialized = True

    @classmethod
    def get_logger(cls, channel: str, redis_client) -> logging.Logger:
        cls._initialize()
        if channel in cls._loggers:
            return cls._loggers[channel]
        logger = logging.getLogger(channel)
        logger.setLevel(logging.INFO)
        logger.propagate = False
        handler = RedisHandler(redis_client, channel)
        handler.setFormatter(cls.formatter)
        logger.addHandler(handler)
        cls._loggers[channel] = logger
        return logger

    @classmethod
    def delete_logger(cls, channel: str):
        logger = cls._loggers.pop(channel, None)
        if not logger:
            return
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)
            handler.close()
        logging.Logger.manager.loggerDict.pop(channel, None)
