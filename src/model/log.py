from datetime import datetime

from sqlmodel import Column, Field, Relationship, Text

from src.model.base import Base
from src.model.runs import Runs


class Log(Base, table=True):
    #
    run_id: str = Field(foreign_key="runs.id", index=True)
    run: Runs = Relationship(back_populates="logs")
    #
    timestamp: datetime = Field()
    level: str = Field()
    message: str | None = Field(sa_column=Column(Text))
