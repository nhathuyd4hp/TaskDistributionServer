from datetime import datetime, timezone
from uuid import uuid4

from pydantic import field_serializer
from sqlalchemy import DateTime
from sqlmodel import Field, SQLModel


class Base(SQLModel):
    id: str = Field(default_factory=lambda: str(uuid4()), primary_key=True)

    created_at: datetime = Field(
        default_factory=lambda: datetime.now(timezone.utc), sa_type=DateTime(timezone=True), nullable=False
    )

    updated_at: datetime = Field(
        default_factory=lambda: datetime.now(timezone.utc),
        sa_type=DateTime(timezone=True),
        nullable=False,
        sa_column_kwargs={"onupdate": lambda: datetime.now(timezone.utc)},
    )

    @field_serializer("created_at", "updated_at")
    def serialize_dt(self, dt: datetime, _info):
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.isoformat().replace("+00:00", "Z")
