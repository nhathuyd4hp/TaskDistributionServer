"""make created_at updated_at timezone aware

Revision ID: 424f429cdd0c
Revises: cea5873d8da4
Create Date: 2026-01-23 10:44:09.860084

"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa
import sqlmodel
from sqlalchemy.dialects import mysql

# revision identifiers, used by Alembic.
revision: str = '424f429cdd0c'
down_revision: Union[str, Sequence[str], None] = 'cea5873d8da4'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    pass


def downgrade() -> None:
    pass
