"""add run_errors table

Revision ID: cea5873d8da4
Revises: 4e4711876b58
Create Date: 2026-01-20 11:07:06.910386

"""

from typing import Sequence, Union

import sqlalchemy as sa
import sqlmodel
from alembic import op

# revision identifiers, used by Alembic.
revision: str = "cea5873d8da4"  # pragma: allowlist secret
down_revision: Union[str, Sequence[str], None] = "4e4711876b58"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.create_table(
        "error",
        sa.Column("id", sqlmodel.sql.sqltypes.AutoString(), nullable=False),
        sa.Column("created_at", sa.DateTime(), nullable=False),
        sa.Column("updated_at", sa.DateTime(), nullable=False),
        sa.Column("run_id", sqlmodel.sql.sqltypes.AutoString(), nullable=False),
        sa.Column("error_type", sa.String(length=255), nullable=True),
        sa.Column("message", sa.String(length=1024), nullable=True),
        sa.Column("traceback", sa.Text(), nullable=False),
        sa.ForeignKeyConstraint(
            ["run_id"],
            ["runs.id"],
            ondelete="CASCADE",
            name="fk_error_run_id",
        ),
        sa.PrimaryKeyConstraint("id"),
    )
    op.create_index("ix_error_run_id", "error", ["run_id"])


def downgrade() -> None:
    op.drop_constraint(
        constraint_name="fk_error_run_id",
        table_name="error",
        type_="foreignkey",
    )
    op.drop_index("ix_error_run_id", table_name="error")
    op.drop_table("error")
