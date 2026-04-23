"""Add user_id to meetings and index it.

Revision ID: 20260423_0001
Revises:
Create Date: 2026-04-23 00:00:00
"""

from alembic import op
import sqlalchemy as sa
from sqlalchemy import inspect


# revision identifiers, used by Alembic.
revision = "20260423_0001"
down_revision = None
branch_labels = None
depends_on = None


def upgrade() -> None:
    conn = op.get_bind()
    inspector = inspect(conn)
    existing_columns = {c["name"] for c in inspector.get_columns("meetings")}
    existing_indexes = {idx["name"] for idx in inspector.get_indexes("meetings")}

    if "user_id" not in existing_columns:
        with op.batch_alter_table("meetings") as batch_op:
            batch_op.add_column(sa.Column("user_id", sa.Integer(), nullable=True))

    if "ix_meetings_user_id" not in existing_indexes:
        with op.batch_alter_table("meetings") as batch_op:
            batch_op.create_index("ix_meetings_user_id", ["user_id"], unique=False)

    admin_row = conn.execute(sa.text("SELECT id FROM users WHERE email = 'admin@meetflow.app' LIMIT 1")).fetchone()
    if admin_row:
        conn.execute(sa.text("UPDATE meetings SET user_id = :uid WHERE user_id IS NULL"), {"uid": admin_row[0]})


def downgrade() -> None:
    conn = op.get_bind()
    inspector = inspect(conn)
    existing_columns = {c["name"] for c in inspector.get_columns("meetings")}
    existing_indexes = {idx["name"] for idx in inspector.get_indexes("meetings")}

    if "ix_meetings_user_id" in existing_indexes:
        with op.batch_alter_table("meetings") as batch_op:
            batch_op.drop_index("ix_meetings_user_id")
    if "user_id" in existing_columns:
        with op.batch_alter_table("meetings") as batch_op:
            batch_op.drop_column("user_id")
