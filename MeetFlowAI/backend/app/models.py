from datetime import datetime

from sqlalchemy import Boolean, DateTime, Integer, Text
from sqlalchemy.orm import Mapped, mapped_column

from .database import Base


class User(Base):
    __tablename__ = "users"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, index=True)
    email: Mapped[str] = mapped_column(Text, unique=True, index=True)
    full_name: Mapped[str] = mapped_column(Text)
    password_hash: Mapped[str] = mapped_column(Text)
    is_active: Mapped[bool] = mapped_column(Boolean, default=True)


class Meeting(Base):
    __tablename__ = "meetings"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, index=True)
    title: Mapped[str] = mapped_column(Text, index=True)
    source_type: Mapped[str] = mapped_column(Text)  # upload, transcript, manual, live
    transcript_raw: Mapped[str] = mapped_column(Text, default="")
    transcript_clean: Mapped[str] = mapped_column(Text, default="")
    detected_language: Mapped[str] = mapped_column(Text, default="pt")
    speakers: Mapped[str] = mapped_column(Text, default="[]")
    ata_formal: Mapped[str] = mapped_column(Text, default="")
    resumo_executivo: Mapped[str] = mapped_column(Text, default="")
    tarefas: Mapped[str] = mapped_column(Text, default="[]")
    decisoes: Mapped[str] = mapped_column(Text, default="[]")
    riscos: Mapped[str] = mapped_column(Text, default="[]")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
