from datetime import datetime

from pydantic import BaseModel, EmailStr, Field


class LoginRequest(BaseModel):
    email: EmailStr
    password: str


class LoginResponse(BaseModel):
    access_token: str
    token_type: str = "bearer"
    user_name: str


class MeetingCreateText(BaseModel):
    title: str = Field(..., max_length=200)
    transcript: str = Field(..., max_length=200_000)
    source_type: str = "manual"


class MeetingOut(BaseModel):
    id: int
    title: str
    source_type: str
    detected_language: str
    ata_formal: str
    resumo_executivo: str
    tarefas: list[dict]
    decisoes: list[str]
    riscos: list[str]
    speakers: list[str]
    transcript_clean: str
    created_at: datetime


class ChatRequest(BaseModel):
    question: str = Field(..., max_length=2_000)


class RegisterRequest(BaseModel):
    email: EmailStr
    full_name: str
    password: str


class UserProfileOut(BaseModel):
    email: EmailStr
    full_name: str


class UserProfileUpdate(BaseModel):
    full_name: str


class ChangePasswordRequest(BaseModel):
    current_password: str
    new_password: str
