import json
import os
import time
from collections import defaultdict, deque

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
from sqlalchemy.orm import Session

from .ai_pipeline import chat_with_meeting, detect_language, normalize_transcript, structure_meeting
from .auth import create_access_token, decode_token, get_password_hash, validate_password_strength, verify_password
from .database import Base, engine, get_db
from .exporters import build_docx, build_markdown, build_pdf
from .models import Meeting, User
from .schemas import (
    ChangePasswordRequest,
    ChatRequest,
    LoginRequest,
    LoginResponse,
    MeetingCreateText,
    RegisterRequest,
    UserProfileOut,
    UserProfileUpdate,
)
from .transcription import transcribe_audio_file

Base.metadata.create_all(bind=engine)

app = FastAPI(title="MeetFlow AI MVP")
DEFAULT_ADMIN_EMAIL = "admin@meetflow.app"
LEGACY_ADMIN_EMAIL = "admin@meetflow.local"
LOGIN_WINDOW_SECONDS = 15 * 60
LOGIN_MAX_ATTEMPTS = 7
LOGIN_LOCK_SECONDS = 15 * 60
_login_attempts: dict[str, deque[float]] = defaultdict(deque)
_login_lock_until: dict[str, float] = {}
# CORS: não use allow_origins=["*"] com allow_credentials=True (comportamento inválido; o browser bloqueia).
_cors = os.getenv(
    "CORS_ALLOW_ORIGINS",
    "http://localhost:3000,http://127.0.0.1:3000,http://localhost:3001,http://127.0.0.1:3001",
)
_cors_list = [o.strip() for o in _cors.split(",") if o.strip()]
app.add_middleware(
    CORSMiddleware,
    allow_origins=_cors_list,
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


def bootstrap_user(db: Session):
    legacy = db.query(User).filter(User.email == LEGACY_ADMIN_EMAIL).first()
    current = db.query(User).filter(User.email == DEFAULT_ADMIN_EMAIL).first()

    # Migração de compatibilidade: email antigo era inválido para EmailStr no login.
    if legacy and not current:
        legacy.email = DEFAULT_ADMIN_EMAIL
        db.commit()
        current = legacy

    if not current:
        user = User(
            email=DEFAULT_ADMIN_EMAIL,
            full_name="MeetFlow Owner",
            password_hash=get_password_hash("admin123"),
        )
        db.add(user)
        db.commit()


def _login_lock_remaining(identity: str) -> int:
    now = time.time()
    lock_until = _login_lock_until.get(identity, 0)
    if lock_until <= now:
        _login_lock_until.pop(identity, None)
        return 0
    return int(lock_until - now)


def _register_login_failure(identity: str):
    now = time.time()
    attempts = _login_attempts[identity]
    attempts.append(now)
    while attempts and now - attempts[0] > LOGIN_WINDOW_SECONDS:
        attempts.popleft()
    if len(attempts) >= LOGIN_MAX_ATTEMPTS:
        _login_lock_until[identity] = now + LOGIN_LOCK_SECONDS
        attempts.clear()


def _clear_login_failures(identity: str):
    _login_attempts.pop(identity, None)
    _login_lock_until.pop(identity, None)


def get_current_user(authorization: str = Header(default=""), db: Session = Depends(get_db)) -> User:
    token = authorization.replace("Bearer ", "").strip()
    payload = decode_token(token)
    if not payload:
        raise HTTPException(status_code=401, detail="Token inválido ou expirado.")
    email = payload.get("sub")
    user = db.query(User).filter(User.email == email).first()
    if not user:
        raise HTTPException(status_code=401, detail="Usuário não encontrado.")
    return user


@app.on_event("startup")
def startup_event():
    db = next(get_db())
    bootstrap_user(db)


@app.get("/api/health")
def health():
    return {"status": "ok", "service": "meetflow-fastapi"}


@app.post("/api/auth/login", response_model=LoginResponse)
def login(payload: LoginRequest, db: Session = Depends(get_db)):
    email = payload.email.strip().lower()
    retry_after = _login_lock_remaining(email)
    if retry_after > 0:
        raise HTTPException(
            status_code=429,
            detail={
                "message": "Muitas tentativas. Tente novamente em alguns minutos.",
                "retry_after_seconds": retry_after,
            },
            headers={"Retry-After": str(retry_after)},
        )
    user = db.query(User).filter(User.email == email).first()
    if not user or not verify_password(payload.password, user.password_hash):
        _register_login_failure(email)
        raise HTTPException(status_code=401, detail="Credenciais inválidas.")
    _clear_login_failures(email)
    token = create_access_token(subject=user.email)
    return LoginResponse(access_token=token, user_name=user.full_name)


@app.post("/api/auth/register", response_model=LoginResponse)
def register(payload: RegisterRequest, db: Session = Depends(get_db)):
    email = payload.email.strip().lower()
    full_name = (payload.full_name or "").strip()
    if not full_name:
        raise HTTPException(status_code=422, detail="Informe seu nome.")
    pwd_error = validate_password_strength(payload.password)
    if pwd_error:
        raise HTTPException(status_code=422, detail=pwd_error)

    exists = db.query(User).filter(User.email == email).first()
    if exists:
        raise HTTPException(status_code=409, detail="Este e-mail já está em uso.")

    user = User(
        email=email,
        full_name=full_name,
        password_hash=get_password_hash(payload.password),
    )
    db.add(user)
    db.commit()
    db.refresh(user)
    token = create_access_token(subject=user.email)
    return LoginResponse(access_token=token, user_name=user.full_name)


@app.get("/api/me", response_model=UserProfileOut)
def get_me(current_user: User = Depends(get_current_user)):
    return UserProfileOut(email=current_user.email, full_name=current_user.full_name)


@app.patch("/api/me", response_model=UserProfileOut)
def update_me(payload: UserProfileUpdate, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    full_name = (payload.full_name or "").strip()
    if not full_name:
        raise HTTPException(status_code=422, detail="Informe seu nome.")
    current_user.full_name = full_name
    db.commit()
    db.refresh(current_user)
    return UserProfileOut(email=current_user.email, full_name=current_user.full_name)


@app.post("/api/me/change-password")
def change_password(
    payload: ChangePasswordRequest,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    if not verify_password(payload.current_password, current_user.password_hash):
        raise HTTPException(status_code=401, detail="Senha atual inválida.")
    pwd_error = validate_password_strength(payload.new_password)
    if pwd_error:
        raise HTTPException(status_code=422, detail=pwd_error)
    if verify_password(payload.new_password, current_user.password_hash):
        raise HTTPException(status_code=422, detail="A nova senha deve ser diferente da senha atual.")

    current_user.password_hash = get_password_hash(payload.new_password)
    db.commit()
    return {"status": "ok"}


@app.get("/api/dashboard")
def dashboard(current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    meetings = db.query(Meeting).filter(Meeting.user_id == current_user.id).order_by(Meeting.created_at.desc()).all()
    tasks_count = sum(len(json.loads(m.tarefas or "[]")) for m in meetings)
    decisions_count = sum(len(json.loads(m.decisoes or "[]")) for m in meetings)
    return {
        "meetings_count": len(meetings),
        "tasks_count": tasks_count,
        "decisions_count": decisions_count,
    }


@app.get("/api/meetings")
def list_meetings(current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    items = db.query(Meeting).filter(Meeting.user_id == current_user.id).order_by(Meeting.created_at.desc()).all()
    return [
        {
            "id": m.id,
            "title": m.title,
            "source_type": m.source_type,
            "detected_language": m.detected_language,
            "created_at": m.created_at,
        }
        for m in items
    ]


def serialize_meeting(m: Meeting):
    return {
        "id": m.id,
        "title": m.title,
        "source_type": m.source_type,
        "detected_language": m.detected_language,
        "ata_formal": m.ata_formal,
        "resumo_executivo": m.resumo_executivo,
        "tarefas": json.loads(m.tarefas or "[]"),
        "decisoes": json.loads(m.decisoes or "[]"),
        "riscos": json.loads(m.riscos or "[]"),
        "speakers": json.loads(m.speakers or "[]"),
        "transcript_clean": m.transcript_clean,
        "created_at": m.created_at,
    }


@app.get("/api/meetings/{meeting_id}")
def meeting_detail(meeting_id: int, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    m = db.query(Meeting).filter(Meeting.id == meeting_id, Meeting.user_id == current_user.id).first()
    if not m:
        raise HTTPException(status_code=404, detail="Reunião não encontrada.")
    return serialize_meeting(m)


@app.post("/api/meetings/process-text")
def process_text(payload: MeetingCreateText, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    clean = normalize_transcript(payload.transcript)
    if not clean:
        raise HTTPException(
            status_code=422,
            detail="A transcrição não pode estar vazia. Cole o texto ou use carregamento/gravação.",
        )
    if not (payload.title or "").strip():
        raise HTTPException(status_code=422, detail="Informe um título para a reunião.")
    language = detect_language(clean)
    structured = structure_meeting(payload.title, clean, language)

    meeting = Meeting(
        user_id=current_user.id,
        title=payload.title,
        source_type=payload.source_type,
        transcript_raw=payload.transcript,
        transcript_clean=clean,
        detected_language=language,
        ata_formal=structured.get("ata_formal", ""),
        resumo_executivo=structured.get("resumo_executivo", ""),
        tarefas=json.dumps(structured.get("tarefas", []), ensure_ascii=False),
        decisoes=json.dumps(structured.get("decisoes", []), ensure_ascii=False),
        riscos=json.dumps(structured.get("riscos", []), ensure_ascii=False),
        speakers=json.dumps(structured.get("speakers", []), ensure_ascii=False),
    )
    db.add(meeting)
    db.commit()
    db.refresh(meeting)
    return serialize_meeting(meeting)


@app.post("/api/meetings/process-upload")
async def process_upload(
    title: str = Form(...),
    file: UploadFile = File(...),
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    content = await file.read()
    transcript = transcribe_audio_file(content, file.filename or "audio.wav")
    clean = normalize_transcript(transcript)
    if not clean:
        raise HTTPException(
            status_code=422,
            detail="Não foi possível obter texto a partir do arquivo. Verifique o áudio, o formato ou informe transcrição manual.",
        )
    if not (title or "").strip():
        raise HTTPException(status_code=422, detail="Informe um título para a reunião.")
    language = detect_language(clean)
    structured = structure_meeting(title, clean, language)

    meeting = Meeting(
        user_id=current_user.id,
        title=title,
        source_type="upload",
        transcript_raw=transcript,
        transcript_clean=clean,
        detected_language=language,
        ata_formal=structured.get("ata_formal", ""),
        resumo_executivo=structured.get("resumo_executivo", ""),
        tarefas=json.dumps(structured.get("tarefas", []), ensure_ascii=False),
        decisoes=json.dumps(structured.get("decisoes", []), ensure_ascii=False),
        riscos=json.dumps(structured.get("riscos", []), ensure_ascii=False),
        speakers=json.dumps(structured.get("speakers", []), ensure_ascii=False),
    )
    db.add(meeting)
    db.commit()
    db.refresh(meeting)
    return serialize_meeting(meeting)


@app.post("/api/meetings/{meeting_id}/chat")
def meeting_chat(meeting_id: int, payload: ChatRequest, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    m = db.query(Meeting).filter(Meeting.id == meeting_id, Meeting.user_id == current_user.id).first()
    if not m:
        raise HTTPException(status_code=404, detail="Reunião não encontrada.")
    q = (payload.question or "").strip()
    if not q:
        raise HTTPException(status_code=422, detail="Escreva uma pergunta antes de enviar.")
    result = chat_with_meeting(q, serialize_meeting(m))
    if isinstance(result, dict):
        return result
    return {"answer": result, "suggested_questions": []}


@app.delete("/api/meetings/{meeting_id}", status_code=204)
def delete_meeting(meeting_id: int, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    deleted = db.query(Meeting).filter(Meeting.id == meeting_id, Meeting.user_id == current_user.id).delete()
    if not deleted:
        raise HTTPException(status_code=404, detail="Reunião não encontrada.")
    db.commit()
    return Response(status_code=204)


@app.get("/api/meetings/{meeting_id}/export/{fmt}")
def export_meeting(meeting_id: int, fmt: str, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    m = db.query(Meeting).filter(Meeting.id == meeting_id, Meeting.user_id == current_user.id).first()
    if not m:
        raise HTTPException(status_code=404, detail="Reunião não encontrada.")
    data = serialize_meeting(m)

    if fmt == "md":
        return Response(build_markdown(data), media_type="text/markdown")
    if fmt == "docx":
        return Response(
            build_docx(data),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    if fmt == "pdf":
        return Response(build_pdf(data), media_type="application/pdf")
    raise HTTPException(status_code=400, detail="Formato não suportado.")
