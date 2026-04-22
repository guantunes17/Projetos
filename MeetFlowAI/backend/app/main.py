import json

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
from sqlalchemy.orm import Session

from .ai_pipeline import chat_with_meeting, detect_language, normalize_transcript, structure_meeting
from .auth import create_access_token, decode_token, get_password_hash, verify_password
from .database import Base, engine, get_db
from .exporters import build_docx, build_markdown, build_pdf
from .models import Meeting, User
from .schemas import ChatRequest, LoginRequest, LoginResponse, MeetingCreateText
from .transcription import transcribe_audio_file

Base.metadata.create_all(bind=engine)

app = FastAPI(title="MeetFlow AI MVP")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def bootstrap_user(db: Session):
    existing = db.query(User).filter(User.email == "admin@meetflow.local").first()
    if not existing:
        user = User(
            email="admin@meetflow.local",
            full_name="MeetFlow Owner",
            password_hash=get_password_hash("admin123"),
        )
        db.add(user)
        db.commit()


def get_current_user(authorization: str = Header(default="")) -> User:
    token = authorization.replace("Bearer ", "").strip()
    payload = decode_token(token)
    if not payload:
        raise HTTPException(status_code=401, detail="Token invalido ou expirado.")
    email = payload.get("sub")
    db = next(get_db())
    user = db.query(User).filter(User.email == email).first()
    if not user:
        raise HTTPException(status_code=401, detail="Usuario nao encontrado.")
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
    user = db.query(User).filter(User.email == payload.email).first()
    if not user or not verify_password(payload.password, user.password_hash):
        raise HTTPException(status_code=401, detail="Credenciais invalidas.")
    token = create_access_token(subject=user.email)
    return LoginResponse(access_token=token, user_name=user.full_name)


@app.get("/api/dashboard")
def dashboard(_: User = Depends(get_current_user), db: Session = Depends(get_db)):
    meetings = db.query(Meeting).order_by(Meeting.created_at.desc()).all()
    tasks_count = sum(len(json.loads(m.tarefas or "[]")) for m in meetings)
    decisions_count = sum(len(json.loads(m.decisoes or "[]")) for m in meetings)
    return {
        "meetings_count": len(meetings),
        "tasks_count": tasks_count,
        "decisions_count": decisions_count,
    }


@app.get("/api/meetings")
def list_meetings(_: User = Depends(get_current_user), db: Session = Depends(get_db)):
    items = db.query(Meeting).order_by(Meeting.created_at.desc()).all()
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
def meeting_detail(meeting_id: int, _: User = Depends(get_current_user), db: Session = Depends(get_db)):
    m = db.query(Meeting).filter(Meeting.id == meeting_id).first()
    if not m:
        raise HTTPException(status_code=404, detail="Reuniao nao encontrada.")
    return serialize_meeting(m)


@app.post("/api/meetings/process-text")
def process_text(payload: MeetingCreateText, _: User = Depends(get_current_user), db: Session = Depends(get_db)):
    clean = normalize_transcript(payload.transcript)
    language = detect_language(clean)
    structured = structure_meeting(payload.title, clean, language)

    meeting = Meeting(
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
    _: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    content = await file.read()
    transcript = transcribe_audio_file(content, file.filename or "audio.wav")
    clean = normalize_transcript(transcript)
    language = detect_language(clean)
    structured = structure_meeting(title, clean, language)

    meeting = Meeting(
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
def meeting_chat(meeting_id: int, payload: ChatRequest, _: User = Depends(get_current_user), db: Session = Depends(get_db)):
    m = db.query(Meeting).filter(Meeting.id == meeting_id).first()
    if not m:
        raise HTTPException(status_code=404, detail="Reuniao nao encontrada.")
    answer = chat_with_meeting(payload.question, serialize_meeting(m))
    return {"answer": answer}


@app.get("/api/meetings/{meeting_id}/export/{fmt}")
def export_meeting(meeting_id: int, fmt: str, _: User = Depends(get_current_user), db: Session = Depends(get_db)):
    m = db.query(Meeting).filter(Meeting.id == meeting_id).first()
    if not m:
        raise HTTPException(status_code=404, detail="Reuniao nao encontrada.")
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
    raise HTTPException(status_code=400, detail="Formato nao suportado.")
