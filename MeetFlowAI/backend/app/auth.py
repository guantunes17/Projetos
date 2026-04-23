import os
import re
from datetime import datetime, timedelta, timezone

from dotenv import load_dotenv
from jose import JWTError, jwt
from passlib.context import CryptContext

load_dotenv()

SECRET_KEY = os.getenv("APP_SECRET_KEY", "meetflow_dev_secret_change_me")
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_HOURS = 10

pwd_context = CryptContext(schemes=["pbkdf2_sha256"], deprecated="auto")


def verify_password(plain_password: str, hashed_password: str) -> bool:
    return pwd_context.verify(plain_password, hashed_password)


def get_password_hash(password: str) -> str:
    return pwd_context.hash(password)


def create_access_token(subject: str) -> str:
    now = datetime.now(timezone.utc)
    expire = now + timedelta(hours=ACCESS_TOKEN_EXPIRE_HOURS)
    payload = {"sub": subject, "iat": now, "exp": expire}
    return jwt.encode(payload, SECRET_KEY, algorithm=ALGORITHM)


def decode_token(token: str) -> dict | None:
    try:
        return jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
    except JWTError:
        return None


def validate_password_strength(password: str) -> str | None:
    """
    Regra forte:
    - mínimo 8 caracteres
    - pelo menos 1 maiúscula, 1 minúscula, 1 número e 1 símbolo
    Retorna mensagem em caso de erro, ou None quando válido.
    """
    if len(password or "") < 8:
        return "A senha deve ter pelo menos 8 caracteres."
    if not re.search(r"[A-Z]", password):
        return "A senha deve conter pelo menos 1 letra maiúscula."
    if not re.search(r"[a-z]", password):
        return "A senha deve conter pelo menos 1 letra minúscula."
    if not re.search(r"\d", password):
        return "A senha deve conter pelo menos 1 número."
    if not re.search(r"[^A-Za-z0-9]", password):
        return "A senha deve conter pelo menos 1 símbolo."
    return None
