import os
import tempfile


def transcribe_audio_file(file_bytes: bytes, filename: str) -> str:
    """Transcricao MVP com Whisper local quando disponivel."""
    try:
        import whisper

        model_name = os.getenv("WHISPER_MODEL", "base")
        model = whisper.load_model(model_name)
        suffix = os.path.splitext(filename)[1] or ".wav"
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as temp:
            temp.write(file_bytes)
            temp_path = temp.name
        result = model.transcribe(temp_path)
        return result.get("text", "").strip()
    except Exception:
        return "Nao foi possivel transcrever localmente. Para teste MVP, envie transcricao em texto na opcao manual."
