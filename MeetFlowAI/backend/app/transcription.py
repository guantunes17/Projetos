import os
import tempfile


def transcribe_audio_file(file_bytes: bytes, filename: str) -> str:
    """Transcrição MVP com Whisper local quando disponível."""
    try:
        import whisper
    except ImportError:
        return ""

    model_name = os.getenv("WHISPER_MODEL", "base")
    temp_path = None
    try:
        model = whisper.load_model(model_name)
        suffix = os.path.splitext(filename)[1] or ".wav"
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as temp:
            temp.write(file_bytes)
            temp_path = temp.name
        result = model.transcribe(temp_path)
        return result.get("text", "").strip()
    except Exception:
        return ""
    finally:
        if temp_path:
            try:
                os.unlink(temp_path)
            except OSError:
                pass
