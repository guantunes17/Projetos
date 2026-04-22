import json
import os
import re
from datetime import datetime

from openai import OpenAI


def normalize_transcript(transcript: str) -> str:
    cleaned = re.sub(r"\s+", " ", transcript).strip()
    return cleaned


def detect_language(transcript: str) -> str:
    text = transcript.lower()
    if any(token in text for token in [" the ", " project ", "deadline", "meeting"]):
        if any(token in text for token in [" nao ", " reunião", "prazo", "cliente"]):
            return "misto"
        return "en"
    return "pt"


def fallback_structuring(title: str, transcript: str, language: str) -> dict:
    chunks = [c.strip() for c in re.split(r"[.!?]", transcript) if c.strip()]
    bullet_chunks = chunks[:8]
    tasks = []
    for sentence in chunks:
        if any(key in sentence.lower() for key in ["vai", "precisa", "fazer", "entregar", "must", "need to"]):
            tasks.append({"task": sentence, "owner": "Nao identificado", "deadline": "Nao citado"})
    tasks = tasks[:10]

    decisions = [c for c in chunks if any(k in c.lower() for k in ["decid", "aprov", "combin", "agreed"])]
    risks = [c for c in chunks if any(k in c.lower() for k in ["risco", "atras", "bloque", "problema", "reclam"])]

    return {
        "ata_formal": (
            f"Ata da reuniao: {title}\n"
            f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}\n"
            f"Idioma detectado: {language}\n"
            "Participantes detectados: Nao identificado automaticamente.\n\n"
            "Temas discutidos:\n- " + "\n- ".join(bullet_chunks[:5]) + "\n\n"
            "Pendencias e proximos passos:\n- " + ("\n- ".join(bullet_chunks[5:8]) if len(bullet_chunks) > 5 else "Nao identificados.")
        ),
        "resumo_executivo": " ".join(bullet_chunks[:5]) or "Resumo nao disponivel.",
        "tarefas": tasks or [{"task": "Sem tarefas claras detectadas", "owner": "-", "deadline": "-"}],
        "decisoes": decisions[:8] or ["Sem decisoes explicitas detectadas."],
        "riscos": risks[:8] or ["Nenhum risco explicito detectado."],
        "speakers": ["Participante 1", "Participante 2"],
    }


def structure_meeting(title: str, transcript: str, language: str) -> dict:
    api_key = os.getenv("OPENAI_API_KEY", "")
    model = os.getenv("MEETFLOW_LLM_MODEL", "gpt-4o-mini")
    if not api_key:
        return fallback_structuring(title, transcript, language)

    client = OpenAI(api_key=api_key)
    prompt = f"""
Voce e um assistente de reunioes. Analise a transcricao e retorne JSON valido.

Formato esperado:
{{
  "ata_formal": "texto",
  "resumo_executivo": "texto objetivo",
  "tarefas": [{{"task":"", "owner":"", "deadline":""}}],
  "decisoes": ["..."],
  "riscos": ["..."],
  "speakers": ["..."]
}}

Titulo: {title}
Idioma detectado: {language}
Transcricao:
{transcript}
"""
    try:
        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
            response_format={"type": "json_object"},
        )
        content = response.choices[0].message.content or "{}"
        parsed = json.loads(content)
        return parsed
    except Exception:
        return fallback_structuring(title, transcript, language)


def chat_with_meeting(question: str, meeting: dict) -> str:
    api_key = os.getenv("OPENAI_API_KEY", "")
    model = os.getenv("MEETFLOW_LLM_MODEL", "gpt-4o-mini")
    context = (
        f"Titulo: {meeting['title']}\n"
        f"Resumo: {meeting['resumo_executivo']}\n"
        f"Transcricao: {meeting['transcript_clean']}\n"
        f"Tarefas: {meeting['tarefas']}\n"
        f"Decisoes: {meeting['decisoes']}\n"
        f"Riscos: {meeting['riscos']}"
    )
    if not api_key:
        if "prazo" in question.lower():
            return "Prazos citados: " + ", ".join([t.get("deadline", "-") for t in meeting["tarefas"]])
        if "respons" in question.lower():
            return "Responsaveis: " + ", ".join([t.get("owner", "-") for t in meeting["tarefas"]])
        return f"Baseado na reuniao: {meeting['resumo_executivo']}"

    client = OpenAI(api_key=api_key)
    prompt = f"Responda com base apenas no contexto da reuniao.\n\n{context}\n\nPergunta: {question}"
    try:
        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
        )
        return response.choices[0].message.content or "Sem resposta."
    except Exception:
        return "Nao foi possivel responder com IA no momento."
