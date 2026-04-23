import json
import os
import re
from datetime import datetime

from openai import OpenAI

from .text_preprocess import preprocess_transcript

# Instrução reutilizada: saídas em português do Brasil com ortografia correta
ORTOGRAFIA_PT = (
    "Escreva sempre em português do Brasil (pt-BR), com acentuação e ortografia corretas "
    "(incluindo ã, õ, ç e acentos quando apropriado). Não use grafia sem acentos."
)


def normalize_transcript(transcript: str) -> str:
    """Compatível com chamadas existentes: delega para pré-processamento completo."""
    return preprocess_transcript(transcript)


def _model_structure() -> str:
    return os.getenv("MEETFLOW_LLM_MODEL_STRUCTURE") or os.getenv("MEETFLOW_LLM_MODEL", "gpt-4o-mini")


def _model_chat() -> str:
    return os.getenv("MEETFLOW_LLM_MODEL_CHAT") or os.getenv("MEETFLOW_LLM_MODEL", "gpt-4o-mini")


def detect_language(transcript: str) -> str:
    text = transcript.lower()
    if any(token in text for token in [" the ", " project ", "deadline", "meeting"]):
        if any(token in text for token in [" não ", " nao ", " reunião", "reuniao", "prazo", "cliente"]):
            return "misto"
        return "en"
    return "pt"


def fallback_structuring(title: str, transcript: str, language: str) -> dict:
    chunks = [c.strip() for c in re.split(r"[.!?]", transcript) if c.strip()]
    bullet_chunks = chunks[:8]
    tasks = []
    for sentence in chunks:
        if any(key in sentence.lower() for key in ["vai", "precisa", "fazer", "entregar", "must", "need to"]):
            tasks.append({"task": sentence, "owner": "Não identificado", "deadline": "Não citado"})
    tasks = tasks[:10]

    decisions = [c for c in chunks if any(k in c.lower() for k in ["decid", "aprov", "combin", "agreed"])]
    risks = [c for c in chunks if any(k in c.lower() for k in ["risco", "atras", "bloque", "problema", "reclam"])]

    return {
        "ata_formal": (
            f"Ata da reunião: {title}\n"
            f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}\n"
            f"Idioma detectado: {language}\n"
            "Participantes detectados: Não identificados automaticamente.\n\n"
            "Temas discutidos:\n- " + "\n- ".join(bullet_chunks[:5]) + "\n\n"
            "Pendências e próximos passos:\n- "
            + ("\n- ".join(bullet_chunks[5:8]) if len(bullet_chunks) > 5 else "Não identificados.")
        ),
        "resumo_executivo": " ".join(bullet_chunks[:5]) or "Resumo não disponível.",
        "tarefas": tasks or [{"task": "Sem tarefas claras detectadas", "owner": "—", "deadline": "—"}],
        "decisoes": decisions[:8] or ["Sem decisões explícitas detectadas."],
        "riscos": risks[:8] or ["Nenhum risco explícito detectado."],
        "speakers": ["Participante 1", "Participante 2"],
    }


def structure_meeting(title: str, transcript: str, language: str) -> dict:
    api_key = os.getenv("OPENAI_API_KEY", "")
    model = _model_structure()
    if not api_key:
        return fallback_structuring(title, transcript, language)

    client = OpenAI(api_key=api_key)
    prompt = f"""{ORTOGRAFIA_PT}
És um assistente de reuniões. Analise a transcrição e devolva JSON válido.
Todos os textos no JSON devem seguir a ortografia em português do Brasil (com acentos).

Formato esperado:
{{
  "ata_formal": "texto",
  "resumo_executivo": "texto objetivo",
  "tarefas": [{{"task":"", "owner":"", "deadline":""}}],
  "decisoes": ["..."],
  "riscos": ["..."],
  "speakers": ["..."]
}}

Título: {title}
Idioma detectado: {language}
Transcrição:
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


def _fallback_chat_answer(question: str, meeting: dict) -> str:
    if "prazo" in question.lower():
        return "Prazos citados: " + ", ".join([str(t.get("deadline", "—")) for t in meeting["tarefas"]])
    if "respons" in question.lower():
        return "Responsáveis: " + ", ".join([str(t.get("owner", "—")) for t in meeting["tarefas"]])
    return f"Com base no resumo da reunião: {meeting['resumo_executivo']}"


def _default_suggested_questions() -> list[str]:
    return [
        "Quais foram as principais decisões e quem ficou responsável por cada entrega?",
        "Existem riscos, bloqueios ou dependências citadas na conversa?",
        "Resuma em cinco linhas o que o cliente ou a liderança precisa saber.",
    ]


def chat_with_meeting(question: str, meeting: dict) -> dict:
    """Resposta estilo assistente: resposta + perguntas de continuação sugeridas."""
    api_key = os.getenv("OPENAI_API_KEY", "")
    model = _model_chat()
    context = (
        f"Título: {meeting['title']}\n"
        f"Resumo executivo: {meeting['resumo_executivo']}\n"
        f"Ata (trecho): {str(meeting.get('ata_formal', ''))[:4000]}\n"
        f"Transcrição: {meeting['transcript_clean']}\n"
        f"Tarefas: {meeting['tarefas']}\n"
        f"Decisões: {meeting['decisoes']}\n"
        f"Riscos: {meeting['riscos']}"
    )
    if not api_key:
        return {
            "answer": _fallback_chat_answer(question, meeting),
            "suggested_questions": _default_suggested_questions(),
        }

    client = OpenAI(api_key=api_key)
    system = (
        f"{ORTOGRAFIA_PT} "
        "Você é um assistente de análise de reuniões, no estilo NotebookLM: claro, profissional e útil. "
        "Use apenas o contexto fornecido. Se não houver informação suficiente, diga isso explicitamente. "
        "Você pode usar markdown leve (títulos ##, listas) quando ajudar a leitura. "
        "O campo answer e as suggested_questions devem estar em português do Brasil, com acentos corretos. "
        "O JSON de resposta deve conter 3 sugestões curtas de acompanhamento (perguntas de follow-up) em português."
    )
    user = (
        f"Contexto da reunião:\n{context}\n\nPergunta do usuário:\n{question}\n\n"
        "Responde em JSON válido, exatamente no formato: "
        '{"answer": "texto em markdown", "suggested_questions": ["pergunta1", "pergunta2", "pergunta3"]}'
    )
    try:
        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "system", "content": system}, {"role": "user", "content": user}],
            temperature=0.3,
            response_format={"type": "json_object"},
        )
        raw = response.choices[0].message.content or "{}"
        parsed = json.loads(raw)
        answer = (parsed.get("answer") or "").strip() or "Sem resposta."
        suggested = parsed.get("suggested_questions") or []
        if not isinstance(suggested, list):
            suggested = []
        suggested = [str(s).strip() for s in suggested if str(s).strip()][:3]
        defaults = _default_suggested_questions()
        for d in defaults:
            if len(suggested) >= 3:
                break
            if d not in suggested:
                suggested.append(d)
        return {"answer": answer, "suggested_questions": suggested[:3]}
    except Exception:
        return {
            "answer": "Não foi possível gerar a resposta com a IA de momento. Tenta novamente dentro de instantes.",
            "suggested_questions": _default_suggested_questions(),
        }
