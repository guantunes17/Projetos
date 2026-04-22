# MeetFlow AI MVP (Privado)

MVP funcional para transformar reunioes em:

- Ata Formal
- Resumo Executivo
- Tarefas acionaveis
- Decisoes consolidadas
- Riscos detectados
- Chat com a reuniao

## Stack

- `backend`: FastAPI + SQLite local + Whisper + LLM (OpenAI opcional)
- `frontend`: Next.js + React + Tailwind CSS
- `desktop-ready`: estrutura preparada para wrapper Electron/Tauri

## Fluxo principal

1. Login local simples (usuario bootstrap)
2. Nova reuniao por texto/manual ou upload de audio/video
3. Pipeline de IA com status visual:
   - Transcrevendo
   - Analisando
   - Gerando ata
   - Finalizando insights
4. Visualizacao em abas e chat contextual
5. Exportacao em PDF, DOCX e Markdown

## Credenciais bootstrap

- email: `admin@meetflow.local`
- senha: `admin123`

## Como rodar (Docker)

Na pasta `novo_web`:

```bash
docker compose up --build
```

Frontend: `http://localhost:3000`  
Backend: `http://localhost:8000`

## Variaveis de ambiente (backend)

Use `backend/.env.example` como base:

- `APP_SECRET_KEY`
- `OPENAI_API_KEY` (opcional)
- `MEETFLOW_LLM_MODEL` (default `gpt-4o-mini`)
- `WHISPER_MODEL` (default `base`)

## Endpoint principais

- `POST /api/auth/login`
- `GET /api/dashboard`
- `GET /api/meetings`
- `POST /api/meetings/process-text`
- `POST /api/meetings/process-upload`
- `POST /api/meetings/{id}/chat`
- `GET /api/meetings/{id}/export/{pdf|docx|md}`
