"use client";

import { useMemo, useState } from "react";

const API = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";

async function apiFetch(path, options = {}, token = "") {
  const res = await fetch(`${API}${path}`, {
    ...options,
    headers: {
      ...(options.headers || {}),
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
    },
  });
  const contentType = res.headers.get("content-type") || "";
  const data = contentType.includes("application/json") ? await res.json() : await res.text();
  if (!res.ok) throw new Error(data.detail || data.error || "Erro na requisicao");
  return data;
}

export default function HomePage() {
  const [email, setEmail] = useState("admin@meetflow.local");
  const [password, setPassword] = useState("admin123");
  const [token, setToken] = useState("");
  const [dashboard, setDashboard] = useState(null);
  const [meetings, setMeetings] = useState([]);
  const [selectedMeeting, setSelectedMeeting] = useState(null);
  const [activeTab, setActiveTab] = useState("ata");
  const [title, setTitle] = useState("Reuniao de Status");
  const [transcript, setTranscript] = useState("");
  const [file, setFile] = useState(null);
  const [status, setStatus] = useState("Pronto");
  const [question, setQuestion] = useState("");
  const [chat, setChat] = useState([]);
  const [isRecording, setIsRecording] = useState(false);
  const [recorder, setRecorder] = useState(null);

  const stats = useMemo(
    () =>
      dashboard || {
        meetings_count: 0,
        tasks_count: 0,
        decisions_count: 0,
      },
    [dashboard],
  );

  async function loadData(authToken) {
    const [dash, list] = await Promise.all([
      apiFetch("/api/dashboard", {}, authToken),
      apiFetch("/api/meetings", {}, authToken),
    ]);
    setDashboard(dash);
    setMeetings(list);
  }

  async function handleLogin(e) {
    e.preventDefault();
    const data = await apiFetch("/api/auth/login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email, password }),
    });
    setToken(data.access_token);
    await loadData(data.access_token);
  }

  async function processText() {
    if (!token) return;
    setStatus("Transcrevendo");
    await new Promise((r) => setTimeout(r, 400));
    setStatus("Analisando");
    const data = await apiFetch(
      "/api/meetings/process-text",
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ title, transcript, source_type: "manual" }),
      },
      token,
    );
    setStatus("Gerando ata");
    await new Promise((r) => setTimeout(r, 300));
    setStatus("Finalizando insights");
    setSelectedMeeting(data);
    await loadData(token);
    setStatus("Concluido");
  }

  async function processUpload() {
    if (!token || !file) return;
    const form = new FormData();
    form.append("title", title);
    form.append("file", file);
    setStatus("Transcrevendo");
    const data = await apiFetch("/api/meetings/process-upload", { method: "POST", body: form }, token);
    setStatus("Analisando");
    await new Promise((r) => setTimeout(r, 300));
    setStatus("Gerando ata");
    await new Promise((r) => setTimeout(r, 300));
    setStatus("Finalizando insights");
    setSelectedMeeting(data);
    await loadData(token);
    setStatus("Concluido");
  }

  async function openMeeting(id) {
    const detail = await apiFetch(`/api/meetings/${id}`, {}, token);
    setSelectedMeeting(detail);
  }

  async function askMeeting() {
    if (!question || !selectedMeeting) return;
    const data = await apiFetch(
      `/api/meetings/${selectedMeeting.id}/chat`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ question }),
      },
      token,
    );
    setChat((prev) => [...prev, { q: question, a: data.answer }]);
    setQuestion("");
  }

  async function startRecording() {
    const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
    const mediaRecorder = new MediaRecorder(stream);
    const chunks = [];
    mediaRecorder.ondataavailable = (e) => chunks.push(e.data);
    mediaRecorder.onstop = () => {
      const blob = new Blob(chunks, { type: "audio/webm" });
      const generatedFile = new File([blob], `gravacao_${Date.now()}.webm`, { type: "audio/webm" });
      setFile(generatedFile);
      stream.getTracks().forEach((track) => track.stop());
    };
    mediaRecorder.start();
    setRecorder(mediaRecorder);
    setIsRecording(true);
  }

  function stopRecording() {
    if (!recorder) return;
    recorder.stop();
    setIsRecording(false);
    setRecorder(null);
  }

  async function exportFile(fmt) {
    if (!selectedMeeting || !token) return;
    const res = await fetch(`${API}/api/meetings/${selectedMeeting.id}/export/${fmt}`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok) return;
    const blob = await res.blob();
    const ext = fmt === "md" ? "md" : fmt;
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `${selectedMeeting.title.replace(/\s+/g, "_")}.${ext}`;
    link.click();
    URL.revokeObjectURL(link.href);
  }

  return (
    <main className="min-h-screen bg-gradient-to-b from-slate-950 to-slate-900 p-6">
      <section className="mx-auto max-w-7xl space-y-6">
        <header className="rounded-2xl border border-slate-800 bg-slate-900/80 p-6">
          <h1 className="text-3xl font-semibold tracking-tight">MeetFlow AI MVP</h1>
          <p className="mt-2 text-slate-300">Transforme reunioes em atas, resumos executivos e tarefas acionaveis.</p>
        </header>

        {!token ? (
          <form onSubmit={handleLogin} className="grid gap-3 rounded-2xl border border-slate-800 bg-slate-900 p-6 md:grid-cols-4">
            <input className="rounded-lg bg-slate-800 p-3" value={email} onChange={(e) => setEmail(e.target.value)} />
            <input
              className="rounded-lg bg-slate-800 p-3"
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
            />
            <button className="rounded-lg bg-indigo-500 px-4 py-3 font-medium">Entrar</button>
          </form>
        ) : (
          <div className="grid gap-6 lg:grid-cols-3">
            <aside className="space-y-4 rounded-2xl border border-slate-800 bg-slate-900 p-5">
              <h2 className="text-xl font-medium">Dashboard</h2>
              <div className="grid gap-3">
                <Card label="Reunioes" value={stats.meetings_count} />
                <Card label="Tarefas geradas" value={stats.tasks_count} />
                <Card label="Decisoes extraidas" value={stats.decisions_count} />
              </div>
              <div className="rounded-xl bg-slate-800 p-3 text-sm">Status pipeline: {status}</div>
              <h3 className="mt-4 font-medium">Historico</h3>
              <div className="max-h-72 space-y-2 overflow-auto">
                {meetings.map((m) => (
                  <button
                    key={m.id}
                    onClick={() => openMeeting(m.id)}
                    className="w-full rounded-lg border border-slate-700 p-2 text-left hover:bg-slate-800"
                  >
                    <p className="text-sm font-medium">{m.title}</p>
                    <p className="text-xs text-slate-400">{m.source_type}</p>
                  </button>
                ))}
              </div>
            </aside>

            <section className="space-y-4 rounded-2xl border border-slate-800 bg-slate-900 p-5 lg:col-span-2">
              <h2 className="text-xl font-medium">Nova Reuniao</h2>
              <input className="w-full rounded-lg bg-slate-800 p-3" value={title} onChange={(e) => setTitle(e.target.value)} />
              <div className="grid gap-3 md:grid-cols-2">
                <textarea
                  placeholder="Cole transcricao/manual aqui..."
                  className="min-h-40 rounded-lg bg-slate-800 p-3"
                  value={transcript}
                  onChange={(e) => setTranscript(e.target.value)}
                />
                <div className="rounded-lg border border-dashed border-slate-700 p-4">
                  <p className="text-sm text-slate-300">Upload de audio/video (mp3, m4a, wav, mp4)</p>
                  <input type="file" className="mt-3 w-full" onChange={(e) => setFile(e.target.files?.[0] || null)} />
                </div>
              </div>
              <div className="flex flex-wrap gap-3">
                <button onClick={processText} className="rounded-lg bg-indigo-500 px-4 py-2">
                  Nova Reuniao (texto/manual)
                </button>
                <button onClick={processUpload} className="rounded-lg bg-emerald-500 px-4 py-2">
                  Nova Reuniao (upload)
                </button>
                {!isRecording ? (
                  <button onClick={startRecording} className="rounded-lg bg-rose-500 px-4 py-2">
                    Iniciar gravacao local
                  </button>
                ) : (
                  <button onClick={stopRecording} className="rounded-lg bg-amber-500 px-4 py-2">
                    Parar gravacao
                  </button>
                )}
              </div>

              {selectedMeeting && (
                <div className="space-y-3 rounded-xl border border-slate-700 p-4">
                  <div className="flex flex-wrap gap-2">
                    {[
                      ["ata", "Ata Formal"],
                      ["resumo", "Resumo Executivo"],
                      ["tarefas", "Tarefas"],
                      ["decisoes", "Decisoes"],
                      ["riscos", "Riscos Detectados"],
                      ["chat", "Chat com a Reuniao"],
                    ].map(([id, label]) => (
                      <button
                        key={id}
                        onClick={() => setActiveTab(id)}
                        className={`rounded-md px-3 py-1 text-sm ${activeTab === id ? "bg-indigo-500" : "bg-slate-800"}`}
                      >
                        {label}
                      </button>
                    ))}
                  </div>

                  {activeTab === "ata" && <pre className="whitespace-pre-wrap rounded-lg bg-slate-950 p-3">{selectedMeeting.ata_formal}</pre>}
                  {activeTab === "resumo" && <pre className="whitespace-pre-wrap rounded-lg bg-slate-950 p-3">{selectedMeeting.resumo_executivo}</pre>}
                  {activeTab === "tarefas" && (
                    <pre className="whitespace-pre-wrap rounded-lg bg-slate-950 p-3">{JSON.stringify(selectedMeeting.tarefas, null, 2)}</pre>
                  )}
                  {activeTab === "decisoes" && (
                    <pre className="whitespace-pre-wrap rounded-lg bg-slate-950 p-3">{JSON.stringify(selectedMeeting.decisoes, null, 2)}</pre>
                  )}
                  {activeTab === "riscos" && (
                    <pre className="whitespace-pre-wrap rounded-lg bg-slate-950 p-3">{JSON.stringify(selectedMeeting.riscos, null, 2)}</pre>
                  )}
                  {activeTab === "chat" && (
                    <div className="space-y-3">
                      <div className="flex gap-2">
                        <input
                          value={question}
                          onChange={(e) => setQuestion(e.target.value)}
                          placeholder='Ex: "quais prazos foram citados?"'
                          className="flex-1 rounded-lg bg-slate-800 p-2"
                        />
                        <button onClick={askMeeting} className="rounded-lg bg-indigo-500 px-4">
                          Perguntar
                        </button>
                      </div>
                      <div className="space-y-2">
                        {chat.map((item, idx) => (
                          <div key={idx} className="rounded-lg bg-slate-800 p-3 text-sm">
                            <p className="font-medium">Q: {item.q}</p>
                            <p className="mt-1 text-slate-300">A: {item.a}</p>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  <div className="flex gap-2">
                    <button onClick={() => exportFile("pdf")} className="rounded-lg bg-slate-700 px-3 py-2">
                      Exportar PDF
                    </button>
                    <button onClick={() => exportFile("docx")} className="rounded-lg bg-slate-700 px-3 py-2">
                      Exportar DOCX
                    </button>
                    <button onClick={() => exportFile("md")} className="rounded-lg bg-slate-700 px-3 py-2">
                      Exportar Markdown
                    </button>
                  </div>
                </div>
              )}
            </section>
          </div>
        )}
      </section>
    </main>
  );
}

function Card({ label, value }) {
  return (
    <div className="rounded-xl border border-slate-700 bg-slate-800 p-4">
      <p className="text-xs uppercase tracking-wide text-slate-400">{label}</p>
      <p className="text-2xl font-semibold">{value}</p>
    </div>
  );
}
