"use client";

import { useEffect, useRef, useState } from "react";
import { Bot, Copy, MessageCircle, Send, Sparkles, Square, User } from "lucide-react";
import ReactMarkdown from "react-markdown";

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { api } from "@/lib/api";
import { cn } from "@/lib/utils";

const DEFAULT_SUGGESTIONS = [
  "Quais decisões foram tomadas e quem responde por cada ação?",
  "Que riscos, bloqueios ou dependências surgem na reunião?",
  "Resume em 5 linhas o que a liderança precisa de saber hoje.",
];

/**
 * @param {object} props
 * @param {string|number} props.meetingId
 * @param {string} props.token
 * @param {string} [props.meetingTitle]
 * @param {string} [props.className]
 */
export function MeetingChat({ meetingId, token, meetingTitle, className }) {
  const [messages, setMessages] = useState([]);
  const [input, setInput] = useState("");
  const [suggested, setSuggested] = useState(DEFAULT_SUGGESTIONS);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const bottomRef = useRef(null);
  const abortRef = useRef(null);

  useEffect(() => {
    setMessages([]);
    setInput("");
    setError("");
    setSuggested(DEFAULT_SUGGESTIONS);
  }, [meetingId]);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, loading]);

  function stop() {
    abortRef.current?.abort();
  }

  async function send(rawText) {
    const text = (rawText ?? input).trim();
    if (!text || loading) return;
    setError("");
    setInput("");
    setMessages((m) => [...m, { role: "user", id: `u-${Date.now()}`, content: text }]);
    setLoading(true);
    abortRef.current = new AbortController();
    const { signal } = abortRef.current;
    try {
      const res = await api.chat(meetingId, text, token, signal);
      const answer = (res && res.answer) || "";
      const sq = (res && res.suggested_questions) || [];
      if (Array.isArray(sq) && sq.length > 0) {
        setSuggested(sq.slice(0, 3));
      } else {
        setSuggested(DEFAULT_SUGGESTIONS);
      }
      setMessages((m) => [...m, { role: "assistant", id: `a-${Date.now()}`, content: answer }]);
    } catch (e) {
      if (e?.name === "AbortError") {
        setMessages((m) => [
          ...m,
          { role: "assistant", id: `a-${Date.now()}`, content: "_Resposta interrompida._" },
        ]);
      } else {
        setError(e?.message || "Falha ao enviar pergunta.");
        setInput(text);
        setMessages((m) => m.slice(0, -1));
      }
    } finally {
      setLoading(false);
      abortRef.current = null;
    }
  }

  return (
    <div
      className={cn(
        "flex h-[min(70vh,680px)] flex-col overflow-hidden rounded-3xl border border-slate-700/60 bg-slate-950/40 shadow-2xl shadow-blue-950/20 ring-1 ring-white/5 backdrop-blur-sm",
        className,
      )}
    >
      <div className="border-b border-slate-800/80 bg-gradient-to-r from-slate-900/90 to-slate-950/90 px-5 py-4">
        <div className="flex items-start gap-3">
          <div className="flex h-10 w-10 items-center justify-center rounded-2xl bg-blue-500/20 text-blue-200">
            <MessageCircle className="h-5 w-5" />
          </div>
          <div>
            <p className="text-sm font-semibold text-slate-100">Assistente da reunião</p>
            {meetingTitle ? (
              <p className="line-clamp-1 text-xs text-slate-400">{meetingTitle}</p>
            ) : null}
            <p className="mt-1 text-xs leading-relaxed text-slate-500">
              Perguntas e respostas com base na transcrição, na ata e no resumo processados. Leitura objetiva, sem
              inventar fatos.
            </p>
          </div>
        </div>
      </div>

      <div className="flex-1 space-y-5 overflow-y-auto px-4 py-5">
        {messages.length === 0 ? (
          <div className="mx-auto max-w-md rounded-2xl border border-dashed border-slate-700/80 bg-slate-900/30 p-5 text-center">
            <Sparkles className="mx-auto mb-2 h-6 w-6 text-amber-200/80" />
            <p className="text-sm text-slate-300">Comece com uma pergunta ou use as sugestões abaixo.</p>
            <p className="mt-2 text-xs text-slate-500">
              Dica: você pode pedir resumos, responsáveis, riscos ou citações do contexto.
            </p>
          </div>
        ) : null}
        {messages.map((msg) => (
          <div
            key={msg.id}
            className={cn("flex gap-2", msg.role === "user" ? "flex-row-reverse" : "flex-row")}
          >
            <div
              className={cn(
                "flex h-8 w-8 shrink-0 items-center justify-center rounded-full",
                msg.role === "user" ? "bg-blue-500/30 text-blue-200" : "bg-lime-500/15 text-lime-200",
              )}
            >
              {msg.role === "user" ? <User className="h-4 w-4" /> : <Bot className="h-4 w-4" />}
            </div>
            <div
              className={cn(
                "max-w-[min(100%,32rem)] rounded-2xl px-4 py-3 text-sm shadow-sm",
                msg.role === "user"
                  ? "bg-blue-600/30 text-slate-50"
                  : "border border-slate-700/60 bg-slate-900/80",
              )}
            >
              {msg.role === "assistant" ? (
                <div className="prose prose-invert prose-sm max-w-none prose-p:mb-2 prose-p:leading-relaxed prose-headings:mb-2 prose-headings:text-slate-100">
                  <ReactMarkdown>{msg.content}</ReactMarkdown>
                </div>
              ) : (
                <p className="whitespace-pre-wrap text-slate-100">{msg.content}</p>
              )}
              {msg.role === "assistant" ? (
                <button
                  type="button"
                  className="mt-2 inline-flex items-center gap-1.5 text-xs text-slate-500 transition hover:text-slate-300"
                  onClick={() => navigator.clipboard.writeText(msg.content)}
                >
                  <Copy className="h-3.5 w-3.5" />
                  Copiar
                </button>
              ) : null}
            </div>
          </div>
        ))}
        {loading ? (
          <div className="flex items-center gap-2 pl-10 text-xs text-slate-500">
            <span className="inline-block h-2 w-2 animate-bounce rounded-full bg-blue-400" />
            <span className="inline-block h-2 w-2 animate-bounce rounded-full bg-blue-400 [animation-delay:0.1s]" />
            <span className="inline-block h-2 w-2 animate-bounce rounded-full bg-blue-400 [animation-delay:0.2s]" />
            <span>Pensando com base no seu conteúdo…</span>
          </div>
        ) : null}
        <div ref={bottomRef} />
      </div>

      {suggested.length > 0 ? (
        <div className="shrink-0 space-y-2 border-t border-slate-800/80 bg-slate-950/40 px-4 py-3">
          <p className="text-[10px] font-medium uppercase tracking-wider text-slate-500">Sugestões</p>
          <div className="flex max-h-20 flex-wrap gap-2 overflow-y-auto">
            {suggested.map((s) => (
              <Button
                key={s}
                type="button"
                variant="ghost"
                size="sm"
                className="h-auto max-w-full rounded-full border border-slate-700/80 bg-slate-900/50 px-3 py-1.5 text-left text-xs text-slate-300 hover:border-blue-500/50 hover:bg-slate-800/80"
                disabled={loading}
                onClick={() => send(s)}
              >
                <Sparkles className="mr-1.5 h-3 w-3 shrink-0 text-lime-300" />
                <span className="whitespace-normal">{s}</span>
              </Button>
            ))}
          </div>
        </div>
      ) : null}

      {error ? <p className="px-4 text-xs text-rose-400">{error}</p> : null}

      <div className="flex flex-wrap items-end gap-2 border-t border-slate-800/80 bg-slate-950/60 p-4">
        <div className="min-w-0 flex-1">
          <Input
            value={input}
            onChange={(e) => setInput(e.target.value)}
            placeholder="Escreva sua pergunta… (Enter envia; Shift+Enter — no MVP, só uma linha)"
            className="min-h-11"
            disabled={loading}
            onKeyDown={(e) => {
              if (e.key === "Enter" && !e.shiftKey) {
                e.preventDefault();
                send();
              }
            }}
          />
        </div>
        {loading ? (
          <Button type="button" variant="secondary" onClick={stop}>
            <Square size={14} className="mr-1" />
            Parar
          </Button>
        ) : null}
        <Button type="button" onClick={() => send()} disabled={loading} className="shrink-0">
          <Send size={16} className="mr-1" />
          Enviar
        </Button>
      </div>
    </div>
  );
}
