"use client";

import { useRouter, useSearchParams } from "next/navigation";
import { Suspense, useCallback, useEffect, useMemo, useState } from "react";
import { ChevronRight, MessageCircle, Sparkles } from "lucide-react";
import { motion, useReducedMotion } from "framer-motion";

import { MeetingChat } from "@/components/meetflow/meeting-chat";
import { useMeetFlow } from "@/components/meetflow/app-context";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { fadeInUp, MOTION } from "@/lib/motion";
import { cn } from "@/lib/utils";

function ChatStudioContent() {
  const { meetings, token, isLoading } = useMeetFlow();
  const searchParams = useSearchParams();
  const router = useRouter();
  const paramId = searchParams.get("meeting");
  const [selectedId, setSelectedId] = useState(null);
  const [q, setQ] = useState("");
  const reduceMotion = useReducedMotion();

  useEffect(() => {
    if (paramId) {
      setSelectedId(paramId);
    }
  }, [paramId]);

  const selected = useMemo(() => meetings.find((m) => String(m.id) === String(selectedId)), [meetings, selectedId]);

  const filtered = useMemo(
    () => meetings.filter((m) => m.title.toLowerCase().includes(q.toLowerCase())),
    [meetings, q],
  );

  const selectMeeting = useCallback(
    (id) => {
      setSelectedId(String(id));
      router.push(`/chat?meeting=${id}`);
    },
    [router],
  );

  return (
    <div className="space-y-4 xl:space-y-6">
      <motion.div {...fadeInUp(0, reduceMotion, { y: 12, baseDelay: 0.04, duration: MOTION.duration.base })}>
        <Card className="border-blue-500/20 bg-slate-950/70">
          <CardContent className="flex flex-wrap items-center justify-between gap-3 py-4">
            <div className="flex items-center gap-2 text-sm text-slate-300">
              <Sparkles className="h-4 w-4 text-lime-300" />
              Assistente em modo de contexto: selecione uma reunião para respostas ancoradas.
            </div>
            <Badge variant="info">{meetings.length} reunião(ões) disponível(is)</Badge>
          </CardContent>
        </Card>
      </motion.div>
      <div className="grid gap-6 xl:gap-8 xl:grid-cols-[320px_1fr]">
        <motion.div {...fadeInUp(1, reduceMotion, { y: 12, baseDelay: 0.04, duration: MOTION.duration.base })}>
          <Card className="h-fit max-h-[calc(100vh-8rem)] xl:max-h-[calc(100vh-10rem)]">
            <CardHeader>
              <div className="flex items-center gap-2">
                <MessageCircle className="h-4 w-4 text-blue-300" />
                <CardTitle className="text-base">Contexto</CardTitle>
              </div>
              <CardDescription>
                Escolha uma reunião processada. O assistente lê apenas o que já foi analisado.
              </CardDescription>
              <Input
                value={q}
                onChange={(e) => setQ(e.target.value)}
                placeholder="Filtrar reuniões…"
                className="text-sm"
              />
            </CardHeader>
            <CardContent className="max-h-72 space-y-1 overflow-y-auto p-2 pt-0">
              {isLoading && meetings.length === 0 ? (
                <div className="space-y-2 px-2 py-1">
                  {[0, 1, 2].map((i) => (
                    <div
                      key={i}
                      className="h-10 animate-pulse rounded-lg border border-slate-800 bg-slate-900/70"
                    />
                  ))}
                </div>
              ) : null}
              {!isLoading && meetings.length === 0 ? (
                <p className="px-2 text-sm text-slate-400">
                  Ainda não há reuniões. Processa áudio ou texto primeiro.
                </p>
              ) : null}
              {!isLoading && meetings.length > 0 && filtered.length === 0 ? (
                <p className="rounded-lg border border-dashed border-slate-700 px-3 py-2 text-xs text-slate-400">
                  Nenhuma reunião encontrada para esse filtro.
                </p>
              ) : null}
              {filtered.map((m) => (
                <motion.button
                  key={m.id}
                  type="button"
                  onClick={() => selectMeeting(m.id)}
                  initial={reduceMotion ? false : { opacity: 0, x: -8 }}
                  animate={{ opacity: 1, x: 0 }}
                  transition={reduceMotion ? { duration: 0 } : { duration: MOTION.duration.quick + 0.02 }}
                  className={cn(
                    "flex w-full items-center justify-between gap-2 rounded-xl px-3 py-2.5 text-left text-sm transition",
                    String(m.id) === String(selectedId)
                      ? "bg-blue-500/25 text-white ring-1 ring-blue-400/50"
                      : "text-slate-300 hover:bg-slate-800/80",
                  )}
                >
                  <span className="truncate font-medium">{m.title}</span>
                  <ChevronRight className="h-4 w-4 shrink-0 text-slate-500" />
                </motion.button>
              ))}
            </CardContent>
          </Card>
        </motion.div>

        <div className="min-w-0">
          {selectedId && selected ? (
            <MeetingChat meetingId={selected.id} token={token} meetingTitle={selected.title} />
          ) : (
            <Card className="min-h-[420px] border-dashed xl:min-h-[520px]">
              <CardContent className="flex flex-col items-center justify-center gap-4 py-16 text-center xl:py-24">
                <Badge variant="info" className="w-fit">
                  Assistente
                </Badge>
                <div>
                  <p className="text-lg font-medium text-slate-200">Escolha uma reunião</p>
                  <p className="mt-1 max-w-md text-sm text-slate-500">
                    O chat está sempre disponível no menu <strong>Assistente</strong> — não é preciso estar na página da
                    reunião. Escolha um contexto na lista ou crie conteúdo novo.
                  </p>
                </div>
                <div className="flex flex-wrap justify-center gap-2">
                  <Button type="button" onClick={() => router.push("/meetings/new")}>
                    Nova reunião
                  </Button>
                  <Button type="button" variant="secondary" onClick={() => router.push("/meetings")}>
                    Ver histórico
                  </Button>
                </div>
              </CardContent>
            </Card>
          )}
        </div>
      </div>
    </div>
  );
}

export default function ChatStudioPage() {
  return (
    <Suspense
      fallback={
        <div className="flex min-h-[40vh] items-center justify-center text-sm text-slate-500">A abrir o assistente…</div>
      }
    >
      <ChatStudioContent />
    </Suspense>
  );
}
