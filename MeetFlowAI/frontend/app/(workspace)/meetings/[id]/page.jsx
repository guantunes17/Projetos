"use client";

import { useEffect, useState } from "react";
import { useParams, useRouter } from "next/navigation";
import {
  AlertTriangle,
  CalendarClock,
  CheckCircle2,
  ClipboardList,
  MessageCircle,
  Trash2,
  User,
} from "lucide-react";
import { motion, useReducedMotion } from "framer-motion";
import ReactMarkdown from "react-markdown";

import { MeetingChat } from "@/components/meetflow/meeting-chat";
import { useMeetFlow } from "@/components/meetflow/app-context";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { api } from "@/lib/api";
import { fadeInUp, MOTION } from "@/lib/motion";

export default function MeetingDetailsPage() {
  const { id } = useParams();
  const router = useRouter();
  const { token, refreshAll } = useMeetFlow();
  const [meeting, setMeeting] = useState(null);
  const [loadError, setLoadError] = useState("");
  const [deleting, setDeleting] = useState(false);
  const reduceMotion = useReducedMotion();

  useEffect(() => {
    if (!token || !id) return;
    setLoadError("");
    api
      .meetingDetail(id, token)
      .then(setMeeting)
      .catch((e) => setLoadError(e.message));
  }, [token, id]);

  async function exportMeeting(format) {
    const blob = await api.exportMeeting(meeting.id, format, token);
    const ext = format === "md" ? "md" : format;
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `${meeting.title.replace(/\s+/g, "_")}.${ext}`;
    link.click();
    URL.revokeObjectURL(link.href);
  }

  async function removeMeeting() {
    if (!meeting) return;
    if (!window.confirm("Excluir esta reunião e todo o conteúdo gerado? Esta ação não pode ser desfeita.")) {
      return;
    }
    setDeleting(true);
    try {
      await api.deleteMeeting(meeting.id, token);
      await refreshAll();
      router.push("/meetings");
    } catch (e) {
      window.alert(e.message);
    } finally {
      setDeleting(false);
    }
  }

  if (loadError) {
    return (
      <Card>
        <CardHeader>
          <CardTitle>Erro</CardTitle>
          <CardDescription>{loadError}</CardDescription>
        </CardHeader>
        <CardContent>
          <Button variant="secondary" onClick={() => router.push("/meetings")}>
            Voltar
          </Button>
        </CardContent>
      </Card>
    );
  }

  if (!meeting) {
    return (
      <Card className="overflow-hidden">
        <CardHeader>
          <CardTitle>Carregando reunião…</CardTitle>
        </CardHeader>
        <CardContent className="space-y-3">
          <div className="h-9 animate-pulse rounded-lg border border-slate-800 bg-slate-900/70" />
          <div className="h-28 animate-pulse rounded-xl border border-slate-800 bg-slate-900/70" />
          <div className="h-52 animate-pulse rounded-xl border border-slate-800 bg-slate-900/70" />
        </CardContent>
      </Card>
    );
  }

  return (
    <div className="space-y-6 xl:space-y-7">
      <motion.div {...fadeInUp(0, reduceMotion, { y: 12, baseDelay: 0.05, duration: MOTION.duration.base })}>
        <Card className="relative overflow-hidden border-blue-500/20 bg-slate-950/80">
          <div className="pointer-events-none absolute inset-0 bg-[radial-gradient(circle_at_14%_18%,rgba(59,130,246,0.18),transparent_45%),radial-gradient(circle_at_85%_84%,rgba(132,204,22,0.15),transparent_40%)]" />
          <CardHeader>
            <div className="flex items-center justify-between gap-3">
              <div>
                <CardTitle>{meeting.title}</CardTitle>
                <CardDescription>Idioma detectado: {meeting.detected_language}</CardDescription>
              </div>
              <div className="hidden items-center gap-1 rounded-xl border border-slate-700/80 bg-slate-900/70 px-3 py-2 text-xs text-slate-300 sm:flex">
                <CalendarClock className="h-3.5 w-3.5 text-blue-300" />
                Contexto pronto para análise
              </div>
            </div>
          </CardHeader>
          <CardContent className="flex flex-wrap gap-2 xl:gap-3">
            <Button variant="default" onClick={() => router.push(`/chat?meeting=${meeting.id}`)}>
              <MessageCircle size={16} className="mr-1" />
              Assistente (tela cheia)
            </Button>
            <Button variant="secondary" onClick={() => exportMeeting("pdf")}>
              Exportar PDF
            </Button>
            <Button variant="secondary" onClick={() => exportMeeting("docx")}>
              Exportar DOCX
            </Button>
            <Button variant="secondary" onClick={() => exportMeeting("md")}>
              Exportar Markdown
            </Button>
            <Button
              variant="secondary"
              onClick={removeMeeting}
              disabled={deleting}
              className="ml-auto border-rose-500/20 text-rose-300 hover:bg-rose-950/40"
            >
              <Trash2 size={16} className="mr-1" />
              {deleting ? "Excluindo…" : "Excluir reunião"}
            </Button>
          </CardContent>
        </Card>
      </motion.div>

      <motion.div {...fadeInUp(1, reduceMotion, { y: 12, baseDelay: 0.05, duration: MOTION.duration.base })}>
        <Tabs defaultValue="ata">
          <TabsList className="flex w-full flex-wrap xl:h-12 xl:px-1.5">
            <TabsTrigger value="ata">Ata</TabsTrigger>
            <TabsTrigger value="resumo">Resumo</TabsTrigger>
            <TabsTrigger value="tarefas">Tarefas</TabsTrigger>
            <TabsTrigger value="decisoes">Decisões</TabsTrigger>
            <TabsTrigger value="riscos">Riscos</TabsTrigger>
            <TabsTrigger value="chat">Chat com a reunião</TabsTrigger>
          </TabsList>
          <TabsContent value="ata">
            <ProseBlock content={meeting.ata_formal} />
          </TabsContent>
          <TabsContent value="resumo">
            <ProseBlock content={meeting.resumo_executivo} />
          </TabsContent>
          <TabsContent value="tarefas">
            <TaskList tasks={meeting.tarefas} />
          </TabsContent>
          <TabsContent value="decisoes">
            <BulletList items={meeting.decisoes} variant="decision" />
          </TabsContent>
          <TabsContent value="riscos">
            <BulletList items={meeting.riscos} variant="risk" />
          </TabsContent>
          <TabsContent value="chat" className="mt-4">
            <MeetingChat meetingId={meeting.id} token={token} meetingTitle={meeting.title} />
          </TabsContent>
        </Tabs>
      </motion.div>
    </div>
  );
}

function ProseBlock({ content }) {
  if (!content?.trim()) {
    return (
      <Card>
        <CardContent className="pt-5">
          <p className="text-sm text-slate-500">Conteúdo não disponível.</p>
        </CardContent>
      </Card>
    );
  }
  return (
    <Card>
      <CardContent className="pt-5">
        <div className="prose prose-invert prose-sm max-w-none prose-headings:mb-2 prose-headings:mt-4 prose-headings:font-semibold prose-headings:text-slate-100 prose-p:mb-3 prose-p:leading-relaxed prose-p:text-slate-200 prose-strong:text-slate-100 prose-li:text-slate-200 prose-li:leading-relaxed">
          <ReactMarkdown>{content}</ReactMarkdown>
        </div>
      </CardContent>
    </Card>
  );
}

function TaskList({ tasks }) {
  const items = Array.isArray(tasks) ? tasks : [];
  if (items.length === 0) {
    return (
      <Card>
        <CardContent className="pt-5">
          <p className="text-sm text-slate-500">Nenhuma tarefa identificada.</p>
        </CardContent>
      </Card>
    );
  }
  return (
    <Card>
      <CardContent className="space-y-3 pt-5">
        {items.map((task, i) => {
          const hasOwner = task.owner && task.owner !== "—" && task.owner !== "Não identificado";
          const hasDeadline = task.deadline && task.deadline !== "—" && task.deadline !== "Não citado";
          return (
            <div
              key={i}
              className="rounded-xl border border-blue-500/15 bg-blue-950/10 p-4 transition-colors hover:border-blue-500/25"
            >
              <div className="flex items-start gap-3">
                <ClipboardList className="mt-0.5 h-4 w-4 shrink-0 text-blue-300" />
                <p className="text-sm font-medium leading-relaxed text-slate-100">{task.task}</p>
              </div>
              {(hasOwner || hasDeadline) && (
                <div className="mt-2 flex flex-wrap gap-4 pl-7">
                  {hasOwner && (
                    <span className="flex items-center gap-1.5 text-xs text-slate-400">
                      <User className="h-3 w-3 text-lime-400" />
                      {task.owner}
                    </span>
                  )}
                  {hasDeadline && (
                    <span className="flex items-center gap-1.5 text-xs text-slate-400">
                      <CalendarClock className="h-3 w-3 text-amber-400" />
                      {task.deadline}
                    </span>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </CardContent>
    </Card>
  );
}

function BulletList({ items, variant = "decision" }) {
  const list = Array.isArray(items) ? items : [];
  const isRisk = variant === "risk";

  if (list.length === 0) {
    return (
      <Card>
        <CardContent className="pt-5">
          <p className="text-sm text-slate-500">
            {isRisk ? "Nenhum risco identificado." : "Nenhuma decisão identificada."}
          </p>
        </CardContent>
      </Card>
    );
  }

  return (
    <Card>
      <CardContent className="space-y-2 pt-5">
        {list.map((item, i) => (
          <div
            key={i}
            className={`flex items-start gap-3 rounded-xl border px-4 py-3 transition-colors ${
              isRisk
                ? "border-amber-500/15 bg-amber-950/10 hover:border-amber-500/25"
                : "border-lime-500/15 bg-lime-950/10 hover:border-lime-500/25"
            }`}
          >
            {isRisk ? (
              <AlertTriangle className="mt-0.5 h-4 w-4 shrink-0 text-amber-400" />
            ) : (
              <CheckCircle2 className="mt-0.5 h-4 w-4 shrink-0 text-lime-400" />
            )}
            <p className="text-sm leading-relaxed text-slate-200">{item}</p>
          </div>
        ))}
      </CardContent>
    </Card>
  );
}
