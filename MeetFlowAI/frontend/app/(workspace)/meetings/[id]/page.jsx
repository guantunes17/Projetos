"use client";

import { useEffect, useState } from "react";
import { useParams, useRouter } from "next/navigation";
import { CalendarClock, MessageCircle, Trash2 } from "lucide-react";
import { motion, useReducedMotion } from "framer-motion";

import { MeetingChat } from "@/components/meetflow/meeting-chat";
import { useMeetFlow } from "@/components/meetflow/app-context";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { api } from "@/lib/api";

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
    if (
      !window.confirm(
        "Excluir esta reunião e todo o conteúdo gerado? Esta ação não pode ser desfeita.",
      )
    ) {
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

  const fadeInUp = (index = 0) => ({
    initial: reduceMotion ? { opacity: 1, y: 0 } : { opacity: 0, y: 12 },
    animate: { opacity: 1, y: 0 },
    transition: reduceMotion ? { duration: 0 } : { delay: 0.05 * index, duration: 0.3, ease: "easeOut" },
  });

  return (
    <div className="space-y-6">
      <motion.div {...fadeInUp(0)}>
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
        <CardContent className="flex flex-wrap gap-2">
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

      <motion.div {...fadeInUp(1)}>
        <Tabs defaultValue="ata">
        <TabsList className="flex w-full flex-wrap">
          <TabsTrigger value="ata">Ata</TabsTrigger>
          <TabsTrigger value="resumo">Resumo</TabsTrigger>
          <TabsTrigger value="tarefas">Tarefas</TabsTrigger>
          <TabsTrigger value="decisoes">Decisões</TabsTrigger>
          <TabsTrigger value="riscos">Riscos</TabsTrigger>
          <TabsTrigger value="chat">Chat com a reunião</TabsTrigger>
        </TabsList>
        <TabsContent value="ata">
          <ContentBlock content={meeting.ata_formal} />
        </TabsContent>
        <TabsContent value="resumo">
          <ContentBlock content={meeting.resumo_executivo} />
        </TabsContent>
        <TabsContent value="tarefas">
          <ContentBlock content={JSON.stringify(meeting.tarefas, null, 2)} />
        </TabsContent>
        <TabsContent value="decisoes">
          <ContentBlock content={JSON.stringify(meeting.decisoes, null, 2)} />
        </TabsContent>
        <TabsContent value="riscos">
          <ContentBlock content={JSON.stringify(meeting.riscos, null, 2)} />
        </TabsContent>
        <TabsContent value="chat" className="mt-4">
          <MeetingChat meetingId={meeting.id} token={token} meetingTitle={meeting.title} />
        </TabsContent>
        </Tabs>
      </motion.div>
    </div>
  );
}

function ContentBlock({ content }) {
  return (
    <Card>
      <CardContent className="pt-5">
        <pre className="whitespace-pre-wrap text-sm text-slate-200">{content}</pre>
      </CardContent>
    </Card>
  );
}
