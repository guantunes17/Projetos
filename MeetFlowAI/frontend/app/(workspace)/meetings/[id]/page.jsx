"use client";

import { useEffect, useState } from "react";
import { useParams, useRouter } from "next/navigation";
import { MessageCircle, Trash2 } from "lucide-react";

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
        "Eliminar esta reunião e todo o conteúdo gerado? Esta ação não pode ser anulada.",
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
      <Card>
        <CardHeader>
          <CardTitle>A carregar a reunião…</CardTitle>
        </CardHeader>
      </Card>
    );
  }

  return (
    <div className="space-y-6">
      <Card>
        <CardHeader>
          <CardTitle>{meeting.title}</CardTitle>
          <CardDescription>Idioma detetado: {meeting.detected_language}</CardDescription>
        </CardHeader>
        <CardContent className="flex flex-wrap gap-2">
          <Button variant="default" onClick={() => router.push(`/chat?meeting=${meeting.id}`)}>
            <MessageCircle size={16} className="mr-1" />
            Assistente (ecrã completo)
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
          <Button variant="secondary" onClick={removeMeeting} disabled={deleting} className="ml-auto text-rose-300">
            <Trash2 size={16} className="mr-1" />
            {deleting ? "A eliminar…" : "Eliminar reunião"}
          </Button>
        </CardContent>
      </Card>

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
