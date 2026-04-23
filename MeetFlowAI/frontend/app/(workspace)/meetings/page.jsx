"use client";

import Link from "next/link";
import { motion } from "framer-motion";
import { Trash2 } from "lucide-react";
import { useMemo, useState } from "react";

import { useMeetFlow } from "@/components/meetflow/app-context";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { api } from "@/lib/api";

export default function MeetingsPage() {
  const { meetings, token, refreshAll } = useMeetFlow();
  const [query, setQuery] = useState("");
  const [deletingId, setDeletingId] = useState(null);

  async function removeMeeting(e, id) {
    e.preventDefault();
    e.stopPropagation();
    if (!window.confirm("Excluir esta reunião e o conteúdo gerado?")) return;
    setDeletingId(id);
    try {
      await api.deleteMeeting(id, token);
      await refreshAll();
    } catch (err) {
      window.alert(err.message);
    } finally {
      setDeletingId(null);
    }
  }

  const filtered = useMemo(
    () => meetings.filter((item) => item.title.toLowerCase().includes(query.toLowerCase())),
    [meetings, query],
  );

  return (
    <div className="space-y-6">
      <Card>
        <CardHeader>
          <CardTitle>Histórico de reuniões</CardTitle>
          <CardDescription>Pesquise por título e aceda aos detalhes processados.</CardDescription>
        </CardHeader>
        <CardContent>
          <Input placeholder="Pesquisar reunião…" value={query} onChange={(e) => setQuery(e.target.value)} />
        </CardContent>
      </Card>

      <section className="grid gap-3">
        {filtered.map((meeting) => (
          <motion.div key={meeting.id} whileHover={{ y: -2 }} transition={{ duration: 0.16 }}>
            <div className="flex items-stretch gap-2 rounded-2xl border border-slate-800 bg-slate-900/70 p-3 transition hover:border-blue-400">
              <Link href={`/meetings/${meeting.id}`} className="min-w-0 flex-1 px-1">
                <p className="font-semibold">{meeting.title}</p>
                <p className="text-xs text-slate-400">
                  Fonte: {meeting.source_type} · Idioma: {meeting.detected_language}
                </p>
              </Link>
              <Button
                type="button"
                variant="ghost"
                className="shrink-0 text-rose-300 hover:bg-rose-950/50"
                disabled={deletingId === meeting.id}
                title="Eliminar reunião"
                onClick={(e) => removeMeeting(e, meeting.id)}
              >
                <Trash2 size={16} />
              </Button>
            </div>
          </motion.div>
        ))}
        {filtered.length === 0 ? <p className="text-sm text-slate-400">Nenhum resultado para a sua pesquisa.</p> : null}
      </section>
    </div>
  );
}
