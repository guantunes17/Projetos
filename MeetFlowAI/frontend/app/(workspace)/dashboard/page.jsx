"use client";

import Link from "next/link";
import { useEffect } from "react";
import { motion } from "framer-motion";
import { Sparkles } from "lucide-react";

import { useMeetFlow } from "@/components/meetflow/app-context";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";

export default function DashboardPage() {
  const { dashboard, meetings, refreshAll, token, isLoading } = useMeetFlow();

  useEffect(() => {
    if (token) refreshAll();
  }, [token]);

  return (
    <div className="space-y-6">
      <Card className="bg-gradient-to-r from-indigo-500/20 to-slate-900/80">
        <CardHeader>
          <Badge variant="info" className="w-fit">
            MeetFlow Performance
          </Badge>
          <CardTitle className="text-2xl">Painel executivo</CardTitle>
          <CardDescription>
            Consulte o volume de reuniões, tarefas acionáveis e decisões registadas.
          </CardDescription>
        </CardHeader>
        <CardContent className="flex flex-wrap gap-3">
          <Link href="/meetings/new">
            <Button>
              <Sparkles size={16} className="mr-2" />
              Nova reunião
            </Button>
          </Link>
          <Link href="/meetings">
            <Button variant="secondary">Ver histórico</Button>
          </Link>
        </CardContent>
      </Card>

      <section className="grid gap-4 md:grid-cols-3">
        {[dashboard.meetings_count, dashboard.tasks_count, dashboard.decisions_count].map((value, idx) => (
          <motion.div
            key={idx}
            initial={{ opacity: 0, y: 8 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: idx * 0.05, duration: 0.2 }}
          >
            <Metric
              title={["Reuniões processadas", "Tarefas extraídas", "Decisões consolidadas"][idx]}
              value={value}
            />
          </motion.div>
        ))}
      </section>

      <Card>
        <CardHeader>
          <CardTitle>Últimas reuniões</CardTitle>
          <CardDescription>
            {isLoading ? "A atualizar dados…" : "Aceda rapidamente às análises recentes."}
          </CardDescription>
        </CardHeader>
        <CardContent className="space-y-2">
          {meetings.slice(0, 6).map((meeting) => (
            <Link
              href={`/meetings/${meeting.id}`}
              key={meeting.id}
              className="block rounded-xl border border-slate-800 bg-slate-900 px-4 py-3 transition hover:border-indigo-400"
            >
              <p className="font-medium">{meeting.title}</p>
              <p className="text-xs text-slate-400">{meeting.source_type}</p>
            </Link>
          ))}
          {meetings.length === 0 ? <p className="text-sm text-slate-400">Ainda sem reuniões processadas.</p> : null}
        </CardContent>
      </Card>
    </div>
  );
}

function Metric({ title, value }) {
  return (
    <Card>
      <CardHeader>
        <CardDescription>{title}</CardDescription>
        <CardTitle className="text-3xl">{value}</CardTitle>
      </CardHeader>
    </Card>
  );
}
