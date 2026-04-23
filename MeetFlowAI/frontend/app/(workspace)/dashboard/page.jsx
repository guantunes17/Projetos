"use client";

import Link from "next/link";
import { useEffect } from "react";
import { motion, useReducedMotion } from "framer-motion";
import { ArrowRight, CalendarDays, CircleCheckBig, Sparkles, WandSparkles } from "lucide-react";

import { useMeetFlow } from "@/components/meetflow/app-context";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { fadeInUp, MOTION } from "@/lib/motion";

export default function DashboardPage() {
  const { dashboard, meetings, refreshAll, token, isLoading } = useMeetFlow();
  const reduceMotion = useReducedMotion();

  useEffect(() => {
    if (token) refreshAll();
  }, [token]);

  const metrics = [
    {
      title: "Reuniões processadas",
      value: dashboard.meetings_count,
      icon: CalendarDays,
      accent: "from-blue-500/25 to-sky-500/20",
      ring: "ring-blue-400/30",
    },
    {
      title: "Tarefas extraídas",
      value: dashboard.tasks_count,
      icon: CircleCheckBig,
      accent: "from-lime-400/20 to-emerald-500/20",
      ring: "ring-lime-300/30",
    },
    {
      title: "Decisões consolidadas",
      value: dashboard.decisions_count,
      icon: WandSparkles,
      accent: "from-sky-500/20 to-lime-400/20",
      ring: "ring-sky-400/30",
    },
  ];

  return (
    <div className="space-y-6 xl:space-y-7">
      <motion.section {...fadeInUp(0, reduceMotion, { y: 18, baseDelay: 0.06, duration: MOTION.duration.slow })}>
        <Card className="relative overflow-hidden border-blue-500/30 bg-slate-950/80">
          <div className="pointer-events-none absolute inset-0 bg-[radial-gradient(circle_at_12%_18%,rgba(59,130,246,0.24),transparent_45%),radial-gradient(circle_at_84%_84%,rgba(132,204,22,0.2),transparent_40%)]" />
          <CardHeader className="relative z-10 space-y-3">
            <Badge variant="info" className="w-fit border-blue-300/40 bg-blue-500/20 text-blue-100">
              MeetFlow Performance
            </Badge>
            <CardTitle className="text-3xl md:text-4xl">Painel executivo</CardTitle>
            <CardDescription className="max-w-2xl text-base text-slate-300">
              Visual premium para acompanhar reuniões, tarefas e decisões com mais ritmo e clareza no fluxo.
            </CardDescription>
          </CardHeader>
          <CardContent className="relative z-10 flex flex-wrap items-center gap-3">
            <Link href="/meetings/new">
              <Button className="group bg-blue-500 text-white hover:bg-blue-400">
                <Sparkles size={16} className="mr-2" />
                Nova reunião
                <ArrowRight
                  size={15}
                  className="ml-2 transition-transform duration-200 group-hover:translate-x-0.5 md:group-hover:translate-x-1"
                />
              </Button>
            </Link>
            <Link href="/meetings">
              <Button variant="secondary">Ver histórico</Button>
            </Link>
          </CardContent>
        </Card>
      </motion.section>

      <section className="grid gap-4 xl:gap-5 md:grid-cols-3">
        {metrics.map((item, idx) => (
          <motion.div key={item.title} {...fadeInUp(idx + 1, reduceMotion, { y: 18, baseDelay: 0.06, duration: MOTION.duration.slow })}>
            <MetricCard item={item} />
          </motion.div>
        ))}
      </section>

      <motion.section {...fadeInUp(4, reduceMotion, { y: 18, baseDelay: 0.06, duration: MOTION.duration.slow })}>
        <Card className="overflow-hidden xl:min-h-[360px]">
          <CardHeader className="border-b border-slate-800/80 bg-slate-900/40">
            <CardTitle>Últimas reuniões</CardTitle>
            <CardDescription>
              {isLoading ? "Atualizando dados…" : "Acesse rapidamente análises e contexto recente."}
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-2 pt-5">
            {meetings.slice(0, 6).map((meeting, index) => (
              <motion.div
                key={meeting.id}
                initial={reduceMotion ? false : { opacity: 0, x: -10 }}
                animate={{ opacity: 1, x: 0 }}
                transition={reduceMotion ? { duration: 0 } : { delay: 0.02 * index, duration: 0.28 }}
              >
                <Link
                  href={`/meetings/${meeting.id}`}
                  className="group block rounded-xl border border-slate-800 bg-slate-900 px-4 py-3 transition hover:border-lime-400/80"
                >
                  <div className="flex items-center justify-between gap-3">
                    <p className="truncate font-medium">{meeting.title}</p>
                    <ArrowRight className="h-4 w-4 shrink-0 text-slate-500 transition group-hover:text-lime-300 md:group-hover:translate-x-1" />
                  </div>
                  <p className="text-xs text-slate-400">{meeting.source_type}</p>
                </Link>
              </motion.div>
            ))}
            {meetings.length === 0 ? (
              <p className="rounded-xl border border-dashed border-slate-700 p-4 text-sm text-slate-400">
                Ainda sem reuniões processadas. Crie a primeira para começar.
              </p>
            ) : null}
          </CardContent>
        </Card>
      </motion.section>
    </div>
  );
}

function MetricCard({ item }) {
  const Icon = item.icon;

  return (
    <Card className={`group border-slate-700/80 bg-gradient-to-br ${item.accent} ring-1 ${item.ring}`}>
      <CardHeader className="space-y-3">
        <div className="flex items-center justify-between">
          <CardDescription className="text-slate-200">{item.title}</CardDescription>
          <span className="rounded-lg border border-slate-700 bg-slate-900/70 p-2 text-slate-300 transition md:group-hover:-translate-y-0.5 md:group-hover:text-white">
            <Icon className="h-4 w-4" />
          </span>
        </div>
        <CardTitle className="text-4xl font-semibold tracking-tight">{item.value}</CardTitle>
      </CardHeader>
    </Card>
  );
}
