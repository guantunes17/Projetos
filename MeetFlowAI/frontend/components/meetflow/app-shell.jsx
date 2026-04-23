"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import { LayoutDashboard, ListChecks, LogOut, MessageCircle, PlusCircle } from "lucide-react";

import { useMeetFlow } from "@/components/meetflow/app-context";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";

const items = [
  { href: "/dashboard", label: "Dashboard", icon: LayoutDashboard },
  { href: "/chat", label: "Assistente", icon: MessageCircle },
  { href: "/meetings/new", label: "Nova reunião", icon: PlusCircle },
  { href: "/meetings", label: "Reuniões", icon: ListChecks },
];

export function AppShell({ children }) {
  const pathname = usePathname();
  const { userName, meetings, logout, syncError, clearSyncError, refreshAll } = useMeetFlow();

  return (
    <div className="min-h-screen bg-gradient-to-b from-slate-950 via-slate-950 to-slate-900 text-slate-100">
      <div className="mx-auto grid min-h-screen max-w-7xl grid-cols-1 gap-6 px-4 py-6 lg:grid-cols-[260px_1fr]">
        <aside className="rounded-3xl border border-slate-800 bg-slate-900/80 p-4 backdrop-blur">
          <div className="mb-6 space-y-2">
            <p className="text-lg font-semibold">MeetFlow AI</p>
            <Badge variant="info">Workspace privado</Badge>
            <p className="text-xs text-slate-400">{userName}</p>
          </div>
          <nav className="space-y-1">
            {items.map((item) => {
              const Icon = item.icon;
              const active = pathname === item.href || pathname.startsWith(`${item.href}/`);
              return (
                <Link
                  key={item.href}
                  href={item.href}
                  className={cn(
                    "flex items-center gap-2 rounded-xl px-3 py-2 text-sm text-slate-300 transition-all duration-200 hover:bg-slate-800 motion-safe:hover:translate-x-0.5",
                    active && "bg-blue-500 text-white hover:bg-blue-500",
                  )}
                >
                  <Icon size={16} />
                  {item.label}
                </Link>
              );
            })}
          </nav>
          <div className="mt-6 rounded-xl border border-slate-800 p-3">
            <p className="text-xs text-slate-400">Reuniões no histórico</p>
            <p className="text-2xl font-semibold">{meetings.length}</p>
          </div>
          <Button className="mt-6 w-full" variant="secondary" onClick={logout}>
            <LogOut size={16} className="mr-2" />
            Sair
          </Button>
        </aside>
        <main className="min-w-0">
          {syncError ? (
            <div className="mb-4 flex flex-wrap items-center justify-between gap-2 rounded-2xl border border-amber-500/30 bg-amber-950/30 px-4 py-2.5 text-xs text-amber-100/90 sm:text-sm">
              <span className="pr-2">
                Sem conexão com o servidor. Seu histórico não é apagado.{" "}
                <button type="button" className="text-blue-300 underline" onClick={() => refreshAll()}>
                  Sincronizar
                </button>{" "}
                ou
                <button type="button" className="ml-1 text-slate-400 underline" onClick={clearSyncError}>
                  ocultar
                </button>
                .
              </span>
            </div>
          ) : null}
          {children}
        </main>
      </div>
    </div>
  );
}
