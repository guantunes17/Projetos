"use client";

import { useEffect, useState } from "react";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { useMeetFlow } from "@/components/meetflow/app-context";

export function AuthGate({ children }) {
  const { token, login, register } = useMeetFlow();
  const [mode, setMode] = useState("login");
  const [name, setName] = useState("");
  const [email, setEmail] = useState("admin@meetflow.app");
  const [password, setPassword] = useState("admin123");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [rememberMe, setRememberMe] = useState(false);
  const [error, setError] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [cooldownSeconds, setCooldownSeconds] = useState(0);

  useEffect(() => {
    if (!cooldownSeconds) return;
    const timer = window.setInterval(() => {
      setCooldownSeconds((prev) => (prev > 0 ? prev - 1 : 0));
    }, 1000);
    return () => window.clearInterval(timer);
  }, [cooldownSeconds]);

  if (token) return children;

  async function handleSubmit(event) {
    event.preventDefault();
    if (mode === "login" && cooldownSeconds > 0) {
      return;
    }
    setSubmitting(true);
    setError("");
    try {
      if (mode === "login") {
        await login(email, password, rememberMe);
      } else {
        if (!name.trim()) {
          throw new Error("Informe seu nome.");
        }
        if (password !== confirmPassword) {
          throw new Error("As senhas não conferem.");
        }
        await register({ full_name: name, email, password }, rememberMe);
      }
    } catch (err) {
      if (mode === "login" && err?.status === 429) {
        setCooldownSeconds(Number(err?.retryAfterSeconds || 0));
      }
      setError(err.message);
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <div className="flex min-h-screen items-center justify-center bg-slate-950 p-6">
      <Card className="w-full max-w-xl">
        <CardHeader>
          <CardTitle>MeetFlow AI</CardTitle>
          <CardDescription>
            {mode === "login"
              ? "Faça login para acessar o workspace privado de reuniões."
              : "Crie sua conta para começar a usar o MeetFlow AI."}
          </CardDescription>
        </CardHeader>
        <CardContent>
          <div className="mb-4 grid grid-cols-2 gap-2 rounded-xl border border-slate-800 bg-slate-900/70 p-1">
            <button
              type="button"
              onClick={() => {
                setMode("login");
                setError("");
              }}
              className={`rounded-lg px-3 py-1.5 text-sm transition ${
                mode === "login" ? "bg-blue-500 text-white" : "text-slate-300 hover:bg-slate-800"
              }`}
            >
              Entrar
            </button>
            <button
              type="button"
              onClick={() => {
                setMode("register");
                setError("");
              }}
              className={`rounded-lg px-3 py-1.5 text-sm transition ${
                mode === "register" ? "bg-blue-500 text-white" : "text-slate-300 hover:bg-slate-800"
              }`}
            >
              Criar conta
            </button>
          </div>
          <form className="space-y-3" onSubmit={handleSubmit}>
            {mode === "register" ? <Input value={name} onChange={(e) => setName(e.target.value)} placeholder="Seu nome" /> : null}
            <Input value={email} onChange={(e) => setEmail(e.target.value)} placeholder="E-mail" />
            <Input type="password" value={password} onChange={(e) => setPassword(e.target.value)} placeholder="Senha" />
            {mode === "register" ? (
              <Input
                type="password"
                value={confirmPassword}
                onChange={(e) => setConfirmPassword(e.target.value)}
                placeholder="Confirmar senha"
              />
            ) : null}
            {mode === "register" ? (
              <p className="text-xs text-slate-500">
                A senha deve ter no mínimo 8 caracteres, com maiúscula, minúscula, número e símbolo.
              </p>
            ) : null}
            <label className="flex items-center gap-2 text-sm text-slate-300">
              <input
                type="checkbox"
                checked={rememberMe}
                onChange={(e) => setRememberMe(e.target.checked)}
                className="h-4 w-4 rounded border-slate-700 bg-slate-900 accent-blue-500"
              />
              Manter conectado
            </label>
            {cooldownSeconds > 0 ? (
              <p className="text-sm text-amber-300">
                Muitas tentativas. Tente novamente em {Math.ceil(cooldownSeconds / 60)} min.
              </p>
            ) : null}
            {error ? <p className="text-sm text-rose-300">{error}</p> : null}
            <Button className="w-full" disabled={submitting || (mode === "login" && cooldownSeconds > 0)}>
              {submitting ? "Processando..." : mode === "login" ? "Entrar" : "Criar conta"}
            </Button>
          </form>
        </CardContent>
      </Card>
    </div>
  );
}
