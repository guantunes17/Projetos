"use client";

import { useState } from "react";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { useMeetFlow } from "@/components/meetflow/app-context";

export function AuthGate({ children }) {
  const { token, login } = useMeetFlow();
  const [email, setEmail] = useState("admin@meetflow.local");
  const [password, setPassword] = useState("admin123");
  const [error, setError] = useState("");
  const [submitting, setSubmitting] = useState(false);

  if (token) return children;

  async function handleSubmit(event) {
    event.preventDefault();
    setSubmitting(true);
    setError("");
    try {
      await login(email, password);
    } catch (err) {
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
          <CardDescription>Inicie sessão para aceder ao workspace privado de reuniões.</CardDescription>
        </CardHeader>
        <CardContent>
          <form className="space-y-3" onSubmit={handleSubmit}>
            <Input value={email} onChange={(e) => setEmail(e.target.value)} />
            <Input type="password" value={password} onChange={(e) => setPassword(e.target.value)} />
            {error ? <p className="text-sm text-rose-300">{error}</p> : null}
            <Button className="w-full" disabled={submitting}>
              {submitting ? "Entrando..." : "Entrar"}
            </Button>
          </form>
        </CardContent>
      </Card>
    </div>
  );
}
