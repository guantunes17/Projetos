"use client";

import { motion } from "framer-motion";
import { AudioLines, Mic, Square, Upload } from "lucide-react";
import { useMemo, useRef, useState } from "react";
import { useRouter } from "next/navigation";

import { useMeetFlow } from "@/components/meetflow/app-context";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Progress } from "@/components/ui/progress";
import { Textarea } from "@/components/ui/textarea";
import { api } from "@/lib/api";

const steps = ["A transcrever", "A analisar", "A gerar ata", "A finalizar conclusões"];

function delay(ms, signal) {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) {
      reject(new DOMException("Aborted", "AbortError"));
      return;
    }
    const id = setTimeout(() => resolve(), ms);
    signal?.addEventListener("abort", () => {
      clearTimeout(id);
      reject(new DOMException("Aborted", "AbortError"));
    });
  });
}

export default function NewMeetingPage() {
  const { token, refreshAll } = useMeetFlow();
  const [title, setTitle] = useState("Reunião de estado");
  const [transcript, setTranscript] = useState("");
  const [file, setFile] = useState(null);
  const [currentStep, setCurrentStep] = useState(-1);
  const [isRecording, setIsRecording] = useState(false);
  const [recorder, setRecorder] = useState(null);
  const [error, setError] = useState("");
  const [processing, setProcessing] = useState(false);
  const abortRef = useRef(null);
  const router = useRouter();

  const progress = useMemo(() => (currentStep < 0 ? 0 : ((currentStep + 1) / steps.length) * 100), [currentStep]);
  const currentStepLabel = currentStep >= 0 ? steps[currentStep] : "Aguardando início";

  function cancelPipeline() {
    abortRef.current?.abort();
  }

  async function runPipeline(executor) {
    setError("");
    const ac = new AbortController();
    abortRef.current = ac;
    const { signal } = ac;
    setProcessing(true);
    setCurrentStep(0);
    try {
      await delay(320, signal);
      setCurrentStep(1);
      const created = await executor(signal);
      setCurrentStep(2);
      await delay(220, signal);
      setCurrentStep(3);
      await delay(180, signal);
      await refreshAll();
      router.push(`/meetings/${created.id}`);
    } catch (err) {
      if (err?.name === "AbortError") {
        setError("Processamento interrompido.");
      } else {
        setError(err.message);
      }
    } finally {
      setCurrentStep(-1);
      setProcessing(false);
      abortRef.current = null;
    }
  }

  function processText() {
    if (!title.trim()) {
      setError("Informe o título da reunião.");
      return;
    }
    if (!transcript.trim()) {
      setError("Cole a transcrição (texto) ou use upload / gravação. Não é possível gerar a ata com o campo vazio.");
      return;
    }
    return runPipeline((signal) => api.processText({ title, transcript, source_type: "manual" }, token, signal));
  }

  function processUpload() {
    if (!title.trim()) {
      setError("Informe o título da reunião.");
      return;
    }
    if (!file) {
      setError("Selecione um arquivo de áudio ou vídeo antes de processar.");
      return;
    }
    const form = new FormData();
    form.append("title", title);
    form.append("file", file);
    return runPipeline((signal) => api.processUpload(form, token, signal));
  }

  async function startRecording() {
    const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
    const mediaRecorder = new MediaRecorder(stream);
    const chunks = [];
    mediaRecorder.ondataavailable = (event) => chunks.push(event.data);
    mediaRecorder.onstop = () => {
      const blob = new Blob(chunks, { type: "audio/webm" });
      setFile(new File([blob], `gravacao_${Date.now()}.webm`, { type: "audio/webm" }));
      stream.getTracks().forEach((track) => track.stop());
    };
    mediaRecorder.start();
    setRecorder(mediaRecorder);
    setIsRecording(true);
  }

  function stopRecording() {
    recorder?.stop();
    setRecorder(null);
    setIsRecording(false);
  }

  return (
    <div className="space-y-6 xl:space-y-7">
      <Card className="relative overflow-hidden border-blue-500/25 bg-slate-950/80">
        <div className="pointer-events-none absolute inset-0 bg-[radial-gradient(circle_at_16%_18%,rgba(59,130,246,0.2),transparent_45%),radial-gradient(circle_at_84%_80%,rgba(132,204,22,0.16),transparent_40%)]" />
        <CardHeader className="relative z-10">
          <div className="flex items-center gap-2">
            <AudioLines className="h-4 w-4 text-blue-300" />
            <CardTitle>Nova reunião</CardTitle>
          </div>
          <CardDescription>
            Escolha a fonte de entrada e deixe a IA gerar ata, resumo, tarefas e riscos em poucos minutos.
          </CardDescription>
        </CardHeader>
        <CardContent className="relative z-10 space-y-4">
          <Input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="Título da reunião" />
          <div className="grid gap-4 xl:gap-5 md:grid-cols-2">
            <Textarea
              value={transcript}
              onChange={(e) => setTranscript(e.target.value)}
              className="min-h-52 xl:min-h-72"
              placeholder="Cole aqui a transcrição ou o texto da reunião"
            />
            <label className="rounded-xl border border-dashed border-slate-700 p-4 text-sm text-slate-400 xl:p-5">
              Upload de áudio/vídeo (mp3, m4a, wav, mp4)
              <Input type="file" className="mt-3" onChange={(e) => setFile(e.target.files?.[0] || null)} />
            </label>
          </div>
          <p className="rounded-xl border border-slate-700/80 bg-slate-900/60 px-3 py-2 text-xs text-slate-400">
            A gravação no navegador usa o <strong className="text-slate-300">microfone</strong> do dispositivo (ou
            headset/fone). Não captura o áudio interno do sistema (ex.: som de videochamada) — isso exigiria um app
            desktop ou extensão com captura em loopback.
          </p>
          <div className="flex flex-wrap gap-2">
            <Button onClick={processText} disabled={processing}>
              Processar texto
            </Button>
            <Button onClick={processUpload} variant="success" disabled={processing}>
              <Upload size={15} className="mr-2" />
              Processar arquivo
            </Button>
            {!isRecording ? (
              <Button onClick={startRecording} variant="secondary" disabled={processing}>
                <Mic size={15} className="mr-2" />
                Iniciar gravação
              </Button>
            ) : (
              <Button onClick={stopRecording} variant="secondary">
                Parar gravação
              </Button>
            )}
            {processing ? (
              <Button type="button" variant="secondary" onClick={cancelPipeline}>
                <Square size={14} className="mr-1" />
                Interromper
              </Button>
            ) : null}
          </div>
          {error ? <p className="rounded-lg bg-rose-500/10 px-3 py-2 text-sm text-rose-300">{error}</p> : null}
        </CardContent>
      </Card>

      <Card>
        <CardHeader>
          <CardTitle>Pipeline</CardTitle>
          <CardDescription>Status visual das etapas de processamento em tempo real.</CardDescription>
        </CardHeader>
        <CardContent className="space-y-4">
          <Progress value={progress} />
          <div className="flex items-center justify-between rounded-xl border border-slate-700/70 bg-slate-900/60 px-3 py-2 text-xs">
            <span className="text-slate-400">Etapa atual</span>
            <span className="font-medium text-blue-200">{currentStepLabel}</span>
          </div>
          <div className="flex flex-wrap gap-2">
            {steps.map((step, index) => (
              <motion.div key={step} animate={{ scale: currentStep === index ? 1.03 : 1 }} transition={{ duration: 0.2 }}>
                <Badge variant={currentStep === index ? "info" : currentStep > index ? "success" : "default"}>
                  {step}
                </Badge>
              </motion.div>
            ))}
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
