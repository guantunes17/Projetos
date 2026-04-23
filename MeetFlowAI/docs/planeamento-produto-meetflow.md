# Planeamento de produto — MeetFlow AI

Use este ficheiro para **registar escolhas** sobre IA, UX e alcance do MVP. Quando estiver preenchido, serve de **contrato leve** para o que será implementado.

**Como usar**

1. Para cada pergunta, marca **uma letra** (A–E) ou escreve em **Outro / notas**.
2. Na secção final, define **prioridades** e **fora de alcance** para a próxima entrega.
3. Opcional: data e responsável na última linha.

---

## Registo rápido (opcional)

| Campo | Valor |
|--------|--------|
| Data | |
| Participantes | |
| Versão do documento | 1 |

---

## 1) Caso de uso principal do MeetFlow

Marca **uma** opção (ou E + detalhe).

| Opção | Descrição | Escolha |
|-------|-----------|---------|
| **A** | Reuniões 1:1 ou internas | [ x] |
| **B** | Stand-ups / sincronizações rápidas | [ ] |
| **C** | Calls com cliente / stakeholders externos | [x ] |
| **D** | Entrevistas ou sessões longas | [ ] |
| **E** | Outro | [ ] |

**Outro / notas:**

---

## 2) Idioma da interface e das saídas da IA

| Opção | Descrição | Escolha |
|-------|-----------|---------|
| **A** | Português europeu (pt-PT) em todo o lado | [ ] |
| **B** | Português do Brasil (pt-BR) em todo o lado | [ x] |
| **C** | UI fixa num idioma; **por reunião** escolho idioma da geração | [ ] |
| **D** | Suporte a **inglês** nas saídas ou na UI | [x ] |
| **E** | Outro | [ ] |

**Outro / notas:**

---

## 3) Fidelidade ao texto original (transcrição)

| Opção | Descrição | Escolha |
|-------|-----------|---------|
| **A** | Só factos sustentados pela transcrição (conservador) | [ ] |
| **B** | Equilíbrio: resume e infere só quando for óbvio | [x ] |
| **C** | Mais interpretativo quando áudio/texto forem fracos | [ ] |
| **D** | Ainda não sei — pedir recomendação à equipa | [ ] |
| **E** | Outro | [ ] |

**Outro / notas:**

---

## 4) Streaming das respostas do chat

| Opção | Descrição | Escolha |
|-------|-----------|---------|
| **A** | Sim, prioridade (must-have) | [x ] |
| **B** | Sim, nice-to-have depois do core | [ ] |
| **C** | Não — resposta completa de uma vez | [ ] |
| **D** | Indiferente | [ ] |
| **E** | Outro | [ ] |

**Outro / notas:**

---

## 5) Citações / provas (ligar respostas ao texto da reunião)

| Opção | Descrição | Escolha |
|-------|-----------|---------|
| **A** | Sim — trechos da transcrição por afirmação relevante | [ ] |
| **B** | Sim — e timestamps quando existirem | [x ] |
| **C** | Não nesta fase | [ ] |
| **D** | Só no chat, não na ata automática | [ ] |
| **E** | Outro | [ ] |

**Outro / notas:**

---

## 6) Pré-processamento da transcrição (antes de estruturar / chat)

| Opção | Descrição | Escolha |
|-------|-----------|---------|
| **A** | Não — texto tal como vem | [ ] |
| **B** | Sim — normalização leve (espaços, pontuação) | [ ] |
| **C** | Sim — correção ortográfica assistida em PT | [ ] |
| **D** | Sim — normalização + ortografia | [x ] |
| **E** | Outro | [ ] |

**Outro / notas:**

---

## 7) Modelos e custo (vários modelos por tarefa)

| Opção | Descrição | Escolha |
|-------|-----------|---------|
| **A** | Um modelo só para tudo | [ ] |
| **B** | Modelo mais barato para extração + modelo mais forte para chat | [ x] |
| **C** | Decidir por orçamento mensal — pedir recomendação | [ ] |
| **D** | Volumes de uso ainda desconhecidos | [ ] |
| **E** | Outro | [ ] |

**Outro / notas:**

*(Esclarecimento — **pergunta 7**)* A ideia é **controlar custo da API OpenAI**: chamadas ao modelo **por token** são caras se usares sempre o modelo mais capaz para tudo. Duas estratégias comuns:

- **B — Dois modelos:** um **mais barato** (ex.: `gpt-4o-mini`) só para extrair JSON da ata (tarefas, decisões, riscos) e um **mais forte** só para o chat longo onde precisas de raciocínio melhor.
- **C — Orçamento:** defines um teto mensal em euros/dólares e ajustas modelo + limites de tokens em função disso.

Se quiseres **recomendação simples para o MVP:** marca **B** (extracção barata + chat mais capaz), salvo orçamento muito apertado — aí **C** com limites por utilizador.

---

## 8) Memória entre várias reuniões

| Opção | Descrição | Escolha |
|-------|-----------|---------|
| **A** | Não — cada reunião isolada | [ ] |
| **B** | Talvez mais tarde — não é MVP | [ ] |
| **C** | Sim — roadmap (ex.: último mês com cliente X) | [ ] |
| **D** | Sim — já no MVP | [x ] |
| **E** | Outro | [ ] |

**Outro / notas:**

---

## 9) Onde “vive” o assistente no produto

| Opção | Descrição | Escolha |
|-------|-----------|---------|
| **A** | Só na página da reunião (detalhe) | [ ] |
| **B** | Só no assistente / chat dedicado (`/chat`) | [ ] |
| **C** | Ambos — mesmo contexto / fluxo coerente | [ ] |
| **D** | Painel lateral global em todo o workspace | [x ] |
| **E** | Outro | [ ] |

**Outro / notas:**

---

## 10) Estado vazio (sem reuniões)

| Opção | Descrição | Escolha |
|-------|-----------|---------|
| **A** | Onboarding — exemplo, modelo de texto, dicas | [x ] |
| **B** | Mínimo — botão “nova reunião” e pouco texto | [ ] |
| **C** | Vídeo ou tour — mais tarde | [x ] |
| **D** | Duas variantes (interno vs. cliente) | [ ] |
| **E** | Outro | [ ] |

**Outro / notas:**

---

## 11) Acessibilidade

| Opção | Descrição | Escolha |
|-------|-----------|---------|
| **A** | Sem requisito formal por agora | [ ] |
| **B** | Boas práticas (contraste, tamanhos) sem auditoria | [x ] |
| **C** | Alvo WCAG (indicar nível: A / AA) | [ ] |
| **D** | Uso interno — baixa prioridade | [ ] |
| **E** | Outro | [ ] |

**Outro / notas:**

---

## 12) Mobile / tablet

| Opção | Descrição | Escolha |
|-------|-----------|---------|
| **A** | Desktop-first por agora | [ ] |
| **B** | Mobile importante nas próximas iterações | [ ] |
| **C** | Tablet / landscape para apresentações | [ ] |
| **D** | Web responsiva básica chega | [ ] |
| **E** | Outro | [x ] |

**Outro / notas:**
Desktop first mas com implementação para Tablet, web responsive e mobile.
---

## 13) Métrica de sucesso do produto

| Opção | Descrição | Escolha |
|-------|-----------|---------|
| **A** | Tempo até ata/resumo útil | [ ] |
| **B** | Menos edições manuais após a IA | [ ] |
| **C** | Satisfação explícita (ex.: 1–5 após cada reunião) | [ ] |
| **D** | Combinação tempo + edições | [x ] |
| **E** | Outro | [ ] |

**Outro / notas:**

*(Esclarecimento — **pergunta 13**)* “Métrica de sucesso” = **como sabes que o MeetFlow vale a pena**:

- **A — Tempo:** quantos **minutos** desde colar transcrição até ter ata/resumo **aceitável** (sem contar edições gigantes).
- **B — Edição:** quantas **alterações manuais** fazes por reunião depois da IA (menos = melhor).
- **C — Satisfação:** **feedback explícito** após cada reunião (estrelas ou NPS curto).
- **D — Os dois primeiros em conjunto** (tempo + edições).

Se ainda não medires nada, começa por **D** ou só **B** — são mais fáceis de perceber num MVP interno.

---

## 14) Visibilidade “gerado por IA / rascunho”

| Opção | Descrição | Escolha |
|-------|-----------|---------|
| **A** | Badge na UI nos conteúdos gerados | [x ] |
| **B** | Badge na UI + nos exports (PDF, DOCX, MD) | [x ] |
| **C** | Só na UI, não nos ficheiros | [ ] |
| **D** | Não destacar — output “limpo” | [ ] |
| **E** | Outro | [ ] |

**Outro / notas:**

---

## Decisões de implementação (derivadas das respostas)

### Síntese consolidada (leitura do preenchimento)

| # | Tema | Leitura proposta |
|---|------|------------------|
| 1 | Caso de uso | Estão **A e C** marcados → foco **duplo:** reuniões internas/1:1 **e** calls com cliente/stakeholders. |
| 2 | Idioma | **B e D** → UI em **pt-BR** + suporte a **inglês** (saídas e/ou UI — definir escopo). O código atual está em **pt-PT**; migrar copy + `lang` e, se necessário, i18n. |
| 3 | Fidelidade | **B** — equilíbrio entre ancoragem na transcrição e síntese útil. |
| 4 | Chat | **A** — **streaming** é prioridade (SSE ou equivalente). |
| 5 | Provas | **B** — citações **com timestamps** quando existirem (depende de pipeline de áudio/tempo). |
| 6 | Pré-processamento | **D** — normalização + ortografia antes da IA. |
| 7 | Modelos | **B** — modelo mais barato para extração (ata/JSON) + modelo mais forte para o chat. |
| 8 | Memória | **D** — **memória entre reuniões já no MVP** (âmbito grande: pesquisa + dados). |
| 9 | Assistente | **D** — **painel lateral global** (além do detalhe/`/chat`). |
| 10 | Vazio | **A e C** → onboarding com exemplo/modelo **agora**; **vídeo/tour** explicitamente **depois**. |
| 11 | A11y | **B** — boas práticas sem auditoria formal. |
| 12 | Device | **E** — desktop first + **tablet, responsivo e mobile** (roadmap por fases dentro do mesmo objetivo). |
| 13 | Métrica | **D** — combinar **tempo até resumo útil** + **menos edições manuais** por reunião. |
| 14 | IA visível | **B** (inclui badge na UI + exports) — **A** redundante se **B** estiver marcado. |

### Incluir no próximo ciclo (MVP / Sprint) — proposta derivada

| # | Funcionalidade ou alteração | Notas |
|---|-----------------------------|-------|
| 1 | **Internacionalização:** pt-BR na UI + modo EN (strings, `lang`, revisão de copy) | Alinha com Q2; impacto transversal no frontend. |
| 2 | **Pipeline de texto:** normalização + correção ortográfica antes de estruturar/chat | Q6; backend + eventual serviço leve. |
| 3 | **Streaming do chat** | Q4; FastAPI streaming + UI consumindo eventos. |
| 4 | **Citações com referência temporal** | Q5; fase 1: excertos na transcrição; timestamps quando o áudio/transcrição os tiverem. |
| 5 | **Badges “gerado por IA”** na UI + rodapé nos exports | Q14. |
| 6 | **Onboarding** no estado vazio (exemplo + dicas); vídeo/tour em sprint posterior | Q10. |
| 7 | **Layout responsivo** (tablet → mobile) mantendo desktop first | Q12. |
| 8 | **Painel lateral global** do assistente | Q9; maior trabalho de UI/estado (partilhar contexto com detalhe da reunião). |
| 9 | **Memória cross-reunião** (MVP mínimo) | Q8; definir: pesquisa por título/tags, ou embedding leve só nas tuas reuniões. |

**Ordem sugerida de execução (dependências):** (2) pré-processamento → (1) i18n copy base → (3) streaming → (4) citações → (5) badges/exports → (7) responsive → (6) onboarding → (8) painel lateral → (9) memória (último por complexidade).

### Fora de alcance (explicitamente não agora)

| Item | Motivo ou “revisitar quando” |
|------|-----------------------------|
| Auditoria WCAG formal (nível A/AA) | Q11 foi **B** — revisitar se houver clientes externos ou requisito legal. |
| Vídeo / tour guiado completo | Q10 — depois do onboarding textual. |

### Riscos ou dependências

- **Timestamps:** sem segmentação temporal na transcrição ou no áudio, “citações com tempo” ficam limitadas — pode ser necessário evoluir o fluxo de upload (Whisper com segmentos) ou aceitar timestamps aproximados por parágrafo.
- **Memória entre reuniões + painel lateral + streaming:** maior superfície de estado e custo de API; convém **modelos por tarefa (Q7 → B)** para não disparar custos.
- **pt-BR vs código atual pt-PT:** decisão única de locale evita manter duas grafias misturadas.

### Próximo passo único (uma frase)

Implementar **pré-processamento da transcrição** (normalização + ortografia) e variáveis de modelo **dual** (`MEETFLOW_*`) alinhadas à **Q7 — B**, em seguida seguir a ordem da tabela MVP acima.

---

**Última atualização:**  
**Assinatura / responsável:**
