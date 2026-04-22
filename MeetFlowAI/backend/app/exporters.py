import io
from datetime import datetime

from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


def build_markdown(meeting: dict) -> bytes:
    content = f"""# {meeting['title']}

## Resumo Executivo
{meeting['resumo_executivo']}

## Ata Formal
{meeting['ata_formal']}

## Tarefas
{meeting['tarefas']}

## Decisoes
{meeting['decisoes']}

## Riscos
{meeting['riscos']}

---
Gerado em {datetime.now().isoformat()}
"""
    return content.encode("utf-8")


def build_docx(meeting: dict) -> bytes:
    doc = Document()
    doc.add_heading(meeting["title"], 0)
    doc.add_heading("Resumo Executivo", level=1)
    doc.add_paragraph(meeting["resumo_executivo"])
    doc.add_heading("Ata Formal", level=1)
    doc.add_paragraph(meeting["ata_formal"])
    doc.add_heading("Tarefas", level=1)
    for task in meeting["tarefas"]:
        doc.add_paragraph(f"- {task.get('task')} | Responsavel: {task.get('owner')} | Prazo: {task.get('deadline')}")

    doc.add_heading("Decisoes", level=1)
    for item in meeting["decisoes"]:
        doc.add_paragraph(f"- {item}")

    doc.add_heading("Riscos", level=1)
    for item in meeting["riscos"]:
        doc.add_paragraph(f"- {item}")

    stream = io.BytesIO()
    doc.save(stream)
    return stream.getvalue()


def build_pdf(meeting: dict) -> bytes:
    stream = io.BytesIO()
    pdf = canvas.Canvas(stream, pagesize=A4)
    width, height = A4
    y = height - 50

    def write_line(text: str, top=15):
        nonlocal y
        if y < 50:
            pdf.showPage()
            y = height - 50
        pdf.drawString(40, y, text[:120])
        y -= top

    write_line(meeting["title"], 20)
    write_line("Resumo Executivo:", 18)
    for line in meeting["resumo_executivo"].split("\n"):
        write_line(line)
    write_line("Ata Formal:", 18)
    for line in meeting["ata_formal"].split("\n"):
        write_line(line)

    pdf.save()
    stream.seek(0)
    return stream.read()
