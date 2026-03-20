import sys
import types
import os
import uuid
import queue
import threading
import json
from datetime import date, datetime

# ── Mock tkinter so we can import business logic without a display ──────────
_tk = types.ModuleType('tkinter')
_tk.END = 'end'
_tk.NW = 'nw'
_tk.YES = True
_tk.BOTH = 'both'
_tk.X = 'x'
_tk.Y = 'y'
_tk.LEFT = 'left'
_tk.RIGHT = 'right'
_tk.TOP = 'top'
_tk.BOTTOM = 'bottom'
_tk.DISABLED = 'disabled'
_tk.NORMAL = 'normal'
_tk.WORD = 'word'

class _Var:
    def __init__(self, value=''): self._v = value
    def get(self): return self._v
    def set(self, v): self._v = v

_tk.StringVar  = _Var
_tk.BooleanVar = _Var
_tk.IntVar     = _Var

for _a in ['Tk', 'Frame', 'Label', 'Button', 'Text', 'Entry', 'Canvas',
           'Scrollbar', 'Toplevel', 'Spinbox', 'OptionMenu', 'PhotoImage']:
    setattr(_tk, _a, lambda *a, **k: None)

_ttk = types.ModuleType('tkinter.ttk')
for _a in ['Combobox', 'Scrollbar', 'Style', 'Progressbar']:
    setattr(_ttk, _a, lambda *a, **k: None)

_fd             = types.ModuleType('tkinter.filedialog')
_fd.askopenfilename = lambda *a, **k: ''
_fd.askdirectory    = lambda *a, **k: ''

_mb             = types.ModuleType('tkinter.messagebox')
_mb.showinfo    = _mb.showwarning = _mb.showerror = lambda *a, **k: None
_mb.askyesno    = lambda *a, **k: True

sys.modules['tkinter']             = _tk
sys.modules['tkinter.ttk']         = _ttk
sys.modules['tkinter.filedialog']  = _fd
sys.modules['tkinter.messagebox']  = _mb

# ── Flask ────────────────────────────────────────────────────────────────────
from flask import (Flask, render_template, request, redirect, url_for,
                   send_file, abort, jsonify, Response, stream_with_context)
from werkzeug.utils import secure_filename

# ── Business logic ───────────────────────────────────────────────────────────
import central_relatorios as cr

# ── Override data paths to use /data/ ───────────────────────────────────────
DATA_DIR   = os.environ.get('DATA_DIR', '/data')
UPLOAD_DIR = os.path.join(DATA_DIR, 'uploads')

cr.PASTA_PEDIDOS         = os.path.join(DATA_DIR, 'Pedidos')
cr.PASTA_FRETES          = os.path.join(DATA_DIR, 'Fretes')
cr.PASTA_ARMAZENAGEM     = os.path.join(DATA_DIR, 'Armazenagem')
cr.PASTA_ESTOQUE         = os.path.join(DATA_DIR, 'Estoque')
cr.PASTA_PRODUTIVIDADE   = os.path.join(DATA_DIR, 'Produtividade')
cr.PASTA_CAP_OPERACIONAL = os.path.join(DATA_DIR, 'CapacidadeOperacional')
cr.PASTA_RECEBIMENTOS    = os.path.join(DATA_DIR, 'Recebimentos')
cr.PASTA_FINANCEIRO      = os.path.join(DATA_DIR, 'Financeiro')
cr.PASTA_CONSOLIDADOS    = os.path.join(DATA_DIR, 'Consolidados')
cr.ESTADO_PATH           = os.path.join(DATA_DIR, 'state', 'state.json')
cr.DE_PARA_PATH          = os.path.join(DATA_DIR, 'state', 'de_para.json')
cr.DB_ESTOQUE_PATH       = os.path.join(DATA_DIR, 'state', 'estoque_db.json')
cr.DB_FAMILIAS_PATH      = os.path.join(DATA_DIR, 'state', 'familias_db.json')
cr.DB_PRECOS_ARM_PATH    = os.path.join(DATA_DIR, 'state', 'precos_arm.json')

for _d in [
    UPLOAD_DIR, os.path.join(DATA_DIR, 'state'),
    cr.PASTA_PEDIDOS, cr.PASTA_FRETES, cr.PASTA_ARMAZENAGEM,
    cr.PASTA_ESTOQUE, cr.PASTA_PRODUTIVIDADE, cr.PASTA_CAP_OPERACIONAL,
    cr.PASTA_RECEBIMENTOS, cr.PASTA_FINANCEIRO, cr.PASTA_CONSOLIDADOS,
]:
    os.makedirs(_d, exist_ok=True)

# ── App ──────────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'central-relatorios-dev')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB upload limit

# Job registry: {job_id: {'queue': Queue, 'status': str, 'result': any}}
_jobs: dict = {}

# ── Helpers ──────────────────────────────────────────────────────────────────
MESES_PT = ['Janeiro', 'Fevereiro', 'Marco', 'Abril', 'Maio', 'Junho',
            'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
MESES_PT_ACC = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
                'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']


def meses_opcoes():
    hoje = date.today()
    out = []
    for i in range(12):
        mes = hoje.month - 1 - i
        ano = hoje.year
        while mes <= 0:
            mes += 12
            ano -= 1
        out.append((f'{mes:02d}-{ano}', f'{MESES_PT_ACC[mes-1]} de {ano}'))
    return out


def save_upload(file_obj) -> str:
    fname = secure_filename(file_obj.filename)
    dest  = os.path.join(UPLOAD_DIR, f'{uuid.uuid4().hex[:8]}_{fname}')
    file_obj.save(dest)
    return dest


def start_job(runner) -> str:
    """
    runner: callable(log) -> result
    Runs runner in a background thread, streams log via SSE.
    Returns job_id.
    """
    job_id = uuid.uuid4().hex
    q      = queue.Queue()
    _jobs[job_id] = {'queue': q, 'status': 'running', 'result': None}

    def log(msg):
        q.put({'type': 'log', 'msg': msg})

    def target():
        try:
            result = runner(log)
        except Exception as exc:
            import traceback as tb
            result = False
            q.put({'type': 'log',
                   'msg': f'\n Erro inesperado: {exc}\n{tb.format_exc()}\n'})
        _jobs[job_id]['result'] = result
        _jobs[job_id]['status'] = 'done'
        if result and isinstance(result, str) and os.path.isfile(result):
            rel   = os.path.relpath(result, DATA_DIR)
            fname = os.path.basename(result)
            q.put({'type': 'done', 'download_url': f'/download/{rel}',
                   'filename': fname})
        else:
            q.put({'type': 'done', 'download_url': None, 'filename': None,
                   'error': result is False})

    threading.Thread(target=target, daemon=True).start()
    return job_id


def start_db_job(runner) -> str:
    """Like start_job but result is True/False (no file download)."""
    job_id = uuid.uuid4().hex
    q      = queue.Queue()
    _jobs[job_id] = {'queue': q, 'status': 'running', 'result': None}

    def log(msg):
        q.put({'type': 'log', 'msg': msg})

    def target():
        try:
            result, info = runner(log)
        except Exception as exc:
            import traceback as tb
            result, info = False, ''
            q.put({'type': 'log',
                   'msg': f'\n Erro inesperado: {exc}\n{tb.format_exc()}\n'})
        _jobs[job_id]['status'] = 'done'
        q.put({'type': 'done', 'download_url': None, 'filename': None,
               'error': not result,
               'info': info if result else ''})

    threading.Thread(target=target, daemon=True).start()
    return job_id


# ── SSE ──────────────────────────────────────────────────────────────────────
@app.route('/stream/<job_id>')
def stream(job_id):
    if job_id not in _jobs:
        def _err():
            yield f'data: {json.dumps({"type":"error","msg":"Job nao encontrado"})}\n\n'
        return Response(_err(), mimetype='text/event-stream')

    def generate():
        q = _jobs[job_id]['queue']
        while True:
            try:
                event = q.get(timeout=30)
                yield f'data: {json.dumps(event)}\n\n'
                if event.get('type') == 'done':
                    break
            except queue.Empty:
                yield 'data: {"type":"keepalive"}\n\n'

    return Response(
        stream_with_context(generate()),
        mimetype='text/event-stream',
        headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'},
    )


# ── Download ─────────────────────────────────────────────────────────────────
@app.route('/download/<path:filepath>')
def download(filepath):
    full = os.path.join(DATA_DIR, filepath)
    if not os.path.isfile(full):
        abort(404)
    return send_file(full, as_attachment=True)


# ── HOME ──────────────────────────────────────────────────────────────────────
@app.route('/')
def home():
    estado    = cr._carregar_estado()
    extracoes = estado.get('extrações', {})
    dias_mod  = estado.get('dias_extracao_modulos', {})
    hoje      = date.today()

    cards = []
    for nome, cfg in cr.RELATORIOS_CONFIG.items():
        dia_mod  = dias_mod.get(nome, estado.get('dia_extracao', 6))
        ultima   = extracoes.get(nome, '—')
        proxima  = cr._proxima_extracao(dia_mod)
        atrasado = False
        try:
            atrasado = datetime.strptime(proxima, '%d/%m/%Y').date() < hoje
        except Exception:
            pass
        cards.append(dict(nome=nome, icone=cfg['icone'], cor=cfg['cor'],
                          ultima=ultima, proxima=proxima, atrasado=atrasado,
                          dia_mod=dia_mod))

    _DIAS  = ['Segunda-feira', 'Terca-feira', 'Quarta-feira', 'Quinta-feira',
              'Sexta-feira', 'Sabado', 'Domingo']
    _DIAS_ACC = ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira',
                 'Sexta-feira', 'Sábado', 'Domingo']
    data_pt = (f"{_DIAS_ACC[hoje.weekday()]}, {hoje.day} de "
               f"{MESES_PT_ACC[hoje.month-1]} de {hoje.year}")
    return render_template('home.html', cards=cards, data_pt=data_pt)


@app.route('/agendar', methods=['POST'])
def agendar():
    nome = request.form.get('nome', '')
    try:
        dia = int(request.form.get('dia', ''))
        if 1 <= dia <= 28:
            est = cr._carregar_estado()
            est.setdefault('dias_extracao_modulos', {})[nome] = dia
            cr._salvar_estado(est)
    except (ValueError, TypeError):
        pass
    return redirect(url_for('home'))


# ── MODULO CONTEXT ────────────────────────────────────────────────────────────
def _ctx(titulo, subtitulo, cor, icone, action, hint=''):
    return dict(titulo=titulo, subtitulo=subtitulo, cor=cor, icone=icone,
                action=action, hint=hint, meses=meses_opcoes())


# ── PEDIDOS ───────────────────────────────────────────────────────────────────
@app.route('/pedidos', methods=['GET', 'POST'])
def pedidos():
    if request.method == 'POST':
        f = request.files.get('arquivo')
        if not f or not f.filename:
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        path    = save_upload(f)
        mes_ref = request.form.get('mes_ref', '')
        job_id  = start_job(lambda log: cr.run_pedidos(path, log, mes_ref=mes_ref))
        return jsonify({'job_id': job_id})
    return render_template('modulo.html', **_ctx(
        'Pedidos e Recebimentos', 'Análise de SLA com feriados brasileiros',
        '#4f8ef7', '📦', '/pedidos',
        'A aba de dados é detectada automaticamente.'))


# ── FRETES ────────────────────────────────────────────────────────────────────
@app.route('/fretes', methods=['GET', 'POST'])
def fretes():
    if request.method == 'POST':
        f = request.files.get('arquivo')
        if not f or not f.filename:
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        path    = save_upload(f)
        mes_ref = request.form.get('mes_ref', '')
        job_id  = start_job(lambda log: cr.run_fretes(path, log, mes_ref=mes_ref))
        return jsonify({'job_id': job_id})
    return render_template('modulo.html', **_ctx(
        'Fretes', 'Embarques · RESCOM · Portadores · Custo de Insumos',
        '#7c3aed', '🚚', '/fretes',
        'Abas de Embarques, RESCOM e PORTADORES são detectadas automaticamente.'))


# ── ARMAZENAGEM ───────────────────────────────────────────────────────────────
@app.route('/armazenagem', methods=['GET', 'POST'])
def armazenagem():
    if request.method == 'POST':
        f = request.files.get('arquivo')
        if not f or not f.filename:
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        path    = save_upload(f)
        mes_ref = request.form.get('mes_ref', '')
        try:
            mm, aaaa   = mes_ref.split('-')
            mes_filtro = f'{aaaa}-{mm}'
        except Exception:
            mes_filtro = None
        job_id = start_job(
            lambda log: cr.run_armazenagem(path, mes_filtro, log, mes_ref=mes_ref))
        return jsonify({'job_id': job_id})
    return render_template('modulo.html', **_ctx(
        'Armazenagem', 'Faturamento mensal por cliente',
        '#10b981', '🏭', '/armazenagem'))


# ── ESTOQUE ───────────────────────────────────────────────────────────────────
@app.route('/estoque')
def estoque():
    total, ultima, clientes = cr._db_estoque_info()
    return render_template('estoque.html',
        db_total=total, db_ultima=ultima or '—', db_clientes=clientes,
        meses=meses_opcoes(),
        dias_opts=[30, 60, 90, 120, 150, 180, 270, 365],
        dias_atual=cr._carregar_estado().get('dias_ocioso', 60))


@app.route('/estoque/carga_inicial', methods=['POST'])
def estoque_carga_inicial():
    f = request.files.get('arquivo')
    if not f or not f.filename:
        return jsonify({'error': 'Nenhum arquivo enviado'}), 400
    path = save_upload(f)

    def runner(log):
        db = cr._carregar_estoque_xlsx(path, log)
        if db:
            cr._salvar_db_estoque(db)
            total = sum(len(v) for v in db.values())
            est = cr._carregar_estado()
            est['ultimo_arquivo_estoque'] = path
            cr._salvar_estado(est)
            return True, f'{total} SKUs em {len(db)} clientes carregados.'
        return False, ''

    return jsonify({'job_id': start_db_job(runner)})


@app.route('/estoque/atualizar_mov', methods=['POST'])
def estoque_atualizar_mov():
    f = request.files.get('arquivo')
    if not f or not f.filename:
        return jsonify({'error': 'Nenhum arquivo enviado'}), 400
    path = save_upload(f)

    def runner(log):
        ok = cr._atualizar_db_com_movimentacao(path, log)
        if ok:
            est = cr._carregar_estado()
            est['ultimo_arquivo_mov'] = path
            cr._salvar_estado(est)
            return True, 'Movimentação atualizada com sucesso.'
        return False, ''

    return jsonify({'job_id': start_db_job(runner)})


@app.route('/estoque/gerar', methods=['POST'])
def estoque_gerar():
    f = request.files.get('arquivo_volumes')
    if not f or not f.filename:
        return jsonify({'error': 'Selecione o arquivo de pico de estoque'}), 400
    pasta_v = save_upload(f)
    mes_ref = request.form.get('mes_ref', '')
    dias    = int(request.form.get('dias_ocioso', 60))
    job_id  = start_job(
        lambda log: cr.run_estoque('', pasta_v, '', dias, log, mes_ref=mes_ref))
    return jsonify({'job_id': job_id})


# ── PRODUTIVIDADE ─────────────────────────────────────────────────────────────
@app.route('/produtividade', methods=['GET', 'POST'])
def produtividade():
    if request.method == 'POST':
        f = request.files.get('arquivo')
        if not f or not f.filename:
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        path    = save_upload(f)
        mes_ref = request.form.get('mes_ref', '')
        try:
            mm, aaaa   = mes_ref.split('-')
            mes_filtro = f'{aaaa}-{mm}'
        except Exception:
            mes_filtro = None
        job_id = start_job(
            lambda log: cr.run_produtividade(path, mes_filtro, log, mes_ref=mes_ref))
        return jsonify({'job_id': job_id})
    return render_template('modulo.html', **_ctx(
        'Produtividade de Equipe',
        'Utilização · Ociosidade · Ranking por etapa · Turno 08h–18h',
        '#06b6d4', '👥', '/produtividade',
        'Horários sobrepostos são mesclados automaticamente. A aba é detectada automaticamente.'))


# ── CAP. OPERACIONAL ──────────────────────────────────────────────────────────
@app.route('/cap_operacional', methods=['GET', 'POST'])
def cap_operacional():
    if request.method == 'POST':
        f = request.files.get('arquivo')
        if not f or not f.filename:
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        path      = save_upload(f)
        mes_ref   = request.form.get('mes_ref', '')
        lim_media = float(request.form.get('lim_media', 3))
        lim_alta  = float(request.form.get('lim_alta',  5))
        job_id    = start_job(
            lambda log: cr.run_cap_operacional_pdf(path, mes_ref, log,
                                                   lim_media, lim_alta))
        return jsonify({'job_id': job_id})
    return render_template('cap_operacional.html', meses=meses_opcoes())


# ── RECEBIMENTOS ──────────────────────────────────────────────────────────────
@app.route('/recebimentos', methods=['GET', 'POST'])
def recebimentos():
    if request.method == 'POST':
        f = request.files.get('arquivo')
        if not f or not f.filename:
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        path    = save_upload(f)
        mes_ref = request.form.get('mes_ref', '')
        job_id  = start_job(lambda log: cr.run_recebimentos(path, mes_ref, log))
        return jsonify({'job_id': job_id})
    return render_template('modulo.html', **_ctx(
        'Recebimentos e Devoluções',
        'Entradas · Devoluções · Retiradas por depositante',
        '#0891b2', '📥', '/recebimentos',
        'Exportar como XLS do sistema ESL, converter para .xlsx antes de enviar.'))


# ── FINANCEIRO (visão geral) ──────────────────────────────────────────────────
@app.route('/financeiro')
def financeiro():
    return render_template('financeiro.html')


# ── CONTAS A PAGAR/RECEBER ────────────────────────────────────────────────────
@app.route('/contas', methods=['GET', 'POST'])
def contas():
    if request.method == 'POST':
        fp = request.files.get('arquivo_pagar')
        fr = request.files.get('arquivo_receber')
        if (not fp or not fp.filename) and (not fr or not fr.filename):
            return jsonify({'error': 'Selecione ao menos um arquivo'}), 400
        path_p  = save_upload(fp) if fp and fp.filename else None
        path_r  = save_upload(fr) if fr and fr.filename else None
        mes_ref = request.form.get('mes_ref', '')
        job_id  = start_job(
            lambda log: cr.run_financeiro(path_p, path_r, mes_ref, log))
        return jsonify({'job_id': job_id})
    return render_template('contas.html', meses=meses_opcoes())


# ── FATURAMENTO ARMAZENAGEM ───────────────────────────────────────────────────
@app.route('/fat_arm', methods=['GET', 'POST'])
def fat_arm():
    if request.method == 'POST':
        action = request.form.get('action', 'gerar')

        if action == 'carregar_familias':
            f = request.files.get('arquivo')
            if not f or not f.filename:
                return jsonify({'error': 'Nenhum arquivo enviado'}), 400
            path = save_upload(f)
            def runner_fam(log):
                db = cr._carregar_familias_xlsx(path, log)
                if db:
                    cr._salvar_db_familias(db)
                    total = sum(len(v) for v in db.values())
                    return True, f'{total} SKUs em {len(db)} clientes.'
                return False, ''
            return jsonify({'job_id': start_db_job(runner_fam)})

        elif action == 'carregar_config':
            f = request.files.get('arquivo')
            if not f or not f.filename:
                return jsonify({'error': 'Nenhum arquivo enviado'}), 400
            path = save_upload(f)
            def runner_cfg(log):
                db = cr._carregar_config_fat_arm_xlsx(path, log)
                if db:
                    cr._salvar_db_precos_arm(db)
                    total = len(db.get('clientes', {}))
                    return True, f'{total} clientes configurados.'
                return False, ''
            return jsonify({'job_id': start_db_job(runner_cfg)})

        else:  # gerar
            fm = request.files.get('arquivo_mov')
            fv = request.files.get('arquivo_volumes')
            if not fm or not fm.filename or not fv or not fv.filename:
                return jsonify(
                    {'error': 'Selecione os arquivos de movimentação e volumes'}), 400
            path_m  = save_upload(fm)
            path_v  = save_upload(fv)
            mes_ref = request.form.get('mes_ref', '')
            ce      = cr._carregar_estado().get('ultimo_arquivo_estoque') or None
            job_id  = start_job(
                lambda log: cr.run_faturamento_armazenagem(
                    path_m, path_v, mes_ref, log, arquivo_estoque=ce))
            return jsonify({'job_id': job_id})

    # GET
    fam_skus, fam_cli, fam_data   = cr._db_familias_info()
    preco_cli, preco_data          = cr._db_precos_arm_info()
    return render_template('fat_arm.html',
        meses=meses_opcoes(),
        fam_skus=fam_skus, fam_cli=fam_cli, fam_data=fam_data,
        preco_cli=preco_cli, preco_data=preco_data)


# ── DE-PARA ───────────────────────────────────────────────────────────────────
@app.route('/de_para')
def de_para():
    data  = cr._carregar_de_para()
    items = sorted(data.items())
    return render_template('de_para.html', items=items)


@app.route('/de_para/salvar', methods=['POST'])
def de_para_salvar():
    data  = request.get_json(silent=True) or {}
    items = data.get('items', [])
    dp    = {i['codigo'].strip(): i['cliente'].strip()
             for i in items if i.get('codigo', '').strip()}
    cr._salvar_de_para(dp)
    return jsonify({'ok': True, 'total': len(dp)})


@app.route('/de_para/importar', methods=['POST'])
def de_para_importar():
    f = request.files.get('arquivo')
    if not f or not f.filename:
        return jsonify({'error': 'Nenhum arquivo enviado'}), 400
    import pandas as pd
    path = save_upload(f)
    try:
        df   = pd.read_excel(path)
        df.columns = [str(c).strip() for c in df.columns]
        col_cod = next((c for c in df.columns
                        if 'sku' in c.lower() or 'cod' in c.lower()), df.columns[0])
        col_cli = next((c for c in df.columns
                        if 'cli' in c.lower() or 'dep' in c.lower()), df.columns[1])
        dp = cr._carregar_de_para()
        for _, row in df.iterrows():
            c  = str(row[col_cod]).strip()
            cl = str(row[col_cli]).strip()
            if c and c not in ('nan', 'None'):
                dp[c] = cl
        cr._salvar_de_para(dp)
        return jsonify({'ok': True, 'total': len(dp)})
    except Exception as exc:
        return jsonify({'error': str(exc)}), 500


# ── ARQUIVOS ──────────────────────────────────────────────────────────────────
@app.route('/arquivos')
def arquivos():
    modulos = [
        ('Pedidos',          cr.PASTA_PEDIDOS),
        ('Fretes',           cr.PASTA_FRETES),
        ('Armazenagem',      cr.PASTA_ARMAZENAGEM),
        ('Estoque',          cr.PASTA_ESTOQUE),
        ('Produtividade',    cr.PASTA_PRODUTIVIDADE),
        ('Cap. Operacional', cr.PASTA_CAP_OPERACIONAL),
        ('Recebimentos',     cr.PASTA_RECEBIMENTOS),
        ('Financeiro',       cr.PASTA_FINANCEIRO),
        ('Consolidados',     cr.PASTA_CONSOLIDADOS),
    ]
    grupos = []
    for nome, pasta in modulos:
        files = []
        if os.path.isdir(pasta):
            for raiz, dirs, fnames in os.walk(pasta):
                dirs.sort()
                for fn in sorted(fnames):
                    if fn.endswith(('.xlsx', '.xls', '.pdf', '.pbids')):
                        full  = os.path.join(raiz, fn)
                        rel   = os.path.relpath(full, DATA_DIR)
                        mtime = datetime.fromtimestamp(os.path.getmtime(full))
                        files.append({
                            'nome': fn,
                            'rel':  rel,
                            'size': round(os.path.getsize(full) / 1024, 1),
                            'data': mtime.strftime('%d/%m/%Y %H:%M'),
                        })
        grupos.append({'nome': nome, 'arquivos': files})
    return render_template('arquivos.html', grupos=grupos)


# ── CONFIG ────────────────────────────────────────────────────────────────────
@app.route('/config', methods=['GET', 'POST'])
def config():
    if request.method == 'POST':
        try:
            dia = int(request.form.get('dia_extracao', '6'))
            dia = max(1, min(28, dia))
        except ValueError:
            dia = 6
        est = cr._carregar_estado()
        est['dia_extracao'] = dia
        cr._salvar_estado(est)
        return jsonify({'ok': True})
    estado = cr._carregar_estado()
    return render_template('config.html',
        dia_extracao=estado.get('dia_extracao', 6),
        data_dir=DATA_DIR,
        pastas={
            'Pedidos':          cr.PASTA_PEDIDOS,
            'Fretes':           cr.PASTA_FRETES,
            'Armazenagem':      cr.PASTA_ARMAZENAGEM,
            'Estoque':          cr.PASTA_ESTOQUE,
            'Produtividade':    cr.PASTA_PRODUTIVIDADE,
            'Cap. Operacional': cr.PASTA_CAP_OPERACIONAL,
            'Recebimentos':     cr.PASTA_RECEBIMENTOS,
            'Financeiro':       cr.PASTA_FINANCEIRO,
            'Consolidados':     cr.PASTA_CONSOLIDADOS,
        })


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False, threaded=True)
