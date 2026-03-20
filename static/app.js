'use strict';

/* ── Core SSE job runner ────────────────────────────────────────────────── */
function startModule(endpoint, formData, opts) {
  const { logEl, btnEl, statusEl, downloadEl, btnText } = opts;

  btnEl.disabled      = true;
  btnEl.textContent   = '⏳ Gerando...';
  logEl.textContent   = '';
  if (downloadEl) downloadEl.style.display = 'none';
  statusEl.textContent = '';
  statusEl.className   = 'status';

  fetch(endpoint, { method: 'POST', body: formData })
    .then(r => {
      if (!r.ok) return r.json().then(e => { throw new Error(e.error || 'Erro no servidor'); });
      return r.json();
    })
    .then(data => {
      if (data.error) throw new Error(data.error);
      listenJob(data.job_id, opts);
    })
    .catch(err => {
      logEl.textContent    = '❌ ' + err.message;
      btnEl.disabled       = false;
      btnEl.textContent    = btnText || 'Gerar';
    });
}

function listenJob(jobId, opts) {
  const { logEl, btnEl, statusEl, downloadEl, btnText } = opts;
  const es = new EventSource('/stream/' + jobId);

  es.onmessage = function (e) {
    const d = JSON.parse(e.data);

    if (d.type === 'log') {
      logEl.textContent += d.msg;
      logEl.scrollTop    = logEl.scrollHeight;

    } else if (d.type === 'done') {
      es.close();
      btnEl.disabled    = false;
      btnEl.textContent = btnText || 'Gerar';

      if (d.download_url && downloadEl) {
        downloadEl.href        = d.download_url;
        downloadEl.textContent = '⬇️  Baixar ' + (d.filename || 'arquivo');
        downloadEl.style.display = 'inline-flex';
        statusEl.textContent   = '✅ Concluído!';
        statusEl.className     = 'status ok';
      } else if (d.error) {
        statusEl.textContent = '❌ Erro — veja o log acima';
        statusEl.className   = 'status err';
      } else if (d.info) {
        statusEl.textContent = '✅ ' + d.info;
        statusEl.className   = 'status ok';
      } else {
        statusEl.textContent = '✅ Concluído!';
        statusEl.className   = 'status ok';
      }

    } else if (d.type === 'error') {
      es.close();
      logEl.textContent    += '\n❌ ' + d.msg;
      btnEl.disabled        = false;
      btnEl.textContent     = btnText || 'Gerar';
      statusEl.textContent  = '❌ Erro';
      statusEl.className    = 'status err';
    }
    /* keepalive: do nothing */
  };

  es.onerror = function () {
    es.close();
    btnEl.disabled       = false;
    btnEl.textContent    = btnText || 'Gerar';
    statusEl.textContent = '❌ Conexão perdida';
    statusEl.className   = 'status err';
  };
}

/* ── File drop zone helper ──────────────────────────────────────────────── */
function setupDropZone(dropId, inputId, labelId) {
  const drop  = document.getElementById(dropId);
  const input = document.getElementById(inputId);
  const label = document.getElementById(labelId);
  if (!drop || !input) return;

  drop.addEventListener('click', () => input.click());

  input.addEventListener('change', function () {
    if (label) label.textContent = this.files[0]?.name || 'Nenhum arquivo selecionado';
  });

  drop.addEventListener('dragover', e => {
    e.preventDefault();
    drop.classList.add('drag-over');
  });
  drop.addEventListener('dragleave', () => drop.classList.remove('drag-over'));
  drop.addEventListener('drop', e => {
    e.preventDefault();
    drop.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) {
      input.files = e.dataTransfer.files;
      if (label) label.textContent = e.dataTransfer.files[0].name;
    }
  });
}

/* ── Standard modulo form (modulo.html) ─────────────────────────────────── */
function initModulo(endpoint, btnText) {
  setupDropZone('drop-zone', 'arquivo', 'file-name');

  document.getElementById('btn-gerar').addEventListener('click', function () {
    const file = document.getElementById('arquivo').files[0];
    if (!file) { alert('Selecione um arquivo primeiro.'); return; }

    const fd = new FormData();
    fd.append('arquivo', file);
    fd.append('mes_ref', document.getElementById('mes_ref').value);

    startModule(endpoint, fd, {
      logEl:       document.getElementById('log-output'),
      btnEl:       this,
      statusEl:    document.getElementById('status'),
      downloadEl:  document.getElementById('btn-download'),
      btnText:     btnText,
    });
  });
}

/* ── DB upload (single-file, no download) ───────────────────────────────── */
function dbUpload(endpoint, fileInputId, logId, statusId, btnEl, btnText) {
  const file = document.getElementById(fileInputId)?.files[0];
  if (!file) { alert('Selecione um arquivo.'); return; }

  const fd = new FormData();
  fd.append('arquivo', file);
  fd.append('action', endpoint.split('/').pop());

  const logEl    = document.getElementById(logId);
  const statusEl = document.getElementById(statusId);

  startModule(endpoint, fd, {
    logEl, btnEl, statusEl,
    downloadEl: null,
    btnText,
  });
}
