// ── CONFIGURAÇÃO SUPABASE ─────────────────────────────────────────────────────
// Para trocar de banco: altere SB_URL e SB_KEY abaixo
const SB_URL = 'https://lavehpbeizrqcuaiqcpx.supabase.co/rest/v1';
const SB_KEY = 'sb_publishable_Q6SWGBjlO3TIdlO7aq16qQ_0QmTZX1R';
const HEADERS = {
  'Content-Type':  'application/json',
  'apikey':        SB_KEY,
  'Authorization': SB_KEY,
  'Prefer':        'return=representation',
};

// ── HELPERS DE BANCO ──────────────────────────────────────────────────────────
async function dbGet(table, params = '') {
  const r = await fetch(`${SB_URL}/${table}${params ? '?' + params : ''}`, { headers: HEADERS });
  if (!r.ok) { const e = await r.json(); throw new Error(e.message || r.statusText); }
  return r.json();
}

async function dbPost(table, body) {
  const r = await fetch(`${SB_URL}/${table}`, {
    method: 'POST', headers: HEADERS, body: JSON.stringify(body),
  });
  if (!r.ok) { const e = await r.json(); throw new Error(e.message || r.statusText); }
  return r.json();
}

async function dbPatch(table, params, body) {
  const h = { ...HEADERS, 'Prefer': 'return=minimal' };
  const r = await fetch(`${SB_URL}/${table}?${params}`, {
    method: 'PATCH', headers: h, body: JSON.stringify(body),
  });
  if (!r.ok) { const e = await r.json(); throw new Error(e.message || r.statusText); }
}

async function dbDelete(table, params) {
  const r = await fetch(`${SB_URL}/${table}?${params}`, { method: 'DELETE', headers: HEADERS });
  if (!r.ok) { const e = await r.json(); throw new Error(e.message || r.statusText); }
}

function setDbStatus(ok, msg) {
  const dot = document.getElementById('dbDot');
  const txt = document.getElementById('sbStatus');
  dot.className = 'db-dot ' + (ok ? 'ok' : 'err');
  txt.textContent = ok ? 'banco conectado' : 'erro: ' + msg;
}

window.addEventListener('load', async () => {
  try {
    await dbGet('pedidos', 'select=id&limit=1');
    setDbStatus(true, '');
    loadSavedCount();
  } catch (e) {
    setDbStatus(false, e.message);
  }
});

// ── ESTADO ────────────────────────────────────────────────────────────────────
let state = { pedido: '', pedidoSeq: '', cliente: '', componentes: [] };
let existingPedidoId = null;
let filterMode = 'all';
let searchTerm = '';

// ── TABS ──────────────────────────────────────────────────────────────────────
function switchTab(tab) {
  ['tabExtrator', 'tabSaved'].forEach(id => document.getElementById(id).classList.remove('active'));
  ['pageExtrator', 'pageSaved'].forEach(id => document.getElementById(id).classList.remove('active'));
  if (tab === 'extrator') {
    document.getElementById('tabExtrator').classList.add('active');
    document.getElementById('pageExtrator').classList.add('active');
  } else {
    document.getElementById('tabSaved').classList.add('active');
    document.getElementById('pageSaved').classList.add('active');
    loadSavedList();
  }
}

// ── EXTRAÇÃO ──────────────────────────────────────────────────────────────────
function extractFromRows(rows) {
  const componentes = [];
  let pedido = '', cliente = '';
  let current = null, pendingOps = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (/pedido:/i.test(row)) {
      const m = row.match(/pedido:\s*(\d+)/i);
      if (m) pedido = m[1];
      const cm = row.match(/cliente:\s*\d+\s+(.+?)(?:\s{2,}|$)/i);
      if (cm) cliente = cm[1].trim();
    }
    const compM = row.match(/^(\d+(?:\.\d+)*)\s*-\s*(\d+\.\d+)\s*-\s*(.+?)(?=\t{2,}|$)/);
    if (compM) {
      if (current) { current.operacoes = [...pendingOps]; componentes.push(current); pendingOps = []; }

      const dM = row.match(/([A-ZÁÇÃÕ]{2,}[\d]+(?:\.[\d]+)+)/);
      let desenho = dM ? dM[0] : '';

      const qM = row.match(/(\d+(?:[.,]\d+)?)\s*PC/i);
      let quantidade = qM ? qM[1].replace(',', '.') : '';

      let descricao = compM[3].trim();
      if (desenho) descricao = descricao.replace(desenho, '').trim();
      descricao = descricao.replace(/\s*[-–]?\s*\d+(?:[.,]\d+)?\s*PC.*/i, '').trim();
      descricao = descricao.replace(/[-–\s]+$/, '').trim();

      current = { posicao: compM[1], codigo: compM[2], descricao, desenho, quantidade, operacoes: [] };
    }
    const parts = row.split('\t');
    if (parts.length >= 7 && /^\d+$/.test(parts[0]) && /^\d+$/.test(parts[1]) && parts[2] && parts[3]) {
      pendingOps.push({
        est: parts[0], seq: parts[1], op: parts[2], desc: parts[3],
        qtd_prev: parts[5] || '', qtd_real: parts[6] || '',
        inicio_data: parts[9] || '', inicio_hora: parts[10] || '',
        fim_data: parts[11] || '', fim_hora: parts[12] || '',
        situacao: parts[13] || parts[12] || '',
      });
    }
  }
  if (current) { current.operacoes = [...pendingOps]; componentes.push(current); }
  return { pedido, cliente, componentes };
}

async function handleFile(file) {
  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      const rows = json.map(r => r.filter(c => c?.toString().trim()).join('\t')).filter(r => r.trim());
      state = extractFromRows(rows);
      state.pedidoSeq = state.pedido;

      document.getElementById('fileTag').style.display = 'flex';
      document.getElementById('fileTagName').textContent = file.name;
      document.getElementById('pedidoPanel').style.display = 'block';
      document.getElementById('pedidoNum').textContent = '...';
      document.getElementById('pedidoCli').textContent = state.cliente || '—';
      document.getElementById('filterBar').style.display = 'flex';

      await checkExisting();
      renderResult();
      const tot = state.componentes.reduce((s, c) => s + (c.operacoes?.length || 0), 0);
      toast(`Extraídos: ${state.componentes.length} itens · ${tot} operações`, 'success');
    } catch (err) {
      toast('Erro ao processar o arquivo', 'error');
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
}

async function checkExisting() {
  if (!state.pedido) return;
  try {
    const data  = await dbGet('pedidos', `select=id,numero&numero=like.${state.pedido}-%&order=numero.asc`);
    const exact = await dbGet('pedidos', `select=id,numero&numero=eq.${state.pedido}&limit=1`);
    const todos = [...(exact || []), ...(data || [])];

    if (todos.length === 0) {
      existingPedidoId = null;
      state.pedidoSeq = state.pedido + '-1';
      setFlag('novo');
      document.getElementById('statusBadge').textContent = 'NOVO PEDIDO';
      document.getElementById('statusBadge').className = 'status-badge ok';
    } else {
      existingPedidoId = null;
      let maxSeq = 0;
      todos.forEach(p => {
        const m = p.numero.match(/-(\d+)$/);
        if (m) maxSeq = Math.max(maxSeq, parseInt(m[1]));
        else maxSeq = Math.max(maxSeq, 1);
      });
      state.pedidoSeq = state.pedido + '-' + (maxSeq + 1);
      setFlag('seq');
      document.getElementById('statusBadge').textContent = 'SEQUÊNCIA ' + (maxSeq + 1);
      document.getElementById('statusBadge').className = 'status-badge warn';
    }

    document.getElementById('pedidoNum').textContent = state.pedidoSeq;
    document.getElementById('saveBtn').disabled = false;
  } catch (e) {
    existingPedidoId = null;
    state.pedidoSeq = state.pedido + '-1';
    document.getElementById('saveBtn').disabled = false;
    setFlag('novo');
  }
}

function setFlag(tipo) {
  const el = document.getElementById('pedidoFlag');
  if (tipo === 'novo')  el.innerHTML = `<span class="pedido-flag flag-novo">✓ NOVO · ${state.pedidoSeq}</span>`;
  if (tipo === 'seq')   el.innerHTML = `<span class="pedido-flag flag-exist">+ SEQUÊNCIA · ${state.pedidoSeq}</span>`;
  if (tipo === 'saved') el.innerHTML = `<span class="pedido-flag flag-saved">✓ SALVO · ${state.pedidoSeq}</span>`;
}

// ── SALVAR NO BANCO ───────────────────────────────────────────────────────────
async function executeSave() {
  if (!state.pedidoSeq) { toast('Nenhum dado para salvar', 'error'); return; }

  const btn = document.getElementById('saveBtn');
  btn.disabled = true; btn.textContent = '⏳ Salvando...';
  const pw = document.getElementById('progressWrap');
  const pf = document.getElementById('progressFill');
  const pl = document.getElementById('progressLabel');
  pw.style.display = 'block'; pf.style.width = '0%';

  try {
    pl.textContent = 'Criando pedido ' + state.pedidoSeq + '...';
    const np = await dbPost('pedidos', {
      numero:          state.pedidoSeq,
      cliente:         state.cliente,
      data_importacao: new Date().toISOString(),
    });
    const pedidoId = np[0].id;
    pf.style.width = '10%';

    const total = state.componentes.length;
    for (let i = 0; i < total; i++) {
      const comp = state.componentes[i];
      pl.textContent = `Salvando item ${i + 1} de ${total}...`;
      pf.style.width = (10 + Math.round((i / total) * 85)) + '%';

      const nc = await dbPost('componentes', {
        pedido_id:  pedidoId,
        posicao:    comp.posicao,
        codigo:     comp.codigo,
        descricao:  comp.descricao,
        desenho:    comp.desenho,
        quantidade: parseFloat(comp.quantidade) || 0,
        unidade:    'PC',
      });
      const compId = nc[0].id;

      if (comp.operacoes?.length) {
        await dbPost('operacoes', comp.operacoes.map(op => ({
          componente_id: compId,
          pedido_id:     pedidoId,
          est:           op.est,
          seq:           op.seq,
          operacao:      op.op,
          descricao:     op.desc,
          qtd_prevista:  parseFloat(op.qtd_prev) || 0,
          qtd_realizada: parseFloat(op.qtd_real) || 0,
          inicio_data:   op.inicio_data || null,
          inicio_hora:   op.inicio_hora || null,
          fim_data:      op.fim_data    || null,
          fim_hora:      op.fim_hora    || null,
          situacao:      op.situacao    || 'Aberta',
        })));
      }
    }

    pf.style.width = '100%';
    pl.textContent = 'Concluído!';
    setTimeout(() => { pw.style.display = 'none'; }, 1500);

    setFlag('saved');
    btn.textContent = '✓ Salvo!';
    document.getElementById('statusBadge').textContent = 'SALVO · ' + state.pedidoSeq;
    document.getElementById('statusBadge').className = 'status-badge ok';
    loadSavedCount();
    toast(`Pedido ${state.pedidoSeq} salvo com sucesso!`, 'success');
  } catch (err) {
    console.error(err);
    pw.style.display = 'none';
    toast('Erro: ' + err.message, 'error');
    btn.disabled = false; btn.textContent = '▶ Salvar no banco';
  }
}

// ── RENDER EXTRATOR ───────────────────────────────────────────────────────────
function getFiltered() {
  let list = state.componentes;
  if (filterMode === 'com-ops') list = list.filter(c => c.operacoes?.length > 0);
  if (filterMode === 'sem-ops') list = list.filter(c => !c.operacoes?.length);
  if (searchTerm) {
    const q = searchTerm.toLowerCase();
    list = list.filter(c =>
      c.codigo.toLowerCase().includes(q) ||
      c.descricao.toLowerCase().includes(q) ||
      c.posicao.toLowerCase().includes(q) ||
      (c.desenho || '').toLowerCase().includes(q)
    );
  }
  return list;
}

function renderResult() {
  const area = document.getElementById('resultContent');
  document.getElementById('emptyState').style.display = 'none';
  area.style.display = 'block';
  const filtered = getFiltered();
  const totalOps = state.componentes.reduce((s, c) => s + (c.operacoes?.length || 0), 0);
  document.getElementById('statComp').textContent = state.componentes.length;
  document.getElementById('statOps').textContent = totalOps;
  document.getElementById('toolbarInfo').innerHTML =
    `<strong>Pedido ${state.pedidoSeq || '—'}</strong> · ${state.componentes.length} itens · ${totalOps} operações · mostrando <strong>${filtered.length}</strong>`;

  if (!filtered.length) {
    area.innerHTML = `<div class="empty-state" style="height:300px"><div class="empty-icon">🔍</div><div class="empty-title">Nenhum item encontrado</div></div>`;
    return;
  }

  let html = `
    <div class="pedido-header">
      <div class="ph-item"><div class="ph-lbl">Pedido</div><div class="ph-val">${state.pedidoSeq || '—'}</div></div>
      <div class="ph-item"><div class="ph-lbl">Cliente</div><div class="ph-val">${state.cliente || '—'}</div></div>
      <div class="ph-item"><div class="ph-lbl">Itens</div><div class="ph-val">${state.componentes.length}</div></div>
      <div class="ph-item"><div class="ph-lbl">Operações</div><div class="ph-val">${totalOps}</div></div>
      <div class="ph-actions">
        <button class="ph-btn" onclick="expandAll()">▼ Expandir tudo</button>
        <button class="ph-btn" onclick="collapseAll()">▶ Recolher tudo</button>
      </div>
    </div>
    <div class="col-header">
      <div class="ch">Pos.</div>
      <div class="ch">Código</div>
      <div class="ch">Descrição</div>
      <div class="ch">Desenho</div>
      <div class="ch">Qtd</div>
      <div class="ch" style="text-align:center">Ops</div>
    </div>`;

  filtered.forEach((comp, idx) => {
    const isLast = idx === filtered.length - 1;
    const opsHtml = comp.operacoes?.length
      ? comp.operacoes.map(op => {
          const s = (op.situacao || '').toLowerCase();
          let cls = 'sit-a';
          if (/concl|finaliz/i.test(s)) cls = 'sit-c';
          else if (/anda|prog|exec/i.test(s)) cls = 'sit-p';
          return `<tr>
            <td>${op.est || '—'}</td><td>${op.seq || '—'}</td><td>${op.op || '—'}</td>
            <td>${op.desc || '—'}</td><td>${op.qtd_prev || '—'}</td><td>${op.qtd_real || '—'}</td>
            <td>${op.inicio_data || '—'} ${op.inicio_hora || ''}</td>
            <td>${op.fim_data || '—'} ${op.fim_hora || ''}</td>
            <td><span class="sit-badge ${cls}">${op.situacao || 'Aberta'}</span></td>
          </tr>`;
        }).join('')
      : `<tr><td colspan="9" style="padding:12px;text-align:center;color:var(--tx3);font-size:11px">Sem operações</td></tr>`;

    const qty = comp.quantidade
      ? parseFloat(comp.quantidade) % 1 === 0 ? parseInt(comp.quantidade) : parseFloat(comp.quantidade)
      : '—';

    html += `
      <div class="comp-wrap${isLast ? ' last' : ''}">
        <div class="comp-row${isLast ? ' last-row' : ''}" onclick="toggleCard(${idx})">
          <div class="cr-cell cr-pos">${comp.posicao}</div>
          <div class="cr-cell cr-cod">${comp.codigo}</div>
          <div class="cr-cell cr-desc">${comp.descricao}</div>
          <div class="cr-cell cr-desenho">${comp.desenho || '—'}</div>
          <div class="cr-cell cr-qty">${qty}</div>
          <div class="cr-cell cr-ops"><span class="tog" id="tog-${idx}">▶</span> ${comp.operacoes?.length || 0}</div>
        </div>
        <div class="ops-wrap" id="ops-${idx}">
          <table class="ops-table">
            <thead><tr>
              <th>EST</th><th>SEQ</th><th>OP</th><th>DESCRIÇÃO</th>
              <th>QTD.PREV</th><th>QTD.REAL</th><th>INÍCIO</th><th>FIM</th><th>SITUAÇÃO</th>
            </tr></thead>
            <tbody>${opsHtml}</tbody>
          </table>
        </div>
      </div>`;
  });

  area.innerHTML = html;
}

function toggleCard(idx) {
  const ops = document.getElementById('ops-' + idx);
  const tog = document.getElementById('tog-' + idx);
  if (!ops) return;
  const open = ops.classList.toggle('open');
  tog.classList.toggle('open', open);
  tog.textContent = open ? '▼' : '▶';
}
function expandAll() {
  document.querySelectorAll('.ops-wrap').forEach(el => el.classList.add('open'));
  document.querySelectorAll('.tog').forEach(el => { el.classList.add('open'); el.textContent = '▼'; });
}
function collapseAll() {
  document.querySelectorAll('.ops-wrap').forEach(el => el.classList.remove('open'));
  document.querySelectorAll('.tog').forEach(el => { el.classList.remove('open'); el.textContent = '▶'; });
}
function setFilter(mode, el) {
  filterMode = mode;
  document.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
  el.classList.add('active');
  renderResult();
}
function applyFilter() {
  searchTerm = document.getElementById('searchInput').value;
  renderResult();
}

// ── PEDIDOS SALVOS ────────────────────────────────────────────────────────────
async function loadSavedCount() {
  try {
    const data = await dbGet('pedidos', 'select=id');
    document.getElementById('savedCount').textContent = data?.length || 0;
  } catch (e) { /* silencioso */ }
}

async function loadSavedList() {
  const area = document.getElementById('savedListArea');
  area.innerHTML = '<div class="empty-state" style="height:200px"><div class="empty-title">Carregando...</div></div>';
  try {
    const pedidos = await dbGet('pedidos', 'select=id,numero,cliente,data_importacao,data_atualizacao&order=data_importacao.desc');
    if (!pedidos?.length) {
      area.innerHTML = '<div class="empty-state" style="height:300px"><div class="empty-title">Nenhum pedido salvo</div><div class="empty-sub">Importe e salve uma planilha na aba Extrator</div></div>';
      return;
    }
    const comps = await dbGet('componentes', 'select=pedido_id');
    const ops   = await dbGet('operacoes',   'select=pedido_id');
    const ccMap = {}, ocMap = {};
    (comps || []).forEach(r => { ccMap[r.pedido_id] = (ccMap[r.pedido_id] || 0) + 1; });
    (ops   || []).forEach(r => { ocMap[r.pedido_id] = (ocMap[r.pedido_id] || 0) + 1; });

    let html = `<table class="saved-table"><thead><tr>
      <th>Nº Pedido</th><th>Cliente</th><th>Itens</th><th>Operações</th><th>Importado em</th><th>Atualizado</th><th>Ações</th>
    </tr></thead><tbody>`;
    pedidos.forEach(p => {
      const dt = p.data_importacao  ? new Date(p.data_importacao ).toLocaleString('pt-BR') : '—';
      const up = p.data_atualizacao ? new Date(p.data_atualizacao).toLocaleString('pt-BR') : '—';
      const cli = (p.cliente || '').replace(/'/g, "\\'");
      html += `<tr>
        <td class="mono" style="font-weight:600;color:var(--blue)">${p.numero}</td>
        <td>${p.cliente || '—'}</td>
        <td class="mono">${ccMap[p.id] || 0}</td>
        <td class="mono">${ocMap[p.id] || 0}</td>
        <td class="mono" style="color:var(--tx2);font-size:11px">${dt}</td>
        <td class="mono" style="color:var(--tx2);font-size:11px">${up}</td>
        <td><div style="display:flex;gap:6px">
          <button class="icon-btn" onclick="viewSavedPedido('${p.id}','${p.numero}','${cli}')">👁 Ver</button>
          <button class="icon-btn danger" onclick="deletePedido('${p.id}','${p.numero}')">🗑 Remover</button>
        </div></td>
      </tr>`;
    });
    html += `</tbody></table>`;
    area.innerHTML = html;
    document.getElementById('savedCount').textContent = pedidos.length;
  } catch (e) {
    area.innerHTML = `<div class="empty-state"><div class="empty-title">Erro: ${e.message}</div></div>`;
  }
}

async function viewSavedPedido(id, numero, cliente) {
  try {
    const comps = await dbGet('componentes', `select=*&pedido_id=eq.${id}&order=posicao`);
    const ops   = await dbGet('operacoes',   `select=*&pedido_id=eq.${id}`);
    const opsMap = {};
    (ops || []).forEach(o => { if (!opsMap[o.componente_id]) opsMap[o.componente_id] = []; opsMap[o.componente_id].push(o); });
    const componentes = (comps || []).map(c => ({
      posicao: c.posicao, codigo: c.codigo, descricao: c.descricao,
      desenho: c.desenho, quantidade: c.quantidade,
      operacoes: (opsMap[c.id] || []).map(o => ({
        est: o.est, seq: o.seq, op: o.operacao, desc: o.descricao,
        qtd_prev: o.qtd_prevista, qtd_real: o.qtd_realizada,
        inicio_data: o.inicio_data, inicio_hora: o.inicio_hora,
        fim_data: o.fim_data, fim_hora: o.fim_hora, situacao: o.situacao,
      })),
    }));
    state = { pedido: numero, pedidoSeq: numero, cliente, componentes };
    existingPedidoId = id;
    switchTab('extrator');
    document.getElementById('fileTag').style.display = 'flex';
    document.getElementById('fileTagName').textContent = 'Pedido ' + numero + ' (banco)';
    document.getElementById('pedidoPanel').style.display = 'block';
    document.getElementById('pedidoNum').textContent = numero;
    document.getElementById('pedidoCli').textContent = cliente || '—';
    document.getElementById('filterBar').style.display = 'flex';
    document.getElementById('saveBtn').disabled = false;
    setFlag('saved');
    document.getElementById('statusBadge').textContent = 'VISUALIZANDO';
    document.getElementById('statusBadge').className = 'status-badge ok';
    renderResult();
  } catch (e) {
    toast('Erro ao carregar pedido: ' + e.message, 'error');
  }
}

async function deletePedido(id, numero) {
  if (!confirm(`Remover pedido Nº ${numero} e todos os seus dados?`)) return;
  try {
    const comps = await dbGet('componentes', `select=id&pedido_id=eq.${id}`);
    if (comps?.length) {
      await dbDelete('operacoes',   `componente_id=in.(${comps.map(c => c.id).join(',')})`);
      await dbDelete('componentes', `pedido_id=eq.${id}`);
    }
    await dbDelete('pedidos', `id=eq.${id}`);
    toast('Pedido ' + numero + ' removido', 'warn');
    loadSavedList();
    loadSavedCount();
  } catch (e) {
    toast('Erro: ' + e.message, 'error');
  }
}

// ── RESET ─────────────────────────────────────────────────────────────────────
function resetAll() {
  state = { pedido: '', pedidoSeq: '', cliente: '', componentes: [] };
  existingPedidoId = null; filterMode = 'all'; searchTerm = '';
  document.getElementById('pedidoPanel').style.display = 'none';
  document.getElementById('fileTag').style.display = 'none';
  document.getElementById('fileInput').value = '';
  document.getElementById('filterBar').style.display = 'none';
  document.getElementById('saveBtn').disabled = true;
  document.getElementById('saveBtn').textContent = '▶ Salvar no banco';
  document.getElementById('statComp').textContent = '0';
  document.getElementById('statOps').textContent = '0';
  document.getElementById('toolbarInfo').textContent = 'Nenhuma planilha importada';
  document.getElementById('statusBadge').textContent = 'AGUARDANDO';
  document.getElementById('statusBadge').className = 'status-badge';
  document.getElementById('emptyState').style.display = 'flex';
  document.getElementById('resultContent').style.display = 'none';
  document.getElementById('resultContent').innerHTML = '';
  document.getElementById('searchInput').value = '';
  document.getElementById('progressWrap').style.display = 'none';
  toast('Tudo limpo', 'warn');
}

// ── TOAST ─────────────────────────────────────────────────────────────────────
function toast(msg, type = 'info') {
  const stack = document.getElementById('toasts');
  const t = document.createElement('div');
  t.className = 'toast ' + (type === 'success' ? 'success' : type === 'error' ? 'error' : type === 'warn' ? 'warn' : '');
  t.textContent = msg;
  stack.appendChild(t);
  requestAnimationFrame(() => requestAnimationFrame(() => t.classList.add('show')));
  setTimeout(() => { t.classList.remove('show'); setTimeout(() => t.remove(), 250); }, 3500);
}

// ── UPLOAD EVENTS ─────────────────────────────────────────────────────────────
const zone = document.getElementById('uploadZone');
const fi   = document.getElementById('fileInput');
zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag'); });
zone.addEventListener('dragleave', () => zone.classList.remove('drag'));
zone.addEventListener('drop', e => {
  e.preventDefault(); zone.classList.remove('drag');
  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
fi.addEventListener('change', e => { if (e.target.files.length) handleFile(e.target.files[0]); });
