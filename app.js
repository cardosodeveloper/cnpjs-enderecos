(() => {
  // ---------- Estado ----------
  const state = {
    wb: null, fileName: '', sheetName: '', rows: [], header: [], colIdx: -1,
    items: [],          // { row, raw, cnpj, valid }
    results: new Map(), // cnpj -> { status: 'ok'|'nf'|'err'|'wait', data, msg, source }
    running: false, paused: false, cancelled: false,
    filter: 'all', startedAt: 0, doneThisRun: 0
  };
  const $ = id => document.getElementById(id);

  // ---------- CNPJ ----------
  function normalizeCnpj(v) {
    if (v === null || v === undefined || v === '') return '';
    if (typeof v === 'number') return String(Math.round(v)).padStart(14, '0');
    let s = String(v).trim().toUpperCase().replace(/[.\-\/\s]/g, '');
    if (/^\d+$/.test(s) && s.length < 14 && s.length >= 8) s = s.padStart(14, '0');
    return s;
  }
  function dv(base) {
    const w = base.length === 12 ? [5,4,3,2,9,8,7,6,5,4,3,2] : [6,5,4,3,2,9,8,7,6,5,4,3,2];
    let sum = 0;
    for (let i = 0; i < base.length; i++) sum += (base.charCodeAt(i) - 48) * w[i];
    const r = sum % 11;
    return r < 2 ? 0 : 11 - r;
  }
  function isValidCnpj(c) {
    if (!/^[0-9A-Z]{12}\d{2}$/.test(c)) return false;
    if (/^(\d)\1{13}$/.test(c)) return false;
    const d1 = dv(c.slice(0, 12));
    const d2 = dv(c.slice(0, 12) + d1);
    return c.endsWith(`${d1}${d2}`);
  }
  const fmtCnpj = c => c.length === 14 ? `${c.slice(0,2)}.${c.slice(2,5)}.${c.slice(5,8)}/${c.slice(8,12)}-${c.slice(12)}` : c;
  const fmtCep = c => { const s = String(c || '').replace(/\D/g, '').padStart(8, '0'); return s === '00000000' ? '' : `${s.slice(0,5)}-${s.slice(5)}`; };

  // ---------- Leitura da planilha ----------
  const drop = $('drop'), fileInput = $('file');
  fileInput.addEventListener('change', e => e.target.files[0] && loadFile(e.target.files[0]));
  drop.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); } });
  ['dragenter','dragover'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave','drop'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => e.dataTransfer.files[0] && loadFile(e.dataTransfer.files[0]));

  async function loadFile(file) {
    if (state.running) return;
    try {
      const buf = await file.arrayBuffer();
      state.wb = XLSX.read(buf, { type: 'array', raw: true });
      state.fileName = file.name;
      const sel = $('sheetSel');
      sel.innerHTML = state.wb.SheetNames.map(n => `<option>${esc(n)}</option>`).join('');
      selectSheet(state.wb.SheetNames[0]);
      $('setup').classList.remove('hidden');
    } catch (err) {
      alert('Não foi possível ler o arquivo. Verifique se é uma planilha .xlsx, .xls ou .csv válida.');
      console.error(err);
    }
  }
  $('sheetSel').addEventListener('change', e => selectSheet(e.target.value));
  $('colSel').addEventListener('change', e => { state.colIdx = +e.target.value; buildItems(); });

  function selectSheet(name) {
    state.sheetName = name;
    const ws = state.wb.Sheets[name];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true, blankrows: false });
    state.header = (aoa[0] || []).map((h, i) => String(h || `Coluna ${i + 1}`));
    state.rows = aoa.slice(1);
    state.colIdx = detectColumn();
    $('colSel').innerHTML = state.header.map((h, i) => `<option value="${i}" ${i === state.colIdx ? 'selected' : ''}>${esc(h)}</option>`).join('');
    buildItems();
  }

  function detectColumn() {
    const byName = state.header.findIndex(h => h.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').includes('cnpj'));
    if (byName >= 0) return byName;
    let best = 0, bestScore = -1;
    state.header.forEach((_, i) => {
      const score = state.rows.slice(0, 80).filter(r => isValidCnpj(normalizeCnpj(r[i]))).length;
      if (score > bestScore) { bestScore = score; best = i; }
    });
    return best;
  }

  function buildItems() {
    state.items = state.rows.map((r, i) => {
      const raw = r[state.colIdx];
      const cnpj = normalizeCnpj(raw);
      return { row: i + 2, raw, cnpj, valid: isValidCnpj(cnpj) };
    }).filter(it => it.raw !== '' && it.raw !== null);
    state.results = new Map();
    const unique = new Set(state.items.filter(i => i.valid).map(i => i.cnpj));
    $('fName').textContent = state.fileName;
    $('fRows').textContent = state.rows.length.toLocaleString('pt-BR');
    $('fValid').textContent = unique.size.toLocaleString('pt-BR');
    $('fInvalid').textContent = state.items.filter(i => !i.valid).length.toLocaleString('pt-BR');
    $('btnStart').disabled = unique.size === 0;
    $('progress').classList.add('hidden');
    renderTable();
    $('results').classList.toggle('hidden', state.items.length === 0);
    $('btnExport').disabled = true;
  }

  // ---------- Fontes de dados ----------
  const PROVIDERS = {
    brasilapi: {
      name: 'BrasilAPI',
      url: c => `https://brasilapi.com.br/api/cnpj/v1/${c}`,
      map: d => fromReceitaSchema(d)
    },
    minhareceita: {
      name: 'Minha Receita',
      url: c => `https://minhareceita.org/${c}`,
      map: d => fromReceitaSchema(d)
    },
    cnpjws: {
      name: 'CNPJ.ws',
      url: c => `https://publica.cnpj.ws/cnpj/${c}`,
      map: d => {
        const e = d.estabelecimento || {};
        return {
          razao_social: d.razao_social || '',
          nome_fantasia: e.nome_fantasia || '',
          situacao: e.situacao_cadastral || '',
          data_situacao: e.data_situacao_cadastral || '',
          tipo_logradouro: e.tipo_logradouro || '',
          logradouro: e.logradouro || '',
          numero: e.numero || '',
          complemento: e.complemento || '',
          bairro: e.bairro || '',
          municipio: (e.cidade && e.cidade.nome) || '',
          uf: (e.estado && e.estado.sigla) || '',
          cep: fmtCep(e.cep),
          telefone: e.ddd1 && e.telefone1 ? `(${e.ddd1}) ${e.telefone1}` : '',
          email: e.email || '',
          codigo_ibge_api: (e.cidade && e.cidade.ibge_id) ? String(e.cidade.ibge_id) : ''
        };
      }
    }
  };
  function fromReceitaSchema(d) {
    return {
      razao_social: d.razao_social || '',
      nome_fantasia: d.nome_fantasia || '',
      situacao: d.descricao_situacao_cadastral || '',
      data_situacao: d.data_situacao_cadastral || '',
      tipo_logradouro: d.descricao_tipo_de_logradouro || '',
      logradouro: d.logradouro || '',
      numero: d.numero || '',
      complemento: d.complemento || '',
      bairro: d.bairro || '',
      municipio: d.municipio || '',
      uf: d.uf || '',
      cep: fmtCep(d.cep),
      telefone: d.ddd_telefone_1 || '',
      email: d.email || '',
      codigo_ibge_api: d.codigo_municipio_ibge ? String(d.codigo_municipio_ibge) : ''
    };
  }
  function fullAddress(a) {
    const tipo = a.tipo_logradouro && !a.logradouro.toUpperCase().startsWith(a.tipo_logradouro.toUpperCase()) ? a.tipo_logradouro + ' ' : '';
    let s = `${tipo}${a.logradouro}`.trim();
    if (a.numero) s += `, ${a.numero}`;
    if (a.complemento) s += ` - ${a.complemento}`;
    if (a.bairro) s += ` - ${a.bairro}`;
    if (a.municipio) s += `, ${a.municipio}${a.uf ? '/' + a.uf : ''}`;
    if (a.cep) s += ` - CEP ${a.cep}`;
    return s.replace(/\s+/g, ' ').trim();
  }

  const sleep = ms => new Promise(r => setTimeout(r, ms));

  async function fetchFrom(provKey, cnpj) {
    const p = PROVIDERS[provKey];
    const waits = [2000, 5000, 12000];
    for (let attempt = 0; attempt <= waits.length; attempt++) {
      if (state.cancelled) throw { kind: 'cancel' };
      let res;
      try {
        res = await fetch(p.url(cnpj), { headers: { Accept: 'application/json' } });
      } catch (e) {
        if (attempt < waits.length) { setStatus(`${p.name}: falha de conexão, tentando de novo…`); await sleep(waits[attempt]); continue; }
        throw { kind: 'err', msg: 'Falha de conexão' };
      }
      if (res.ok) return { data: p.map(await res.json()), source: p.name };
      if (res.status === 404 || res.status === 400) throw { kind: 'nf', msg: 'CNPJ não encontrado na base' };
      if ((res.status === 429 || res.status >= 500) && attempt < waits.length) {
        setStatus(`${p.name} limitou as consultas (HTTP ${res.status}). Aguardando ${waits[attempt] / 1000}s…`);
        await sleep(waits[attempt]);
        continue;
      }
      throw { kind: 'err', msg: `HTTP ${res.status}` };
    }
  }

  async function lookup(cnpj) {
    const primary = $('provSel').value;
    const order = [primary, ...Object.keys(PROVIDERS).filter(k => k !== primary)];
    let lastErr = null, notFoundCount = 0;
    for (const key of order) {
      try {
        const found = await fetchFrom(key, cnpj);
        await enrichWithIbge(found.data);
        return { status: 'ok', ...found };
      } catch (e) {
        if (e.kind === 'cancel') throw e;
        if (e.kind === 'nf') notFoundCount++;
        lastErr = e;
        // Se a fonte principal disse "não encontrado", confirma em só mais uma fonte
        if (notFoundCount >= 2) break;
      }
    }
    return notFoundCount > 0
      ? { status: 'nf', msg: 'CNPJ não encontrado na base' }
      : { status: 'err', msg: lastErr ? lastErr.msg : 'Erro desconhecido' };
  }


  // ---------- IBGE: código do município e coordenadas ----------
  // Lista de municípios (view nivelado) -> mapa "NOME|UF" => código IBGE
  const IBGE_MUNICIPIOS = 'https://servicodados.ibge.gov.br/api/v1/localidades/municipios?view=nivelado';
  const IBGE_MALHA = id => `https://servicodados.ibge.gov.br/api/v3/malhas/municipios/${id}`;
  let municipiosPromise = null;
  const centroidCache = new Map(); // código IBGE -> Promise<{lat, lng} | null>

  const normName = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase().replace(/[^A-Z0-9]/g, '');

  function loadMunicipios() {
    if (!municipiosPromise) {
      municipiosPromise = (async () => {
        const res = await fetch(IBGE_MUNICIPIOS);
        if (!res.ok) throw new Error(`IBGE HTTP ${res.status}`);
        const list = await res.json();
        const map = new Map();
        list.forEach(m => map.set(`${normName(m['municipio-nome'])}|${m['UF-sigla']}`, String(m['municipio-id'])));
        return map;
      })().catch(err => { municipiosPromise = null; throw err; });
    }
    return municipiosPromise;
  }

  // Centroide de um polígono/multipolígono GeoJSON (média ponderada pela área)
  function geoCentroid(geojson) {
    let A = 0, cx = 0, cy = 0;
    const ring = pts => {
      for (let i = 0, j = pts.length - 1; i < pts.length; j = i++) {
        const [x0, y0] = pts[j], [x1, y1] = pts[i];
        const f = x0 * y1 - x1 * y0;
        A += f; cx += (x0 + x1) * f; cy += (y0 + y1) * f;
      }
    };
    const geom = g => {
      if (!g) return;
      if (g.type === 'Polygon') g.coordinates.forEach(ring);
      else if (g.type === 'MultiPolygon') g.coordinates.forEach(p => p.forEach(ring));
      else if (g.type === 'GeometryCollection') g.geometries.forEach(geom);
    };
    const feats = geojson.type === 'FeatureCollection' ? geojson.features : [geojson.type === 'Feature' ? geojson : { geometry: geojson }];
    feats.forEach(f => geom(f.geometry));
    if (!A) return null;
    A /= 2;
    return { lng: cx / (6 * A), lat: cy / (6 * A) };
  }

  function getCentroid(id) {
    if (!centroidCache.has(id)) {
      centroidCache.set(id, (async () => {
        // 1ª opção: metadados da malha, que já trazem o centroide calculado pelo IBGE
        try {
          const res = await fetch(`${IBGE_MALHA(id)}/metadados`);
          if (res.ok) {
            const j = await res.json();
            const c = (Array.isArray(j) ? j[0] : j)?.centroide;
            if (c && c.latitude != null && c.longitude != null) return { lat: +c.latitude, lng: +c.longitude };
          }
        } catch (e) { /* tenta a malha */ }
        // 2ª opção: baixa a malha do município e calcula o centroide
        try {
          const res = await fetch(`${IBGE_MALHA(id)}?formato=application/vnd.geo+json&qualidade=minima`);
          if (res.ok) return geoCentroid(await res.json());
        } catch (e) { /* sem coordenada */ }
        return null;
      })());
    }
    return centroidCache.get(id);
  }

  async function enrichWithIbge(a) {
    a.codigo_ibge = ''; a.latitude = ''; a.longitude = ''; a.geo_status = '';
    try {
      const map = await loadMunicipios();
      a.codigo_ibge = map.get(`${normName(a.municipio)}|${a.uf}`) || '';
      if (!a.codigo_ibge && a.codigo_ibge_api) a.codigo_ibge = a.codigo_ibge_api; // nomes grafados diferente na Receita
    } catch (e) {
      a.codigo_ibge = a.codigo_ibge_api || '';
    }
    if (!a.codigo_ibge) { a.geo_status = 'Município não localizado no IBGE'; return; }
    const c = await getCentroid(a.codigo_ibge);
    if (c) {
      a.latitude = Math.round(c.lat * 1e6) / 1e6;
      a.longitude = Math.round(c.lng * 1e6) / 1e6;
      a.geo_status = 'Centroide do município (IBGE)';
    } else a.geo_status = 'Coordenada indisponível no IBGE';
  }

  // ---------- Execução ----------
  $('btnStart').addEventListener('click', () => run(false));
  $('btnRetry').addEventListener('click', () => run(true));
  $('btnPause').addEventListener('click', () => {
    state.paused = !state.paused;
    $('btnPause').textContent = state.paused ? 'Retomar' : 'Pausar';
    setStatus(state.paused ? 'Consulta pausada.' : 'Retomando…');
  });
  $('btnStop').addEventListener('click', () => { state.cancelled = true; state.paused = false; setStatus('Cancelando…'); });

  async function run(onlyErrors) {
    const unique = [...new Set(state.items.filter(i => i.valid).map(i => i.cnpj))];
    const queue = unique.filter(c => {
      const r = state.results.get(c);
      return onlyErrors ? r && r.status === 'err' : !r || r.status === 'wait';
    });
    if (!queue.length) return;
    if (!onlyErrors) unique.forEach(c => { if (!state.results.has(c)) state.results.set(c, { status: 'wait' }); });

    Object.assign(state, { running: true, paused: false, cancelled: false, startedAt: Date.now(), doneThisRun: 0 });
    toggleControls(true);
    setStatus('Carregando lista de municípios do IBGE…');
    try { await loadMunicipios(); } catch (e) { setStatus('Não foi possível carregar a lista do IBGE agora. As coordenadas serão tentadas pelo código do município.'); }
    $('progress').classList.remove('hidden');
    $('results').classList.remove('hidden');
    renderTable();

    const conc = Math.min(5, Math.max(1, +$('conc').value || 1));
    const delay = Math.max(0, +$('delay').value || 0);
    const total = queue.length;
    let idx = 0;

    const worker = async () => {
      while (idx < queue.length && !state.cancelled) {
        while (state.paused && !state.cancelled) await sleep(250);
        if (state.cancelled) break;
        const cnpj = queue[idx++];
        setStatus(`Consultando ${fmtCnpj(cnpj)}…`);
        try {
          state.results.set(cnpj, await lookup(cnpj));
        } catch (e) {
          if (e.kind === 'cancel') break;
          state.results.set(cnpj, { status: 'err', msg: 'Erro inesperado' });
        }
        state.doneThisRun++;
        updateProgress(total);
        scheduleRender();
        if (delay) await sleep(delay);
      }
    };
    await Promise.all(Array.from({ length: conc }, worker));

    state.running = false;
    toggleControls(false);
    updateProgress(total);
    renderTable();
    setStatus(state.cancelled ? 'Consulta cancelada. Você pode exportar o que já foi consultado.' : 'Consulta concluída.');
  }

  function toggleControls(running) {
    $('btnStart').disabled = running;
    $('btnPause').disabled = !running;
    $('btnStop').disabled = !running;
    $('btnPause').textContent = 'Pausar';
    $('sheetSel').disabled = $('colSel').disabled = $('provSel').disabled = running;
    $('btnExport').disabled = running || ![...state.results.values()].some(r => r.status !== 'wait');
    $('btnRetry').disabled = running || ![...state.results.values()].some(r => r.status === 'err');
    $('file').disabled = running;
  }

  function updateProgress(totalThisRun) {
    const all = [...state.results.values()];
    const done = all.filter(r => r.status !== 'wait').length;
    const total = all.length;
    $('pCount').innerHTML = `${done.toLocaleString('pt-BR')} <small>de ${total.toLocaleString('pt-BR')} consultados</small>`;
    $('pBar').style.width = total ? `${(done / total) * 100}%` : '0';
    $('sOk').textContent = all.filter(r => r.status === 'ok').length.toLocaleString('pt-BR');
    $('sNf').textContent = all.filter(r => r.status === 'nf').length.toLocaleString('pt-BR');
    $('sErr').textContent = all.filter(r => r.status === 'err').length.toLocaleString('pt-BR');
    const elapsed = Date.now() - state.startedAt;
    const remaining = totalThisRun - state.doneThisRun;
    if (state.running && state.doneThisRun > 0 && remaining > 0) {
      const s = Math.round((elapsed / state.doneThisRun) * remaining / 1000);
      $('sEta').textContent = s >= 60 ? `${Math.floor(s / 60)} min ${s % 60}s` : `${s}s`;
    } else $('sEta').textContent = '—';
  }
  const setStatus = t => { $('statusLine').textContent = t; };

  // ---------- Tabela ----------
  let renderPending = false;
  function scheduleRender() {
    if (renderPending) return;
    renderPending = true;
    setTimeout(() => { renderPending = false; renderTable(); }, 300);
  }

  document.querySelectorAll('.filters button').forEach(b => b.addEventListener('click', () => {
    state.filter = b.dataset.f;
    document.querySelectorAll('.filters button').forEach(x => x.setAttribute('aria-pressed', x === b));
    renderTable();
  }));

  function rowStatus(it) {
    if (!it.valid) return { status: 'inv', msg: 'CNPJ inválido' };
    return state.results.get(it.cnpj) || { status: 'idle' };
  }

  function renderTable() {
    const MAX = 1000;
    const rows = state.items.filter(it => {
      const s = rowStatus(it).status;
      if (state.filter === 'ok') return s === 'ok';
      if (state.filter === 'fail') return ['nf', 'err', 'inv'].includes(s);
      return true;
    });
    if (!rows.length) {
      $('tbody').innerHTML = `<tr><td colspan="8" class="empty">${state.filter === 'all' ? 'Nenhum CNPJ encontrado na coluna selecionada.' : 'Nenhum resultado neste filtro.'}</td></tr>`;
      return;
    }
    const html = rows.slice(0, MAX).map(it => {
      const r = rowStatus(it);
      const cnpjCell = `<td class="cnpj">${esc(it.valid ? fmtCnpj(it.cnpj) : String(it.raw))}</td>`;
      if (r.status === 'ok') {
        const a = r.data;
        const sit = a.situacao || '—';
        const cls = /ATIVA/i.test(sit) ? 'ok' : 'warn';
        return `<tr><td>${it.row}</td>${cnpjCell}<td><span class="chip ${cls}">${esc(sit)}</span></td>
          <td>${esc(a.razao_social)}${a.nome_fantasia ? `<div class="sub">${esc(a.nome_fantasia)}</div>` : ''}</td>
          <td class="addr">${esc(fullAddress({ ...a, municipio: '', uf: '', cep: '' }))}<div class="sub">Fonte: ${esc(r.source)}</div></td>
          <td>${esc(a.municipio)}${a.uf ? '/' + esc(a.uf) : ''}</td><td style="white-space:nowrap">${esc(a.cep)}</td>
          <td style="white-space:nowrap">${a.latitude !== '' ? `${a.latitude}, ${a.longitude}<div class="sub">IBGE ${esc(a.codigo_ibge)}</div>` : `<span class="sub">${esc(a.geo_status || '—')}</span>`}</td></tr>`;
      }
      const chip = {
        inv: '<span class="chip err">CNPJ inválido</span>',
        nf: '<span class="chip warn">Não encontrado</span>',
        err: `<span class="chip err">Erro</span>`,
        wait: '<span class="chip wait">Na fila</span>',
        idle: '<span class="chip wait">Aguardando</span>'
      }[r.status];
      const msg = r.status === 'err' ? esc(r.msg) : '';
      return `<tr><td>${it.row}</td>${cnpjCell}<td>${chip}</td><td colspan="5" class="sub">${msg}</td></tr>`;
    }).join('');
    const more = rows.length > MAX ? `<tr><td colspan="8" class="empty">Mostrando ${MAX.toLocaleString('pt-BR')} de ${rows.length.toLocaleString('pt-BR')} linhas. A exportação inclui todas.</td></tr>` : '';
    $('tbody').innerHTML = html + more;
  }

  // ---------- Exportação ----------
  $('btnExport').addEventListener('click', () => {
    const extra = ['CNPJ (formatado)', 'Status da consulta', 'Razão social', 'Nome fantasia', 'Situação cadastral', 'Data da situação',
      'Tipo de logradouro', 'Logradouro', 'Número', 'Complemento', 'Bairro', 'Município', 'UF', 'CEP', 'Endereço completo', 'Telefone', 'E-mail', 'Fonte', 'Código IBGE do município', 'Latitude', 'Longitude', 'Precisão da coordenada'];
    const statusLabel = { ok: 'Encontrado', nf: 'Não encontrado', err: 'Erro', inv: 'CNPJ inválido', wait: 'Não consultado', idle: 'Não consultado' };
    const out = [state.header.concat(extra)];
    state.rows.forEach((r, i) => {
      const base = state.header.map((_, c) => r[c] !== undefined ? r[c] : '');
      const raw = r[state.colIdx];
      if (raw === '' || raw === null || raw === undefined) { out.push(base.concat(extra.map(() => ''))); return; }
      const cnpj = normalizeCnpj(raw);
      const valid = isValidCnpj(cnpj);
      const res = valid ? (state.results.get(cnpj) || { status: 'idle' }) : { status: 'inv' };
      const a = res.status === 'ok' ? res.data : null;
      out.push(base.concat([
        valid ? fmtCnpj(cnpj) : String(raw),
        statusLabel[res.status] + (res.status === 'err' && res.msg ? ` (${res.msg})` : ''),
        a ? a.razao_social : '', a ? a.nome_fantasia : '', a ? a.situacao : '', a ? a.data_situacao : '',
        a ? a.tipo_logradouro : '', a ? a.logradouro : '', a ? a.numero : '', a ? a.complemento : '',
        a ? a.bairro : '', a ? a.municipio : '', a ? a.uf : '', a ? a.cep : '',
        a ? fullAddress(a) : '', a ? a.telefone : '', a ? a.email : '', a ? res.source : '',
        a ? a.codigo_ibge : '', a ? a.latitude : '', a ? a.longitude : '', a ? a.geo_status : ''
      ]));
    });
    const ws = XLSX.utils.aoa_to_sheet(out);
    ws['!cols'] = out[0].map((h, i) => ({ wch: h === 'Endereço completo' ? 60 : Math.min(40, Math.max(12, String(h).length + 2)) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Endereços');
    const baseName = state.fileName.replace(/\.[^.]+$/, '');
    XLSX.writeFile(wb, `${baseName}_enderecos.xlsx`);
  });

  function esc(s) {
    return String(s ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  }
})();
