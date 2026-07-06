// ==UserScript==
// @name         FGTS Digital - Automacao CPF (planilha -> soma Totais -> planilha)
// @namespace    marcosprzn.fgts
// @version      1.0.0
// @description  Le uma planilha de CPFs, pesquisa cada um no FGTS Digital, soma os Totais por competencia, lida com o cronometro de sessao (recarrega e retoma) e gera planilha + log.
// @match        https://*.sistema.gov.br/*
// @match        https://fgtsdigital.*/*
// @require      https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';
  if (window.top !== window.self) return; // nao roda em iframes

  /* ============================ CONFIG ============================ */
  const CFG = {
    COL_CPF: 0,        // coluna A = CPF
    COL_NOME: 2,       // coluna C = Nome
    COL_FILTRO: 5,     // coluna F = filtro
    VALOR_FILTRO: 'D', // processa apenas linhas com "D" na coluna F ('' desativa)
    RELOAD_AT_SEC: 5,  // recarrega quando faltar <= X segundos no cronometro
    MAX_WAIT_MS: 20000,
    POLL_MS: 250,
    PAUSE_MS: 300,
    MAX_LOG_LINHAS: 6000,
    LS_KEY: 'FGTS_AUTO_STATE'
  };

  /* ============================ ESTADO ============================ */
  let state = {
    running: false,
    index: 0,
    limite: 0,          // quantos processar a partir do inicio (0 = todos)
    registros: [],      // [[cpf, nome], ...]
    resultados: [],     // {i, cpf, nome, competencia, qtd, soma}
    semDebito: [],      // {i, cpf, nome, obs}
    logs: [],
    iniciadoEm: ''
  };
  let curIndex = 0;
  let reloadArmed = false;
  let panel, elStatus, elProg, elLog;

  /* ============================ HELPERS ============================ */
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const digits = s => (s || '').toString().replace(/\D/g, '');
  const fmtCpf = d => d.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4');
  const parseBR = s => { s = (s || '').toString().trim(); if (!s) return 0; s = s.replace(/\./g, '').replace(',', '.'); const n = parseFloat(s); return isNaN(n) ? 0 : n; };
  const fmtBR = n => Number(n).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const stamp = () => new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
  const nowISO = () => new Date().toISOString().slice(0, 23).replace('T', ' ');

  function saveState() {
    try { localStorage.setItem(CFG.LS_KEY, JSON.stringify(state)); }
    catch (e) { console.warn('Falha ao salvar estado:', e); }
  }
  function loadState() {
    try {
      const raw = localStorage.getItem(CFG.LS_KEY);
      if (raw) state = Object.assign(state, JSON.parse(raw));
    } catch (e) { console.warn('Falha ao ler estado:', e); }
  }
  function limparState() {
    localStorage.removeItem(CFG.LS_KEY);
    state = { running: false, index: 0, limite: 0, registros: [], resultados: [], semDebito: [], logs: [], iniciadoEm: '' };
    curIndex = 0; reloadArmed = false;
    log('Estado limpo.');
    render();
  }

  function log(msg, tipo) {
    const linha = '[' + nowISO() + '] ' + (tipo ? '[' + tipo + '] ' : '') + msg;
    state.logs.push(linha);
    if (state.logs.length > CFG.MAX_LOG_LINHAS) state.logs.splice(0, state.logs.length - CFG.MAX_LOG_LINHAS);
    (tipo === 'ERRO' ? console.error : tipo === 'AVISO' ? console.warn : console.log)(linha);
    if (elLog) {
      elLog.textContent = state.logs.slice(-25).join('\n');
      elLog.scrollTop = elLog.scrollHeight;
    }
  }

  /* ==================== SELETORES DA PAGINA ==================== */
  const getCpfInput = () => document.querySelector('input[name="cpfTrabalhador"]');
  const getPesquisar = () => [...document.querySelectorAll('button')].find(b => b.textContent.trim().toLowerCase() === 'pesquisar');
  const isLoading = () => { const l = document.querySelector('br-loading'); return !!(l && l.querySelector('*')); };
  const indicesText = () => { const el = document.querySelector('.indices'); return el ? el.textContent.trim() : ''; };

  function clockSeconds() {
    const el = document.querySelector('.clock');
    if (!el) return -1;
    const parts = el.textContent.trim().split(':').map(n => parseInt(n, 10));
    if (!parts.length || parts.some(isNaN)) return -1;
    let s = 0; for (const p of parts) s = s * 60 + p; return s;
  }

  function setNativeValue(el, value) {
    const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
    setter.call(el, value);
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  }

  function getColMap() {
    const heads = [...document.querySelectorAll('datatable-header-cell')];
    let apur = -1, total = -1;
    heads.forEach((h, i) => {
      const t = h.textContent.replace(/\s+/g, ' ').trim().toLowerCase();
      const title = (h.querySelector('[title]') ? h.querySelector('[title]').getAttribute('title') : '').toLowerCase();
      if (apur < 0 && (title.includes('apura') || t.includes('apura'))) apur = i;
      if (total < 0 && (t === 'total' || title === 'total')) total = i;
    });
    return { apur, total, nCols: heads.length };
  }

  const readRows = () =>
    [...document.querySelectorAll('datatable-body-row')]
      .map(r => [...r.querySelectorAll('datatable-body-cell')]
        .map(c => c.textContent.replace(/\s+/g, ' ').trim()));

  function firstRowCpf() {
    const rows = readRows();
    if (!rows.length) return '';
    for (const cell of rows[0]) { if (digits(cell).length === 11) return digits(cell); }
    return '';
  }

  async function waitResults(cpf) {
    const t0 = Date.now();
    await sleep(350);
    while (Date.now() - t0 < CFG.MAX_WAIT_MS) {
      if (!isLoading()) {
        if (firstRowCpf() === cpf) return 'ok';
        if (readRows().length === 0) {
          await sleep(300);
          if (!isLoading() && readRows().length === 0 && firstRowCpf() === '') return 'vazio';
        }
      }
      await sleep(CFG.POLL_MS);
    }
    return 'timeout';
  }

  async function coletarPaginas(cpf) {
    let all = readRows();
    let guard = 0;
    while (guard++ < 60) {
      const next = document.querySelector('#btn-next-page');
      if (!next || next.disabled) break;
      const before = indicesText();
      next.click();
      const t0 = Date.now(); await sleep(200);
      while (Date.now() - t0 < CFG.MAX_WAIT_MS) {
        if (!isLoading() && indicesText() !== before && firstRowCpf() === cpf) break;
        await sleep(CFG.POLL_MS);
      }
      all = all.concat(readRows());
    }
    return all;
  }

  async function waitTelaPesquisa(maxMs) {
    const t0 = Date.now();
    while (Date.now() - t0 < (maxMs || 120000)) {
      if (getCpfInput() && getPesquisar()) return true;
      setStatus('Aguardando a tela de pesquisa (campo CPF)...');
      await sleep(500);
    }
    return false;
  }

  /* ==================== LEITURA DA PLANILHA ==================== */
  function lerPlanilha(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(new Error('Erro ao ler o arquivo.'));
      reader.onload = e => {
        try {
          if (typeof XLSX === 'undefined') throw new Error('Biblioteca XLSX (SheetJS) nao carregou. Verifique o @require do userscript / conexao.');
          const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
          resolve(rows);
        } catch (err) { reject(err); }
      };
      reader.readAsArrayBuffer(file);
    });
  }

  function extrairRegistros(rows) {
    const seen = new Set();
    const registros = [];
    let totalLinhas = 0, comFiltro = 0, invalidos = 0, duplicados = 0;
    for (const row of rows) {
      const bruto = row[CFG.COL_CPF];
      const d0 = digits(bruto);
      if (!d0) continue; // linha sem CPF (ex: cabecalho)
      totalLinhas++;
      const filtro = (row[CFG.COL_FILTRO] == null ? '' : row[CFG.COL_FILTRO].toString().trim().toUpperCase());
      if (CFG.VALOR_FILTRO && filtro !== CFG.VALOR_FILTRO.toUpperCase()) continue;
      comFiltro++;
      const cpf = d0.padStart(11, '0');
      if (cpf.length !== 11) { invalidos++; continue; }
      if (seen.has(cpf)) { duplicados++; continue; }
      seen.add(cpf);
      const nome = (row[CFG.COL_NOME] == null ? '' : row[CFG.COL_NOME].toString().trim());
      registros.push([cpf, nome]);
    }
    log('Planilha lida: ' + totalLinhas + ' linhas com CPF | filtro coluna F=="' + CFG.VALOR_FILTRO + '": ' + comFiltro +
        ' | duplicados removidos: ' + duplicados + ' | invalidos: ' + invalidos + ' | FINAL: ' + registros.length + ' CPFs');
    return registros;
  }

  /* ==================== FLUXO PRINCIPAL ==================== */
  async function iniciarComArquivo(file) {
    try {
      log('Arquivo selecionado: ' + file.name + ' (' + file.size + ' bytes)');
      const rows = await lerPlanilha(file);
      const registros = extrairRegistros(rows);
      if (!registros.length) { alert('Nenhum CPF valido encontrado na planilha (apos o filtro coluna F=="' + CFG.VALOR_FILTRO + '").'); return; }

      let limite = registros.length;
      const resp = prompt('Foram encontrados ' + registros.length + ' CPFs.\nQuantos deseja processar?\n(ENTER = todos)');
      if (resp !== null && resp.trim() !== '') {
        const n = parseInt(resp.replace(/\D/g, ''), 10);
        if (!isNaN(n) && n > 0) limite = Math.min(n, registros.length);
      }
      log('Quantidade a processar: ' + limite + ' de ' + registros.length);

      state = {
        running: true, index: 0, limite: limite, registros: registros,
        resultados: [], semDebito: [], logs: state.logs || [], iniciadoEm: nowISO()
      };
      log('=== INICIO DA AUTOMACAO === URL: ' + location.href);
      log('UserAgent: ' + navigator.userAgent);
      saveState();
      render();
      const ok = await waitTelaPesquisa();
      if (!ok) { log('Tela de pesquisa nao encontrada. Abra a tela de pesquisa de debitos e clique em Retomar.', 'AVISO'); state.running = false; saveState(); render(); return; }
      await runLoop(0);
    } catch (err) {
      log('Erro ao iniciar: ' + err.message + '\n' + (err.stack || ''), 'ERRO');
      alert('Erro ao iniciar: ' + err.message);
    }
  }

  async function retomar() {
    if (!state.registros.length) { log('Nada para retomar (sem registros salvos).', 'AVISO'); return; }
    state.running = true; reloadArmed = false; saveState(); render();
    log('=== RETOMANDO apos reload/pausa no index ' + state.index + ' === URL: ' + location.href);
    const ok = await waitTelaPesquisa();
    if (!ok) { log('Tela de pesquisa nao encontrada ao retomar.', 'AVISO'); return; }
    await runLoop(state.index);
  }

  async function runLoop(startIdx) {
    const limite = state.limite && state.limite > 0 ? Math.min(state.limite, state.registros.length) : state.registros.length;
    for (let i = startIdx; i < limite; i++) {
      if (!state.running) { log('Loop interrompido (running=false) no index ' + i, 'AVISO'); return; }
      curIndex = i;
      state.index = i;
      const [cpf, nome] = state.registros[i];
      const cpfFmt = fmtCpf(cpf);
      setStatus('Processando ' + (i + 1) + '/' + limite + '  ' + cpfFmt);
      setProgress(i, limite);
      try {
        // remove qualquer resultado parcial anterior deste indice (caso de redo pos-reload)
        state.resultados = state.resultados.filter(r => r.i !== i);
        state.semDebito = state.semDebito.filter(s => s.i !== i);

        const input = getCpfInput();
        if (!input) { log('[' + i + '] Campo CPF sumiu. Aguardando tela...', 'AVISO'); const ok = await waitTelaPesquisa(); if (!ok) { state.running = false; saveState(); return; } }
        setNativeValue(getCpfInput(), cpfFmt);
        getCpfInput().blur();
        await sleep(120);
        getPesquisar().click();
        log('[' + i + '] Pesquisando ' + cpfFmt + ' ' + nome);

        const status = await waitResults(cpf);
        if (status === 'vazio') {
          state.semDebito.push({ i, cpf, nome, obs: 'sem debitos' });
          log('[' + i + '] ' + cpfFmt + ' -> SEM DEBITOS');
          saveState(); await sleep(CFG.PAUSE_MS); continue;
        }
        if (status === 'timeout') {
          state.semDebito.push({ i, cpf, nome, obs: 'TIMEOUT' });
          log('[' + i + '] ' + cpfFmt + ' -> TIMEOUT (revisar manualmente)', 'AVISO');
          saveState(); await sleep(CFG.PAUSE_MS); continue;
        }

        const chk = document.querySelector('#selecionar-todos');
        if (chk && !chk.checked) { chk.click(); await sleep(120); }

        const { apur, total, nCols } = getColMap();
        if (i === startIdx) log('Mapa de colunas -> apuracao=' + apur + ' total=' + total + ' (nCols=' + nCols + ')');
        const linhas = await coletarPaginas(cpf);

        const porData = {};
        for (const row of linhas) {
          const data = (apur >= 0 ? row[apur] : '') || '(sem data)';
          const val = parseBR(total >= 0 ? row[total] : '');
          porData[data] = (porData[data] || 0) + val;
        }
        const datas = Object.keys(porData);
        let somaCpf = 0;
        for (const d of datas) {
          somaCpf += porData[d];
          state.resultados.push({ i, cpf: cpfFmt, nome, competencia: d, qtd: linhas.length, soma: porData[d] });
        }
        log('[' + i + '] ' + cpfFmt + ' -> ' + linhas.length + ' debito(s) | ' +
            datas.map(d => d + '=' + fmtBR(porData[d])).join(' ; ') + ' | TOTAL=' + fmtBR(somaCpf));
        saveState();
      } catch (err) {
        state.semDebito.push({ i, cpf, nome, obs: 'ERRO: ' + err.message });
        log('[' + i + '] ' + cpfFmt + ' ERRO: ' + err.message + '\n' + (err.stack || ''), 'ERRO');
        saveState();
      }
      await sleep(CFG.PAUSE_MS);
    }
    finalizar();
  }

  function finalizar() {
    state.running = false;
    state.index = state.limite && state.limite > 0 ? Math.min(state.limite, state.registros.length) : state.registros.length;
    saveState();
    setStatus('CONCLUIDO');
    setProgress(1, 1);
    log('=== CONCLUIDO === ' + state.resultados.length + ' linhas de resultado | ' + state.semDebito.length + ' sem debito/erros');
    baixarResultado();
    baixarLogs();
  }

  /* ==================== GERAR SAIDAS ==================== */
  function baixarResultado() {
    try {
      const wb = XLSX.utils.book_new();
      const rows = state.resultados.map(r => ({
        Indice: r.i, CPF: r.cpf, Nome: r.nome, Competencia: r.competencia,
        Qtd_Debitos: r.qtd, Soma_Total: Number(Number(r.soma).toFixed(2))
      }));
      const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{ Indice: '', CPF: '', Nome: '', Competencia: '', Qtd_Debitos: '', Soma_Total: '' }]);
      XLSX.utils.book_append_sheet(wb, ws, 'Resultados');

      const rows2 = state.semDebito.map(s => ({ Indice: s.i, CPF: fmtCpf(s.cpf), Nome: s.nome, Obs: s.obs || 'sem debitos' }));
      const ws2 = XLSX.utils.json_to_sheet(rows2.length ? rows2 : [{ Indice: '', CPF: '', Nome: '', Obs: '' }]);
      XLSX.utils.book_append_sheet(wb, ws2, 'Sem_Debito_Erros');

      XLSX.writeFile(wb, 'fgts_resultados_' + stamp() + '.xlsx');
      log('Planilha de resultados gerada.');
    } catch (err) {
      log('Erro ao gerar planilha: ' + err.message, 'ERRO');
    }
  }

  function baixarLogs() {
    const txt = state.logs.join('\n');
    const blob = new Blob([txt], { type: 'text/plain;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'fgts_logs_' + stamp() + '.log';
    document.body.appendChild(a); a.click(); a.remove();
  }

  /* ==================== CRONOMETRO (RELOAD + RETOMA) ==================== */
  function startClockWatcher() {
    setInterval(() => {
      if (!state.running || reloadArmed) return;
      const s = clockSeconds();
      if (s >= 0 && s <= CFG.RELOAD_AT_SEC) {
        reloadArmed = true;
        state.index = curIndex; // refaz o CPF atual apos o reload
        state.resultados = state.resultados.filter(r => r.i !== curIndex);
        state.semDebito = state.semDebito.filter(x => x.i !== curIndex);
        log('CRONOMETRO em ' + s + 's -> salvando e RECARREGANDO a pagina. Retomo no index ' + curIndex, 'AVISO');
        saveState();
        setTimeout(() => location.reload(), 150);
      }
    }, 500);
  }

  /* ==================== PAINEL / UI ==================== */
  function setStatus(t) { if (elStatus) elStatus.textContent = t; }
  function setProgress(done, total) {
    if (!elProg) return;
    const pct = total ? Math.round((done / total) * 100) : 0;
    elProg.textContent = done + ' / ' + total + '  (' + pct + '%)';
  }
  function render() {
    if (state.running) setStatus('Rodando... index ' + state.index);
    else if (state.registros.length) setStatus('Pausado no index ' + state.index + ' (clique Retomar)');
    else setStatus('Pronto. Selecione a planilha.');
    setProgress(state.index, state.limite && state.limite > 0 ? Math.min(state.limite, state.registros.length) : state.registros.length);
  }

  function buildPanel() {
    panel = document.createElement('div');
    panel.style.cssText = 'position:fixed;top:10px;right:10px;z-index:2147483647;width:330px;font:12px/1.4 Arial,sans-serif;background:#0b3d2e;color:#eafff5;border:1px solid #10b981;border-radius:8px;box-shadow:0 6px 24px rgba(0,0,0,.4);';
    panel.innerHTML =
      '<div style="background:#10b981;color:#04231a;font-weight:bold;padding:6px 10px;border-radius:7px 7px 0 0;cursor:move;">Automacao FGTS Digital</div>' +
      '<div style="padding:8px 10px;">' +
      '  <div id="fgts-status" style="margin-bottom:4px;">Pronto.</div>' +
      '  <div id="fgts-prog" style="margin-bottom:6px;font-weight:bold;">0 / 0</div>' +
      '  <input id="fgts-file" type="file" accept=".xlsx,.xls" style="display:none">' +
      '  <div style="display:flex;flex-wrap:wrap;gap:4px;margin-bottom:6px;">' +
      '    <button id="fgts-sel"  style="flex:1 1 100%;padding:6px;cursor:pointer;background:#10b981;border:0;border-radius:5px;color:#04231a;font-weight:bold;">Selecionar planilha e iniciar</button>' +
      '    <button id="fgts-ret"  style="flex:1;padding:5px;cursor:pointer;">Retomar</button>' +
      '    <button id="fgts-stop" style="flex:1;padding:5px;cursor:pointer;">Parar</button>' +
      '    <button id="fgts-xlsx" style="flex:1;padding:5px;cursor:pointer;">Baixar XLSX</button>' +
      '    <button id="fgts-log"  style="flex:1;padding:5px;cursor:pointer;">Baixar log</button>' +
      '    <button id="fgts-clr"  style="flex:1 1 100%;padding:5px;cursor:pointer;background:#7f1d1d;border:0;border-radius:5px;color:#fff;">Limpar tudo</button>' +
      '  </div>' +
      '  <pre id="fgts-logbox" style="height:150px;overflow:auto;background:#04231a;color:#9be7c4;margin:0;padding:6px;border-radius:5px;white-space:pre-wrap;word-break:break-word;"></pre>' +
      '</div>';
    document.body.appendChild(panel);

    elStatus = panel.querySelector('#fgts-status');
    elProg = panel.querySelector('#fgts-prog');
    elLog = panel.querySelector('#fgts-logbox');
    const fileInput = panel.querySelector('#fgts-file');

    panel.querySelector('#fgts-sel').onclick = () => fileInput.click();
    fileInput.onchange = () => { if (fileInput.files[0]) iniciarComArquivo(fileInput.files[0]); };
    panel.querySelector('#fgts-ret').onclick = () => retomar();
    panel.querySelector('#fgts-stop').onclick = () => { state.running = false; saveState(); log('PARADO pelo usuario no index ' + state.index, 'AVISO'); render(); };
    panel.querySelector('#fgts-xlsx').onclick = () => baixarResultado();
    panel.querySelector('#fgts-log').onclick = () => baixarLogs();
    panel.querySelector('#fgts-clr').onclick = () => { if (confirm('Limpar todo o progresso salvo?')) limparState(); };

    // arrastar painel
    let drag = false, ox = 0, oy = 0;
    const head = panel.firstChild;
    head.onmousedown = e => { drag = true; ox = e.clientX - panel.offsetLeft; oy = e.clientY - panel.offsetTop; };
    document.addEventListener('mousemove', e => { if (drag) { panel.style.left = (e.clientX - ox) + 'px'; panel.style.top = (e.clientY - oy) + 'px'; panel.style.right = 'auto'; } });
    document.addEventListener('mouseup', () => drag = false);
  }

  /* ==================== INIT ==================== */
  function init() {
    loadState();
    buildPanel();
    if (elLog) elLog.textContent = state.logs.slice(-25).join('\n');
    startClockWatcher();
    render();
    log('Script carregado. XLSX ' + (typeof XLSX !== 'undefined' ? 'OK (' + XLSX.version + ')' : 'NAO CARREGOU'));
    if (state.running && state.registros.length) {
      log('Estado running=true detectado -> retomando automaticamente.');
      retomar();
    }
  }

  if (document.readyState === 'complete' || document.readyState === 'interactive') setTimeout(init, 300);
  else window.addEventListener('DOMContentLoaded', () => setTimeout(init, 300));
})();
