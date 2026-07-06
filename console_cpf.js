// ====================================================================
//  COLAR NO CONSOLE DO NAVEGADOR (F12 > Console)
//  Processa UM CPF: pesquisa, le todas as paginas, soma por competencia
// ====================================================================

(async () => {
  const CPF = '16997155480';
  const MAX_WAIT_MS = 20000;
  const POLL_MS = 250;
  let datasSetadas = false;

  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const digits = s => (s || '').toString().replace(/\D/g, '');
  const fmtCpf = d => d.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4');
  const parseBR = s => { s = (s || '').toString().trim(); if (!s) return 0; s = s.replace(/\./g, '').replace(',', '.'); const n = parseFloat(s); return isNaN(n) ? 0 : n; };
  const fmtBR = n => Number(n).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  const getCpfInput = () => document.querySelector('input[name="cpfTrabalhador"]');
  const getPesquisar = () => [...document.querySelectorAll('button')].find(b => b.textContent.trim().toLowerCase() === 'pesquisar');
  const getExpandir = () => [...document.querySelectorAll('button')].find(b => b.textContent.trim().toLowerCase() === 'expandir pesquisa');
  const isLoading = () => { const l = document.querySelector('br-loading'); return !!(l && l.querySelector('*')); };
  const indicesText = () => { const el = document.querySelector('.indices'); return el ? el.textContent.trim() : ''; };

  async function expandirPesquisa() {
    const btn = getExpandir();
    if (btn) { btn.click(); await sleep(300); console.log('Painel expandido.'); }
  }

  async function selectNgOption(labelText, value) {
    const label = [...document.querySelectorAll('br-label label')].find(l => l.textContent.trim() === labelText);
    if (!label) { console.warn('Label nao encontrada:', labelText); return false; }
    const wrapper = label.closest('br-select') || label.closest('.brx-input-wrapper');
    const ngSelect = wrapper.querySelector('ng-select');
    if (!ngSelect) { console.warn('ng-select nao encontrado para:', labelText); return false; }
    ngSelect.querySelector('.ng-select-container').click();
    await sleep(400);
    const input = ngSelect.querySelector('.ng-input input');
    if (input) {
      const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
      nativeSetter.call(input, value);
      input.dispatchEvent(new Event('input', { bubbles: true }));
    }
    await sleep(500);
    const opts = document.querySelectorAll('.ng-option');
    for (const opt of opts) {
      const lbl = opt.querySelector('.ng-option-label');
      if (lbl && lbl.textContent.trim() === value) { opt.click(); await sleep(300); return true; }
    }
    console.warn('Opcao nao encontrada no dropdown:', value);
    return false;
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
    while (Date.now() - t0 < MAX_WAIT_MS) {
      if (!isLoading()) {
        if (firstRowCpf() === cpf) return 'ok';
        if (readRows().length === 0) {
          await sleep(300);
          if (!isLoading() && readRows().length === 0 && firstRowCpf() === '') return 'vazio';
        }
      }
      await sleep(POLL_MS);
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
      while (Date.now() - t0 < MAX_WAIT_MS) {
        if (!isLoading() && indicesText() !== before && firstRowCpf() === cpf) break;
        await sleep(POLL_MS);
      }
      all = all.concat(readRows());
    }
    return all;
  }

  // --- INICIO ---
  console.log('%c=== FGTS Digital - Automacao CPF ===', 'font-size:14px;font-weight:bold;color:#10b981');
  const cpfFmt = fmtCpf(CPF);
  console.log('CPF:', cpfFmt);

  // Expande o painel de pesquisa PRIMEIRO (senao o campo CPF nem existe no DOM)
  await expandirPesquisa();

  const input = getCpfInput();
  const btn = getPesquisar();
  if (!input || !btn) {
    return console.error('ERRO: Campo CPF ou botao Pesquisar nao encontrados. Certifique-se de estar na tela de pesquisa de debitos.');
  }

  // Seleciona datas de competencia (so na primeira vez ou apos reload)
  if (!datasSetadas) {
    console.log('Selecionando datas de competencia...');
    await selectNgOption('Inicial', '09/2025');
    await selectNgOption('Final', '04/2026');
    datasSetadas = true;
    await sleep(300);
  }

  // Digita o CPF e pesquisa
  setNativeValue(input, cpfFmt);
  input.blur();
  await sleep(120);
  btn.click();
  console.log('Pesquisando', cpfFmt, '...');

  const status = await waitResults(CPF);
  if (status === 'vazio') { console.log('%cSEM DEBITOS para este CPF', 'color:#f59e0b'); return; }
  if (status === 'timeout') { console.error('TIMEOUT aguardando resultados'); return; }

  // Marca "selecionar todos"
  const chk = document.querySelector('#selecionar-todos');
  if (chk && !chk.checked) { chk.click(); await sleep(120); }

  // Clica em "Adicionar à guia"
  const btnAdicionar = [...document.querySelectorAll('button')].find(b => { const t = b.textContent.trim().toLowerCase(); return t === 'adicionar à guia' || t === 'adicionar a guia'; });
  if (btnAdicionar) { btnAdicionar.click(); await sleep(300); console.log('Adicionado a guia.'); }

  const map = getColMap();
  console.log('Colunas: apuracao=' + map.apur + ' total=' + map.total + ' (total de colunas=' + map.nCols + ')');

  const linhas = await coletarPaginas(CPF);
  console.log('Linhas lidas:', linhas.length);

  // Soma por competencia
  const porData = {};
  for (const row of linhas) {
    const data = (map.apur >= 0 ? row[map.apur] : '') || '(sem data)';
    const val = parseBR(map.total >= 0 ? row[map.total] : '');
    porData[data] = (porData[data] || 0) + val;
  }

  let totalGeral = 0;
  console.log('');
  console.log('%c=== RESULTADOS POR COMPETENCIA ===', 'font-weight:bold;color:#10b981');
  console.log('Competencia       | Valor');
  console.log('------------------|----------');
  for (const d of Object.keys(porData).sort()) {
    totalGeral += porData[d];
    console.log(d.padEnd(18) + '| ' + fmtBR(porData[d]));
  }
  console.log('------------------|----------');
  console.log('TOTAL GERAL'.padEnd(18) + '| ' + fmtBR(totalGeral));
  console.log('');
  console.log('Linhas de debito:', linhas.length);
  console.log('%c=== CONCLUIDO ===', 'font-weight:bold;color:#10b981');

  // Copia o resultado como CSV para area de transferencia
  const csv = Object.entries(porData).map(([d, v]) => `${CPF},${d},${v.toFixed(2)}`).join('\n');
  console.log('');
  console.log('%cCSV para copiar (CTRL+C):', 'font-weight:bold');
  console.log('CPF,Competencia,Valor');
  console.log(csv);

  // Clica em "Avancar" para proxima etapa
  await sleep(500);
  const btnAvancar = [...document.querySelectorAll('button')].find(b => b.textContent.trim().toLowerCase() === 'avançar');
  if (btnAvancar) {
    btnAvancar.click();
    console.log('%cClique em "Avançar" executado.', 'color:#10b981');
  } else {
    console.warn('Botao "Avançar" nao encontrado.');
  }
})();
