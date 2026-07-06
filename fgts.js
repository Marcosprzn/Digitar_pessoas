/* =====================================================================
   FGTS Digital - Automacao (Node + Selenium + Google Chrome)
   ---------------------------------------------------------------------
   Fluxo:
     1) Abre a planilha (dialogo do Windows ou caminho por argumento).
     2) Pergunta quantos CPFs processar (ENTER = todos).
     3) Abre o Chrome no FGTS Digital e ESPERA voce fazer login.
     4) Mostra o botao "INICIAR AUTOMACAO" na pagina; ao clicar, comeca.
     5) Para cada CPF: pesquisa, marca selecionar-todos, soma os Totais
        por competencia (duplicados na mesma competencia sao somados).
     6) Cronometro de sessao: faltando <=5s, recarrega e continua.
     7) No fim gera a planilha de resultados + arquivo de log.

   Uso:  node fgts.js  [caminho\\planilha.xlsx]
   Obs.: No Windows 8 use o Google Chrome 109 (versoes 110+ nao rodam no Win8).
   ===================================================================== */
'use strict';

const fs = require('fs');
const path = require('path');
const readline = require('readline');
const { execFileSync } = require('child_process');

const { Builder } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const XLSX = require('xlsx');

/* ============================ CONFIG ============================ */
const CFG = {
  START_URL: 'https://fgtsdigital.sistema.gov.br/',
  COL_CPF: 0, COL_NOME: 2, COL_FILTRO: 5, VALOR_FILTRO: 'D',
  RELOAD_AT_SEC: 5,
  PAUSE_MS: 300,
  SCRIPT_TIMEOUT_MS: 120000,
  PAGELOAD_TIMEOUT_MS: 120000
};

/* ============================ LOG ============================ */
const STAMP = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
const LOG_PATH = path.join(process.cwd(), 'fgts_log_' + STAMP + '.log');
const logStream = fs.createWriteStream(LOG_PATH, { flags: 'a' });
function log(msg, tipo) {
  const line = '[' + new Date().toISOString().slice(0, 23).replace('T', ' ') + '] ' + (tipo ? '[' + tipo + '] ' : '') + msg;
  (tipo === 'ERRO' ? console.error : tipo === 'AVISO' ? console.warn : console.log)(line);
  try { logStream.write(line + '\n'); } catch (e) {}
}

/* ============================ HELPERS ============================ */
const sleep = ms => new Promise(r => setTimeout(r, ms));
const digits = s => (s == null ? '' : s.toString()).replace(/\D/g, '');
const fmtCpf = d => d.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4');
const fmtBR = n => Number(n).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

function ask(q) {
  return new Promise(res => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question(q, a => { rl.close(); res(a); });
  });
}

/* ---------- selecionar planilha via dialogo do Windows ---------- */
function escolherPlanilha() {
  // 1) argumento de linha de comando
  if (process.argv[2] && fs.existsSync(process.argv[2])) return process.argv[2];
  // 2) dialogo do Windows (PowerShell OpenFileDialog)
  try {
    const ps =
      "Add-Type -AssemblyName System.Windows.Forms | Out-Null;" +
      "$f = New-Object System.Windows.Forms.OpenFileDialog;" +
      "$f.Title='Selecione a planilha de CPFs';" +
      "$f.Filter='Planilhas (*.xlsx;*.xls)|*.xlsx;*.xls|Todos (*.*)|*.*';" +
      "if($f.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){[Console]::Out.Write($f.FileName)}";
    const out = execFileSync('powershell', ['-NoProfile', '-STA', '-Command', ps], { encoding: 'utf8' }).trim();
    if (out && fs.existsSync(out)) return out;
  } catch (e) { log('Dialogo de arquivo falhou: ' + e.message, 'AVISO'); }
  // 3) auto-deteccao de um unico .xlsx na pasta
  const xs = fs.readdirSync(process.cwd()).filter(f => /\.xlsx?$/i.test(f) && !/resultado/i.test(f));
  if (xs.length === 1) return path.join(process.cwd(), xs[0]);
  return null;
}

function lerRegistros(caminho) {
  const wb = XLSX.readFile(caminho);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
  const seen = new Set();
  const registros = [];
  let comCpf = 0, comFiltro = 0, dup = 0;
  for (const row of rows) {
    const d0 = digits(row[CFG.COL_CPF]);
    if (!d0) continue;
    comCpf++;
    const filtro = (row[CFG.COL_FILTRO] == null ? '' : row[CFG.COL_FILTRO].toString().trim().toUpperCase());
    if (CFG.VALOR_FILTRO && filtro !== CFG.VALOR_FILTRO.toUpperCase()) continue;
    comFiltro++;
    const cpf = d0.padStart(11, '0');
    if (cpf.length !== 11) continue;
    if (seen.has(cpf)) { dup++; continue; }
    seen.add(cpf);
    const nome = (row[CFG.COL_NOME] == null ? '' : row[CFG.COL_NOME].toString().trim());
    registros.push([cpf, nome]);
  }
  log('Planilha: ' + comCpf + ' linhas com CPF | filtro F=="' + CFG.VALOR_FILTRO + '": ' + comFiltro +
      ' | duplicados: ' + dup + ' | FINAL: ' + registros.length + ' CPFs');
  return registros;
}

/* ---------- localizar Chrome e chromedriver ---------- */
function acharChrome() {
  const cands = [
    process.env.ProgramFiles + '\\Google\\Chrome\\Application\\chrome.exe',
    process.env['ProgramFiles(x86)'] + '\\Google\\Chrome\\Application\\chrome.exe',
    process.env.LOCALAPPDATA + '\\Google\\Chrome\\Application\\chrome.exe'
  ];
  for (const c of cands) { if (c && fs.existsSync(c)) return c; }
  return null;
}
function acharChromedriver() {
  const cands = [
    path.join(__dirname, 'node_modules', 'chromedriver', 'lib', 'chromedriver', 'chromedriver.exe'),
    path.join(__dirname, 'node_modules', '.bin', 'chromedriver.exe'),
    path.join(__dirname, 'chromedriver.exe')
  ];
  for (const c of cands) { if (fs.existsSync(c)) return c; }
  return null; // deixa o Selenium Manager resolver automaticamente
}

/* =====================================================================
   CODIGO INJETADO NA PAGINA (roda no contexto do Firefox)
   Define window.__fgts.process(cpf) -> Promise<{status, nlinhas, porData}>
   ===================================================================== */
const PAGE_HELPERS = `
window.__fgts = (function(){
  var CFG = { MAX_WAIT_MS: 20000, POLL_MS: 250 };
  var panel = null, elStatus = null, elProg = null, elPause = null, paused = false, datasSetadas = false;
  function sleep(ms){ return new Promise(function(r){ setTimeout(r, ms); }); }
  function digits(s){ return (s||'').toString().replace(/\\D/g,''); }
  function parseBR(s){ s=(s||'').toString().trim(); if(!s) return 0; s=s.replace(/\\./g,'').replace(',','.'); var n=parseFloat(s); return isNaN(n)?0:n; }
  function getCpfInput(){ return document.querySelector('input[name="cpfTrabalhador"]'); }
  function getPesquisar(){ return Array.prototype.slice.call(document.querySelectorAll('button')).find(function(b){ return b.textContent.trim().toLowerCase()==='pesquisar'; }); }
  function getExpandir(){ return Array.prototype.slice.call(document.querySelectorAll('button')).find(function(b){ return b.textContent.trim().toLowerCase()==='expandir pesquisa'; }); }
  function isLoading(){ var l=document.querySelector('br-loading'); return !!(l && l.querySelector('*')); }
  function indicesText(){ var el=document.querySelector('.indices'); return el?el.textContent.trim():''; }
  async function expandirPesquisa(){ if(getCpfInput()) return; var btn=getExpandir(); if(btn){ btn.click(); await sleep(300); } }
  function criarPanel(total){
    if(panel) return;
    panel=document.createElement('div');
    panel.id='fgts-panel';
    panel.style.cssText='position:fixed;top:10px;left:10px;z-index:2147483647;width:280px;font:12px/1.4 Arial,sans-serif;background:#0b3d2e;color:#eafff5;border:1px solid #10b981;border-radius:8px;box-shadow:0 6px 24px rgba(0,0,0,.4);';
    panel.innerHTML='<div style="background:#10b981;color:#04231a;font-weight:bold;padding:6px 10px;border-radius:7px 7px 0 0;">Automacao FGTS</div><div style="padding:8px 10px;"><div id="fgts-p-status" style="margin-bottom:4px;">Aguardando...</div><div id="fgts-p-prog" style="margin-bottom:6px;font-weight:bold;font-size:14px;">0 / '+total+'</div><div style="background:#04231a;height:8px;border-radius:4px;margin-bottom:6px;overflow:hidden;"><div id="fgts-p-bar" style="height:100%;width:0%;background:#10b981;border-radius:4px;transition:width .3s;"></div></div><button id="fgts-p-pause" style="width:100%;padding:6px;cursor:pointer;background:#f59e0b;border:0;border-radius:5px;color:#000;font-weight:bold;">PAUSAR</button></div>';
    document.body.appendChild(panel);
    elStatus=panel.querySelector('#fgts-p-status');
    elProg=panel.querySelector('#fgts-p-prog');
    elPause=panel.querySelector('#fgts-p-pause');
    elPause.onclick=function(){ paused=!paused; elPause.textContent=paused?'CONTINUAR':'PAUSAR'; elPause.style.background=paused?'#10b981':'#f59e0b'; };
  }
  function atualizarPainel(i, total, cpf){
    if(!panel) criarPanel(total);
    if(elStatus) elStatus.textContent=cpf?'Processando: '+cpf:'Aguardando...';
    if(elProg) elProg.textContent=i+' / '+total;
    var bar=document.querySelector('#fgts-p-bar');
    if(bar) bar.style.width=total?Math.round(i/total*100)+'%':'0%';
  }
  async function aguardarSePausado(){
    while(paused){ await sleep(500); }
  }
  async function selectNgOption(labelText, value){
    var labels=Array.prototype.slice.call(document.querySelectorAll('br-label label'));
    var label=labels.find(function(l){ return l.textContent.trim()===labelText; });
    if(!label) return false;
    var wrapper=label.closest('br-select')||label.closest('.brx-input-wrapper');
    var ngSelect=wrapper.querySelector('ng-select');
    if(!ngSelect) return false;
    ngSelect.querySelector('.ng-select-container').click();
    await sleep(400);
    var input=ngSelect.querySelector('.ng-input input');
    if(input){
      var setter=Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype,'value').set;
      setter.call(input, value);
      input.dispatchEvent(new Event('input',{bubbles:true}));
    }
    await sleep(500);
    var opts=document.querySelectorAll('.ng-option');
    for(var i=0;i<opts.length;i++){
      var lbl=opts[i].querySelector('.ng-option-label');
      if(lbl && lbl.textContent.trim()===value){ opts[i].click(); await sleep(300); return true; }
    }
    return false;
  }
  function setNativeValue(el, value){
    var setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype,'value').set;
    setter.call(el, value);
    el.dispatchEvent(new Event('input',{bubbles:true}));
    el.dispatchEvent(new Event('change',{bubbles:true}));
  }
  function getColMap(){
    var heads = Array.prototype.slice.call(document.querySelectorAll('datatable-header-cell'));
    var apur=-1, total=-1;
    heads.forEach(function(h,i){
      var t=h.textContent.replace(/\\s+/g,' ').trim().toLowerCase();
      var te=h.querySelector('[title]'); var title=(te?te.getAttribute('title'):'').toLowerCase();
      if(apur<0 && (title.indexOf('apura')>=0 || t.indexOf('apura')>=0)) apur=i;
      if(total<0 && (t==='total' || title==='total')) total=i;
    });
    return { apur: apur, total: total, nCols: heads.length };
  }
  function readRows(){
    return Array.prototype.slice.call(document.querySelectorAll('datatable-body-row')).map(function(r){
      return Array.prototype.slice.call(r.querySelectorAll('datatable-body-cell')).map(function(c){
        return c.textContent.replace(/\\s+/g,' ').trim();
      });
    });
  }
  function firstRowCpf(){
    var rows=readRows(); if(!rows.length) return '';
    for(var i=0;i<rows[0].length;i++){ if(digits(rows[0][i]).length===11) return digits(rows[0][i]); }
    return '';
  }
  async function waitResults(cpf){
    var t0=Date.now(); await sleep(350);
    while(Date.now()-t0 < CFG.MAX_WAIT_MS){
      if(!isLoading()){
        if(firstRowCpf()===cpf) return 'ok';
        if(readRows().length===0){ await sleep(300); if(!isLoading() && readRows().length===0 && firstRowCpf()==='') return 'vazio'; }
      }
      await sleep(CFG.POLL_MS);
    }
    return 'timeout';
  }
  async function coletarPaginas(cpf){
    var all=readRows(); var guard=0;
    while(guard++ < 60){
      var next=document.querySelector('#btn-next-page');
      if(!next || next.disabled) break;
      var before=indicesText(); next.click();
      var t0=Date.now(); await sleep(200);
      while(Date.now()-t0 < CFG.MAX_WAIT_MS){
        if(!isLoading() && indicesText()!==before && firstRowCpf()===cpf) break;
        await sleep(CFG.POLL_MS);
      }
      all=all.concat(readRows());
    }
    return all;
  }
  async function process(cpf, i, total){
    await aguardarSePausado();
    if(!panel) criarPanel(total);
    atualizarPainel(i+1, total, cpf);
    await expandirPesquisa();
    var input=getCpfInput(); var btn=getPesquisar();
    if(!input || !btn) return { status:'sem-tela' };
    if(!datasSetadas){
      await selectNgOption('Inicial', '09/2025');
      await selectNgOption('Final', '04/2026');
      datasSetadas = true;
    }
    var cpfFmt=cpf.replace(/(\\d{3})(\\d{3})(\\d{3})(\\d{2})/, '$1.$2.$3-$4');
    setNativeValue(input, cpfFmt); input.blur(); await sleep(120);
    btn.click();
    var st=await waitResults(cpf);
    if(st==='vazio') return { status:'vazio' };
    if(st==='timeout') return { status:'timeout' };
    var chk=document.querySelector('#selecionar-todos'); if(chk && !chk.checked){ chk.click(); await sleep(120); }
    var btnAdicionar=Array.prototype.slice.call(document.querySelectorAll('button')).find(function(b){ return b.textContent.trim().toLowerCase()==='adicionar à guia' || b.textContent.trim().toLowerCase()==='adicionar a guia'; });
    if(btnAdicionar){ btnAdicionar.click(); await sleep(300); }
    var map=getColMap();
    var linhas=await coletarPaginas(cpf);
    var porData={};
    linhas.forEach(function(row){
      var data=(map.apur>=0 ? row[map.apur] : '') || '(sem data)';
      var val=parseBR(map.total>=0 ? row[map.total] : '');
      porData[data]=(porData[data]||0)+val;
    });
    return { status:'ok', nlinhas: linhas.length, porData: porData, colMap: map };
  }
  return { process: process };
})();
`;

/* ---------- funcoes de driver ---------- */
async function ensureHelpers(driver) {
  const has = await driver.executeScript('return (typeof window.__fgts!=="undefined") && (typeof window.__fgts.process==="function")');
  if (!has) { await driver.executeScript(PAGE_HELPERS); log('Helpers injetados na pagina.'); }
}
async function clockSeconds(driver) {
  try {
    return await driver.executeScript(function () {
      var el = document.querySelector('.clock'); if (!el) return -1;
      var p = el.textContent.trim().split(':').map(function (n) { return parseInt(n, 10); });
      if (!p.length || p.some(isNaN)) return -1;
      var s = 0; for (var i = 0; i < p.length; i++) s = s * 60 + p[i]; return s;
    });
  } catch (e) { return -1; }
}
async function esperarCampo(driver, maxMs) {
  const t0 = Date.now();
  while (Date.now() - t0 < (maxMs || 120000)) {
    try {
      const ok = await driver.executeScript('return !!document.querySelector(\'input[name="cpfTrabalhador"]\') && !!Array.prototype.slice.call(document.querySelectorAll("button")).find(function(b){return b.textContent.trim().toLowerCase()==="pesquisar";})');
      if (ok) return true;
    } catch (e) {}
    await sleep(500);
  }
  return false;
}
async function esperarInicio(driver) {
  log('Aguardando login + clique em "INICIAR AUTOMACAO" (botao verde na pagina)...');
  let aviso = false;
  let jaInjetou = false;
  while (true) {
    let st = null;
    try {
      st = await driver.executeScript(function () {
        var injetou = false;
        if (!document.getElementById('fgts-iniciar-btn') && (document.body || document.documentElement)) {
          var b = document.createElement('button');
          b.id = 'fgts-iniciar-btn';
          b.textContent = 'INICIAR AUTOMACAO';
          b.style.cssText = 'position:fixed;top:10px;right:10px;z-index:2147483647;padding:12px 16px;background:#10b981;color:#04231a;font:bold 14px Arial;border:0;border-radius:8px;cursor:pointer;box-shadow:0 4px 16px rgba(0,0,0,.4)';
          b.onclick = function () { window.__FGTS_START = true; b.textContent = 'INICIANDO...'; b.disabled = true; };
          (document.body || document.documentElement).appendChild(b);
          injetou = true;
        }
        var url = window.location.href;
        return { injetou: injetou, start: window.__FGTS_START === true, campo: !!document.querySelector('input[name="cpfTrabalhador"]'), url: url };
      });
    } catch (e) { log('Aguardando pagina carregar... (' + e.message + ')', 'AVISO'); }
    if (st) {
      if (st.injetou && !jaInjetou) { jaInjetou = true; log('Botao INICIAR AUTOMACAO injetado na pagina: ' + st.url); }
      if (st.start && st.campo) return true;
      if (st.start && !st.campo && !aviso) {
        aviso = true;
        log('Campo CPF nao visivel. Tentando expandir pesquisa automaticamente...', 'AVISO');
        try { await driver.executeScript("var b=Array.prototype.slice.call(document.querySelectorAll('button')).find(function(b){return b.textContent.trim().toLowerCase()==='expandir pesquisa'}); if(b)b.click();"); await sleep(500); } catch (e) {}
        // verifica se funcionou
        try { var st2 = await driver.executeScript(function(){ return {campo: !!document.querySelector('input[name=\"cpfTrabalhador\"]')}; }); if (st2 && st2.campo) { aviso = false; continue; } } catch (e) {}
        log('Nao achou o campo CPF. Abra a tela de pesquisa manualmente e clique em INICIAR de novo.', 'AVISO');
        try { await driver.executeScript('window.__FGTS_START=false; var b=document.getElementById("fgts-iniciar-btn"); if(b){b.textContent="INICIAR AUTOMACAO"; b.disabled=false;}'); } catch (e) {}
      } else if (st && !st.start) { aviso = false; }
    }
    await sleep(800);
  }
}

/* ============================ MAIN ============================ */
(async () => {
  console.log('==============================================================');
  console.log('  FGTS Digital - Automacao (Node + Selenium + Firefox ESR)');
  console.log('  Log: ' + LOG_PATH);
  console.log('==============================================================\n');

  // 1) planilha
  const planilha = escolherPlanilha();
  if (!planilha) { log('Nenhuma planilha selecionada. Passe o caminho: node fgts.js "C:\\...\\Trabalhadores.xlsx"', 'ERRO'); process.exit(1); }
  log('Planilha: ' + planilha);
  let registros;
  try { registros = lerRegistros(planilha); }
  catch (e) { log('Erro ao ler a planilha: ' + e.message, 'ERRO'); process.exit(1); }
  if (!registros.length) { log('Nenhum CPF valido apos o filtro.', 'ERRO'); process.exit(1); }

  // 2) quantidade
  let limite = registros.length;
  const resp = await ask('Foram encontrados ' + registros.length + ' CPFs. Quantos processar? (ENTER = todos): ');
  if (resp && resp.trim() !== '') {
    const n = parseInt(resp.replace(/\D/g, ''), 10);
    if (!isNaN(n) && n > 0) limite = Math.min(n, registros.length);
  }
  log('Processar: ' + limite + ' de ' + registros.length);

  // 3) navegador
  const chromeBin = acharChrome();
  if (!chromeBin) log('chrome.exe nao encontrado nos caminhos padrao. Rode o instalador de dependencias.', 'AVISO');
  else log('Chrome: ' + chromeBin);
  const driverBin = acharChromedriver();
  log('chromedriver: ' + (driverBin || '(Selenium Manager resolve automaticamente)'));

  // Chrome aberto como processo NORMAL via --remote-debugging-port
  // Assim o Selenium conecta sem injetar flags de automacao e o captcha funciona
  const CHROME_PORT = 9222;
  const http = require('http');

  function checkChromeDebug() {
    return new Promise((res) => {
      const req = http.get('http://127.0.0.1:' + CHROME_PORT + '/json/version', (r) => { res(true); r.resume(); });
      req.on('error', () => res(false));
      req.setTimeout(2000, () => { req.destroy(); res(false); });
    });
  }

  if (!(await checkChromeDebug())) {
    log('Fechando Chrome existente e reiniciando com --remote-debugging-port...');
    try { require('child_process').execSync('taskkill /f /im chrome.exe', { stdio: 'ignore' }); } catch (e) {}
    await sleep(1000);
    const args = ['--remote-debugging-port=' + CHROME_PORT, '--start-maximized', '--no-first-run', CFG.START_URL];
    require('child_process').execFile(chromeBin, args, { detached: true }).unref();
    for (let i = 0; i < 30; i++) {
      await sleep(1000);
      if (await checkChromeDebug()) { log('Chrome iniciado na porta ' + CHROME_PORT + '.'); break; }
      if (i === 5) log('Aguardando Chrome iniciar...');
    }
  } else {
    log('Chrome ja esta aberto na porta ' + CHROME_PORT + '. Conectando...');
  }

  if (!(await checkChromeDebug())) {
    log('Chrome nao iniciou. Abra manualmente com: chrome.exe --remote-debugging-port=' + CHROME_PORT, 'ERRO');
    process.exit(1);
  }

  const opts = new chrome.Options();
  opts.options_['debuggerAddress'] = '127.0.0.1:' + CHROME_PORT;
  let builder = new Builder().forBrowser('chrome').setChromeOptions(opts);
  let driver;
  try {
    driver = await builder.build();
    await driver.manage().setTimeouts({ script: CFG.SCRIPT_TIMEOUT_MS, pageLoad: CFG.PAGELOAD_TIMEOUT_MS, implicit: 0 });
    log('Conectado ao Chrome. Navegando para ' + CFG.START_URL + '...');
    await driver.get(CFG.START_URL);
    log('Pagina carregada. Faca login se necessario e clique em INICIAR AUTOMACAO.');
  } catch (e) {
    log('Falha ao conectar ao Chrome: ' + e.message, 'ERRO');
    process.exit(1);
  }

  // 4) espera login + botao Iniciar
  await esperarInicio(driver);
  log('=== INICIANDO AUTOMACAO ===');

  const resultados = []; // {i,cpf,nome,competencia,qtd,soma}
  const semDebito = [];  // {i,cpf,nome,obs}

  for (let i = 0; i < limite; i++) {
    const [cpf, nome] = registros[i];
    const cpfFmt = fmtCpf(cpf);
    try {
      // cronometro -> recarrega e continua no mesmo indice
      const cs = await clockSeconds(driver);
      if (cs >= 0 && cs <= CFG.RELOAD_AT_SEC) {
        log('Cronometro em ' + cs + 's -> recarregando pagina e retomando no index ' + i, 'AVISO');
        await driver.navigate().refresh();
        await esperarCampo(driver);
      }
      await ensureHelpers(driver);

      const res = await driver.executeAsyncScript(function (cpf, i, total) {
        var done = arguments[arguments.length - 1];
        try { window.__fgts.process(cpf, i, total).then(done, function (e) { done({ status: 'erro', erro: String(e && e.message || e) }); }); }
        catch (e) { done({ status: 'erro', erro: String(e) }); }
      }, cpf, i, limite);

      if (!res || res.status === 'sem-tela') {
        semDebito.push({ i, cpf, nome, obs: 'tela de pesquisa ausente' });
        log('[' + i + '] ' + cpfFmt + ' -> tela ausente, aguardando...', 'AVISO');
        await esperarCampo(driver); await ensureHelpers(driver);
        i--; // tenta o mesmo de novo
        continue;
      }
      if (res.status === 'vazio') { semDebito.push({ i, cpf, nome, obs: 'sem debitos' }); log('[' + i + '] ' + cpfFmt + ' ' + nome + ' -> SEM DEBITOS'); await sleep(CFG.PAUSE_MS); continue; }
      if (res.status === 'timeout') { semDebito.push({ i, cpf, nome, obs: 'TIMEOUT' }); log('[' + i + '] ' + cpfFmt + ' ' + nome + ' -> TIMEOUT', 'AVISO'); await sleep(CFG.PAUSE_MS); continue; }
      if (res.status === 'erro') { semDebito.push({ i, cpf, nome, obs: 'ERRO: ' + res.erro }); log('[' + i + '] ' + cpfFmt + ' ERRO: ' + res.erro, 'ERRO'); await sleep(CFG.PAUSE_MS); continue; }

      if (i === 0 && res.colMap) log('Mapa de colunas -> apuracao=' + res.colMap.apur + ' total=' + res.colMap.total + ' (nCols=' + res.colMap.nCols + ')');
      const porData = res.porData || {};
      const datas = Object.keys(porData);
      let somaCpf = 0;
      for (const d of datas) { somaCpf += porData[d]; resultados.push({ i, cpf: cpfFmt, nome, competencia: d, qtd: res.nlinhas, soma: porData[d] }); }
      log('[' + i + '] ' + cpfFmt + ' ' + nome + ' -> ' + res.nlinhas + ' debito(s) | ' + datas.map(d => d + '=' + fmtBR(porData[d])).join(' ; ') + ' | TOTAL=' + fmtBR(somaCpf));
    } catch (e) {
      semDebito.push({ i, cpf, nome, obs: 'ERRO: ' + e.message });
      log('[' + i + '] ' + cpfFmt + ' ERRO: ' + e.message, 'ERRO');
    }
    await sleep(CFG.PAUSE_MS);
  }

  // 5) gera planilha
  try {
    const wb = XLSX.utils.book_new();
    const rows = resultados.map(r => ({ Indice: r.i, CPF: r.cpf, Nome: r.nome, Competencia: r.competencia, Qtd_Debitos: r.qtd, Soma_Total: Number(Number(r.soma).toFixed(2)) }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{ Indice: '', CPF: '', Nome: '', Competencia: '', Qtd_Debitos: '', Soma_Total: '' }]), 'Resultados');
    const rows2 = semDebito.map(s => ({ Indice: s.i, CPF: fmtCpf(s.cpf), Nome: s.nome, Obs: s.obs }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows2.length ? rows2 : [{ Indice: '', CPF: '', Nome: '', Obs: '' }]), 'Sem_Debito_Erros');
    const outPath = path.join(process.cwd(), 'fgts_resultados_' + STAMP + '.xlsx');
    XLSX.writeFile(wb, outPath);
    log('=== CONCLUIDO === ' + resultados.length + ' linhas | ' + semDebito.length + ' sem debito/erros');
    log('Planilha gerada: ' + outPath);
  } catch (e) { log('Erro ao gerar planilha: ' + e.message, 'ERRO'); }

  await ask('\nPressione ENTER para fechar o navegador...');
  try { await driver.quit(); } catch (e) {}
  try { logStream.end(); } catch (e) {}
  process.exit(0);
})().catch(e => { log('ERRO FATAL: ' + e.message + '\n' + (e.stack || ''), 'ERRO'); process.exit(1); });
