# =====================================================================
#  Instalador de dependencias - Automacao FGTS Digital (Node + Chrome)
#  - Detecta versao do Windows e arquitetura (32/64 bits)
#  - Instala Node.js (LTS 14 - compativel com Win8.1) e Google Chrome
#  - Roda "npm install" na pasta do projeto
#  Compativel com PowerShell 3.0+ (Windows 8)
# =====================================================================

$ErrorActionPreference = 'Continue'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$Stamp = (Get-Date).ToString('yyyy-MM-dd-HH-mm-ss')
$LogFile = Join-Path $ScriptDir ("fgts_instalador_" + $Stamp + ".log")
try { Start-Transcript -Path $LogFile -Append | Out-Null } catch {}

function Escrever($txt, $cor) { if (-not $cor) { $cor = 'Gray' }; Write-Host $txt -ForegroundColor $cor }

Escrever "==============================================================" 'Cyan'
Escrever "  Instalador de dependencias - FGTS Digital (Node + Chrome)" 'Cyan'
Escrever "  Log: $LogFile" 'DarkGray'
Escrever "==============================================================`n" 'Cyan'

# ---------------------- TLS 1.2 ----------------------
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 }
catch { try { [Net.ServicePointManager]::SecurityProtocol = 3072 } catch {} }

# ---------------------- DETECCAO DO SISTEMA ----------------------
$osVer = [Environment]::OSVersion.Version
$verStr = "$($osVer.Major).$($osVer.Minor)"
$caption = ""; try { $caption = (Get-WmiObject Win32_OperatingSystem).Caption } catch { $caption = "Windows $verStr" }
$is64os = $false
try { $is64os = [Environment]::Is64BitOperatingSystem } catch { $is64os = ($env:PROCESSOR_ARCHITECTURE -eq 'AMD64') -or (Test-Path Env:\PROCESSOR_ARCHITEW6432) }
$arch = if ($is64os) { '64 bits' } else { '32 bits' }
$nomeWindows = switch ($verStr) { '6.1' {'Windows 7'} '6.2' {'Windows 8'} '6.3' {'Windows 8.1'} '10.0' {'Windows 10 / 11'} default {"Windows (versao $verStr)"} }
$legado = ($osVer.Major -lt 10)

Escrever "----------------------- SISTEMA DETECTADO -----------------------" 'Yellow'
Escrever "  Windows...........: $nomeWindows" 'White'
Escrever "  Nome completo.....: $caption" 'White'
Escrever "  Versao (kernel)...: $verStr  (build $($osVer.Build))" 'White'
Escrever "  Arquitetura.......: $arch" 'White'
Escrever "-----------------------------------------------------------------`n" 'Yellow'

if ($legado) {
  Escrever "AVISO (Windows 8): o Chrome 110+ NAO roda neste Windows." 'Red'
  Escrever "       Use o Google Chrome 109 (ultima versao compativel com Win8)." 'Red'
  Escrever "       Este instalador baixa o Chrome padrao; se ele nao abrir," 'Red'
  Escrever "       instale manualmente o Chrome 109 (32 ou 64 bits).`n" 'Red'
}

function Baixar($url, $destino) {
  try { Escrever "  Baixando: $url" 'DarkGray'; (New-Object Net.WebClient).DownloadFile($url, $destino)
        return ((Test-Path $destino) -and ((Get-Item $destino).Length -gt 500KB)) }
  catch { Escrever "  Falha no download: $($_.Exception.Message)" 'DarkYellow'; return $false }
}
function InstalarMsi($msi) {
  try { Start-Process msiexec.exe -ArgumentList '/i', ('"' + $msi + '"'), '/qn', '/norestart' -Wait; return $true }
  catch { Escrever "  Falha ao instalar $msi : $($_.Exception.Message)" 'Red'; return $false }
}

# ---------------------- NODE.JS ----------------------
$node = $null
try { $node = (Get-Command node -ErrorAction SilentlyContinue).Source } catch {}
if (-not $node) { $p = "$env:ProgramFiles\nodejs\node.exe"; if (Test-Path $p) { $node = $p } }
if ($node) {
  $nv = ""; try { $nv = & $node -v } catch {}
  Escrever "Node.js ja instalado: $node $nv" 'Green'
} else {
  Escrever "Node.js nao encontrado. Instalando (v14.21.3 $arch)..." 'Cyan'
  $nodeUrl = if ($is64os) { 'https://nodejs.org/dist/v14.21.3/node-v14.21.3-x64.msi' } else { 'https://nodejs.org/dist/v14.21.3/node-v14.21.3-x86.msi' }
  $nodeMsi = Join-Path $env:TEMP ("node_" + $Stamp + ".msi")
  if (Baixar $nodeUrl $nodeMsi) {
    if (InstalarMsi $nodeMsi) { Escrever "Node.js instalado." 'Green' }
    Remove-Item $nodeMsi -Force -ErrorAction SilentlyContinue
  } else { Escrever "Nao consegui baixar o Node.js. Instale manualmente: https://nodejs.org/dist/v14.21.3/" 'Red' }
}
# atualiza PATH da sessao atual
$env:Path = "$env:ProgramFiles\nodejs;" + $env:Path

# ---------------------- GOOGLE CHROME ----------------------
function AcharChrome {
  foreach ($c in @("$env:ProgramFiles\Google\Chrome\Application\chrome.exe", "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe", "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe")) {
    if ($c -and (Test-Path $c)) { return $c }
  }
  return $null
}
$chrome = AcharChrome
if ($chrome) {
  Escrever "Google Chrome ja instalado: $chrome" 'Green'
} else {
  Escrever "Google Chrome nao encontrado. Baixando instalador ($arch)..." 'Cyan'
  $chUrl = if ($is64os) { 'https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi' } else { 'https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise.msi' }
  $chMsi = Join-Path $env:TEMP ("chrome_" + $Stamp + ".msi")
  if (Baixar $chUrl $chMsi) {
    if (InstalarMsi $chMsi) { Escrever "Chrome instalado." 'Green' }
    Remove-Item $chMsi -Force -ErrorAction SilentlyContinue
    $chrome = AcharChrome
  } else { Escrever "Nao consegui baixar o Chrome. Instale manualmente (no Win8, versao 109)." 'Red' }
  if ($legado -and $chrome) { Escrever "Lembre-se: no Windows 8 o Chrome precisa ser a versao 109 para abrir." 'DarkYellow' }
}

# ---------------------- NPM INSTALL ----------------------
$npm = "$env:ProgramFiles\nodejs\npm.cmd"
if (-not (Test-Path $npm)) { try { $npm = (Get-Command npm.cmd -ErrorAction SilentlyContinue).Source } catch {} }
if ($npm -and (Test-Path (Join-Path $ScriptDir 'package.json'))) {
  Escrever "`nInstalando dependencias do projeto (npm install)... pode demorar." 'Cyan'
  try {
    Start-Process -FilePath $npm -ArgumentList 'install' -WorkingDirectory $ScriptDir -Wait -NoNewWindow
    Escrever "npm install concluido." 'Green'
  } catch { Escrever "Falha no npm install: $($_.Exception.Message). Rode manualmente na pasta." 'Red' }
} else {
  Escrever "`nNao rodei npm install (npm ou package.json ausentes)." 'DarkYellow'
  Escrever "Feche e reabra o terminal e rode:  npm install" 'DarkYellow'
}

# ---------------------- FIM ----------------------
Escrever "`n----------------------- COMO USAR -----------------------" 'Yellow'
Escrever " 1) De dois cliques em RODAR.bat  (ou: node fgts.js)" 'White'
Escrever " 2) Selecione a planilha e informe quantos CPFs (ENTER = todos)." 'White'
Escrever " 3) O Chrome abre no FGTS Digital -> faca login e va para a tela de pesquisa." 'White'
Escrever " 4) Clique no botao verde 'INICIAR AUTOMACAO'." 'White'
Escrever "---------------------------------------------------------`n" 'Yellow'
Escrever "Concluido. Log salvo em: $LogFile" 'Green'
try { Stop-Transcript | Out-Null } catch {}
