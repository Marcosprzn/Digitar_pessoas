# Digitar Pessoas — Automação FGTS Digital

Automatiza a pesquisa de débitos por CPF no **FGTS Digital** (módulo de cobrança),
soma os valores da coluna **Total** por competência de apuração e gera uma planilha
de resultados + log detalhado.

Há **duas formas** de usar:

## 1) Node.js + Selenium + Google Chrome (recomendado — arquivo `fgts.js`)

Script standalone que abre a planilha, abre o Chrome, espera você logar e mostra um
botão **INICIAR AUTOMAÇÃO** na página.

### Fluxo
1. `RODAR.bat` (ou `node fgts.js`) → abre um diálogo para **selecionar a planilha** `.xlsx`.
2. Pergunta **quantos CPFs** processar (ENTER = todos).
3. Abre o **Chrome** no FGTS Digital e **espera você fazer login** e chegar à tela de pesquisa.
4. Clique no botão verde **INICIAR AUTOMAÇÃO** → começa.
5. Para cada CPF: digita, pesquisa, marca *selecionar-todos*, lê a tabela (todas as páginas)
   e **soma o Total por competência de apuração** (duplicados na mesma competência são somados).
6. **Cronômetro de sessão**: faltando ≤ 5s, recarrega a página e continua do mesmo ponto
   (a sessão do Chrome é preservada; o índice fica na memória do Node).
7. No fim gera `fgts_resultados_AAAA-MM-DD-HH-MM-SS.xlsx` e `fgts_log_...log`.

### Instalação
- Dê dois cliques em **`INSTALAR.bat`** (pede Administrador). Ele:
  - detecta **versão do Windows e 32/64 bits**;
  - instala **Node.js** e **Google Chrome**;
  - roda `npm install` (baixa `selenium-webdriver` e `xlsx`; o chromedriver certo é
    resolvido automaticamente pelo Selenium Manager).
- Depois use **`RODAR.bat`**.

> ⚠️ **Windows 8:** o Chrome **110+ não abre** no Win8. Use o **Chrome 109** (última versão
> compatível, 32 ou 64 bits). O `fgts.js` foi escrito para essa realidade.
> Node.js: use a linha **14.x** (compatível com Win8.1).

## 2) Userscript no navegador (alternativa — `fgts_automacao.user.js`)

Roda dentro do navegador via **Tampermonkey**, com painel próprio e botão *Selecionar
planilha e iniciar*. Usa **SheetJS** via `@require`. Bom para Win7/8 com **Firefox ESR 115**.
Veja os comentários no topo do arquivo.

## Configuração comum (`CFG` no topo dos scripts)
- `VALOR_FILTRO` — valor exigido na coluna F (padrão `'D'`; `''` desativa o filtro).
- `RELOAD_AT_SEC` — segundos restantes no cronômetro para recarregar (padrão `5`).
- `COL_CPF` / `COL_NOME` / `COL_FILTRO` — índices das colunas (0 = A).
- `START_URL` (em `fgts.js`) — endereço inicial aberto no Chrome.

## Arquivos
| Arquivo | Função |
|---------|--------|
| `fgts.js` | Automação Node + Selenium + Chrome |
| `package.json` | Dependências npm |
| `INSTALAR.bat` / `instalar_dependencias.ps1` | Instalador (Node + Chrome + npm install) |
| `RODAR.bat` | Executa a automação |
| `fgts_automacao.user.js` | Alternativa via Tampermonkey |

> **Privacidade:** planilhas com CPFs/nomes reais e os logs **não** são versionados (ver `.gitignore`).
