# Digitar Pessoas — Automação FGTS Digital

Userscript (Tampermonkey) que automatiza a pesquisa de débitos por CPF no **FGTS Digital**
(módulo de cobrança), soma os valores da coluna **Total** por competência de apuração e gera
uma planilha de resultados.

## Recursos

- **Seleciona a planilha** (`.xlsx`) direto pelo painel na tela — nada fica embutido no script.
- Ao selecionar, **pergunta quantos CPFs** processar (ENTER = todos) e **inicia sozinho**.
- Lê **coluna A = CPF**, **coluna C = Nome**, filtra **coluna F = "D"** (configurável em `CFG`).
- Para cada CPF: digita no campo *CPF do Trabalhador*, clica em *Pesquisar*, marca *selecionar-todos*,
  lê a tabela (todas as páginas) e **soma o Total agrupado por competência de apuração**.
  Quando o mesmo CPF tem mais de um débito na mesma competência, os valores são somados.
- **Cronômetro de sessão**: quando faltam ≤ 5s, salva o progresso, **recarrega a página e retoma**
  automaticamente de onde parou (estado em `localStorage`; por isso é um userscript, não um script
  de console — ele sobrevive ao reload e não precisa re-selecionar a planilha).
- **Log detalhado** de tudo (com timestamps), baixável como `.log` a qualquer momento e no fim.
- No fim, gera a planilha `fgts_resultados_AAAA-MM-DD-HH-MM-SS.xlsx` (abas *Resultados* e
  *Sem_Debito_Erros*).

## Instalação

1. Instale a extensão **Tampermonkey** (Chrome/Edge/Firefox).
2. Tampermonkey → *Criar novo script* → cole o conteúdo de [`fgts_automacao.user.js`](fgts_automacao.user.js) → salve.
   (O `@require` do SheetJS é baixado pelo Tampermonkey na instalação.)
3. Faça login no FGTS Digital e abra a tela de **pesquisa de débitos** (com o campo *CPF do Trabalhador*).
4. Use o painel no canto superior direito: **Selecionar planilha e iniciar**.

## Botões do painel

| Botão | Ação |
|-------|------|
| Selecionar planilha e iniciar | Escolhe o `.xlsx`, pergunta a quantidade e começa |
| Retomar | Continua a partir do índice salvo (após parar/recarregar) |
| Parar | Interrompe (não apaga o progresso) |
| Baixar XLSX | Gera a planilha de resultados a qualquer momento |
| Baixar log | Baixa o `.log` com tudo que aconteceu |
| Limpar tudo | Apaga o progresso salvo no `localStorage` |

## Configuração (`CFG` no topo do script)

- `VALOR_FILTRO` — valor exigido na coluna F (padrão `'D'`; use `''` para não filtrar).
- `RELOAD_AT_SEC` — segundos restantes no cronômetro para recarregar (padrão `5`).
- `COL_CPF` / `COL_NOME` / `COL_FILTRO` — índices das colunas (0 = A).

> **Privacidade:** as planilhas com CPFs/nomes reais **não** são versionadas (ver `.gitignore`).
