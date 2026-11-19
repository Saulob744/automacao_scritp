# 🚀 Guia Prático — Automação PDTIC em 2º Plano (`main_background.py`)

Este guia mostra como preencher, atualizar e auditar **100% em segundo plano** (sem mexer no mouse e sem usar o teclado físico) os documentos PDTIC no sistema **Nex Privacy (Tecpar)**.

---

## 📋 Pré-requisitos Rápidos

1. **Planilha Excel na pasta do projeto:**
   Deixe o arquivo consolidado na raiz da pasta `automacao`:
   - `Campos Necessários para Sistema da Tecpar - Consolidado.xlsx`
   *(Caso baixe uma versão nova do Google Sheets em `Arquivo -> Baixar -> Microsoft Excel (.xlsx)`, basta substituir esse arquivo ou passar `--planilha "caminho.xlsx"`).*

2. **Guia do documento aberta no Google Chrome:**
   Abra e faça login normalmente no Nex Privacy na tela de edição do documento (ex.: `https://parana.nexprivacy.com/admin/institutional-documents/157/edit`).

3. **Terminal aberto na pasta do projeto:**
   ```powershell
   cd C:\Users\saulo.borges\Downloads\automacao
   ```

---

## ⚡ Passo a Passo Principal — Preencher Tudo do Zero (Etapas 3 a 6)

Preenche todas as **849 linhas (3.188 campos)**, faz a **2ª Passada de Auditoria (100%)** e clica automaticamente em **`Salvar rascunho`** ao final:

### Passo 1 — No Terminal (PowerShell)
```powershell
python main_background.py --tudo --timeout 600
```

### Passo 2 — Na Guia do Nex Privacy no Chrome
1. Aperte **`F12`** e clique na aba **`Console`**.
2. *(Se for a primeira vez no Chrome e aparecer aviso de colagem, digite `allow pasting` e dê Enter).*
3. Cole o comando abaixo e aperte **Enter**:
   ```js
   fetch('http://127.0.0.1:8765/bridge.js').then(r=>r.text()).then(eval)
   ```

Pronto! Você pode minimizar ou usar outras janelas normalmente (a Ponte `v5` usa *Web Worker* dedicado e não congela em 2º plano). Ao terminar (~5 minutos para um documento em branco), o robô salva o rascunho e gera o relatório em `dados/relatorio_pdtic_AAAAMMDD_HHMMSS.md`.

---

## 🔎 Como Auditar um Documento Já Salvo (100% Somente-Leitura)

Para conferir **célula por célula** se todos os **3.188 campos** salvos no navegador estão idênticos à planilha **sem alterar nada e sem clicar em Salvar**:

1. No terminal:
   ```powershell
   python main_background.py --analisar --timeout 600
   ```
2. Na guia aberta do documento (**`F12` -> `Console`**):
   ```js
   fetch('http://127.0.0.1:8765/bridge.js').then(r=>r.text()).then(eval)
   ```

> 💡 **Dica:** Se a guia não foi recarregada desde a última conexão, a ponte já está ativa e conecta sozinha sem precisar colar no `F12` novamente! Ao clicar em *Salvar rascunho* a página recarrega, então basta colar o comando `fetch(...)` uma vez quando quiser rodar outra ação.

Os arquivos da auditoria ficam salvos em:
- `dados/auditoria_somente_leitura.json`
- `dados/diagnostico_campos_dom.json`

---

## 🩹 Como Atualizar Apenas 1 Bloco ou Página (Modo Incremental)

Por padrão, o robô roda em **Modo Incremental (`somente_faltantes=True`)**: ele lê o que já está no rascunho, **preserva tudo o que já está igual à planilha** e escreve **apenas os campos novos ou alterados**!

| O que você quer atualizar | Comando no Terminal |
| :--- | :--- |
| **Somente `3.1 SWOT`** (Forças, Fraquezas, Oportunidades, Ameaças) | `python main_background.py --bloco 3.1 --timeout 600` |
| **Somente `3.2 Inventário de necessidades`** (`ID - Descrição`, Prioridade, Área, Justificativa) | `python main_background.py --bloco 3.2 --timeout 600` |
| **Somente `4.1 Metas e Ações`** | `python main_background.py --bloco 4.1 --timeout 600` |
| **Somente `4.2 Plano Orçamentário`** | `python main_background.py --bloco 4.2 --timeout 600` |
| **Somente `4.3 Gestão de Riscos`** | `python main_background.py --bloco 4.3 --timeout 600` |
| **Somente `5.1 Plano de Capacitação`** | `python main_background.py --bloco 5.1 --timeout 600` |
| **Somente `5.3 Ações Específicas`** (Macro, Micro e Extras) | `python main_background.py --bloco 5.3 --timeout 600` |
| **Somente `6.1 Indicadores (KPIs)`** | `python main_background.py --bloco 6.1 --timeout 600` |
| **Página 3 inteira** (`3.1` + `3.2`) | `python main_background.py --pagina3 --timeout 600` |
| **Página 4 inteira** (`4.1` + `4.2` + `4.3`) | `python main_background.py --pagina4 --timeout 600` |
| **Página 5 inteira** (`5.1` + `5.3`) | `python main_background.py --pagina5 --timeout 600` |
| **Página 6 inteira** (`6.1`) | `python main_background.py --pagina6 --timeout 600` |

---

## ⚙️ Opções Úteis de Linha de Comando

- **Menu interativo (sem decorar comandos):**
  ```powershell
  python main_background.py
  ```
- **Preencher na tela SEM clicar em `Salvar rascunho` ao final:**
  ```powershell
  python main_background.py --tudo --sem-salvar-rascunho --timeout 600
  ```
- **Teste rápido com poucas linhas (ex.: apenas 3 linhas de cada bloco):**
  ```powershell
  python main_background.py --tudo --limite 3 --sem-salvar-rascunho
  ```
- **Forçar reescrita de todos os campos (mesmo os já preenchidos):**
  ```powershell
  python main_background.py --tudo --reescrever-tudo --timeout 600
  ```
- **Usar outro arquivo Excel específico:**
  ```powershell
  python main_background.py --tudo --planilha "Minha_Planilha.xlsx" --timeout 600
  ```

---

## 🧠 Regras Automáticas de Formatação Aplicadas pelo Robô

Ao ler `Campos Necessários para Sistema da Tecpar - Consolidado.xlsx`, o módulo `src/leitor_dados.py` já aplica automaticamente:
1. **`3.2 Inventário de necessidades` (`Descrição`):** concatena `ID DA INICIATIVA` + ` - ` + `DESCRIÇÃO SUSCINTA...` *(ex.: `SESP.P.44 - Contratação de Assessment de Maturidade e Governança de Dados`)*.
2. **`4.1 Metas e Ações` (`Objetivo estratégico`):** remove o prefixo repetido `OE.TIC.X - ` (pois o código já vai no campo `ID`).
3. **`4.1 Metas e Ações` (`Ação principal`):** concatena `ID AÇÃO PRINCIPAL (KR)` + ` - ` + `AÇÃO PRINCIPAL`.
4. **`4.1 Metas e Ações` (`Prazo`):** une automaticamente as subcolunas `QUADRANTE` + `ANO`.
5. **`4.3 Gestão de Riscos` (`Risco identificado` e `Nível`):** concatena `ID DO RISCO` + ` - ` + `RISCO IDENTIFICADO` *(ex.: `SESP.R1 - Inclusão no PDTIC de necessidades...`)* e `NÍVEL` + ` - ` + `CLASSIFICAÇÃO` *(ex.: `8 - Baixo`)*.
6. **`6.1 Indicadores (KPIs)`:** extrai automaticamente os 108 indicadores e fórmulas (`MÉTRICA / FÓRMULA`) da aba `4.1. Metas e Ações`.
