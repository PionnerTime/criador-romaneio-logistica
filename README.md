# 📦 Criador de Romaneio — Automação Logística (Python)

Aplicação desktop desenvolvida em **Python** e **Tkinter** para otimizar a triagem diária de encomendas, consulta rápida de destinatários em planilhas Excel e montagem automatizada de romaneios em operações logísticas (*Last Mile*).

## 🎯 O Problema de Negócio
Na rotina de entregas locais, a conferência manual de pacotes em planilhas extensas e a digitação repetitiva de listas de destinatários geram gargalos operacionais, lentidão na roteirização e risco de falhas manuais.

Este software foi projetado e implementado em operação real para eliminar esse atrito, reduzindo drasticamente o tempo de preparação de rotas.

## ✨ Funcionalidades Principais
- **Busca Instantânea em Planilhas:** Leitura dinâmica de arquivos Excel (`.xlsx` / `.xlsm`) utilizando `Pandas` e `OpenPyXL`, identificando automaticamente abas e colunas de contatos.
- **Comando de Voz Integrado:** Pesquisa de destinatários e inserção direta por linha falada (ex: *"adicionar linha 2"*) via `SpeechRecognition` + `PyAudioWPatch`, acionada pelo atalho `F11`.
- **Gestão Ágil de Romaneio:** Inserção por duplo clique ou entrada manual rápida (`Enter`), ordenação alfabética (`A-Z`), exclusão via `Delete` e rolagem automática.
- **Exportação Rápida (Clipboard):** Cópia formatada com um clique para envio imediato via WhatsApp ou sistemas de rota (lista completa `Nome,Telefone` ou apenas nomes).
- **Interface Adaptável (UX):** Alternância entre **Modo Escuro (Dark Mode)** e **Modo Claro**, linhas zebradas para leitura confortável, tooltips explicativos e ajuste dinâmico de fonte (`A-` / `A+`).

## 🛠️ Tecnologias Utilizadas
- **Linguagem:** Python 3
- **Interface Gráfica (GUI):** Tkinter / ttk
- **Manipulação de Dados:** Pandas, OpenPyXL
- **Reconhecimento de Voz:** SpeechRecognition, PyAudioWPatch

## 🚀 Como Executar o Projeto

1. **Clone o repositório:**
   `git clone [https://github.com/SEU-USUARIO/criador-romaneio-logistica.git](https://github.com/SEU-USUARIO/criador-romaneio-logistica.git)`
   `cd criador-romaneio-logistica`

2. **Crie e ative um ambiente virtual:**
   `python -m venv venv`
   `.\venv\Scripts\Activate.ps1`

3. **Instale as dependências:**
   `pip install -r requirements.txt`

4. **Execute a aplicação:**
   `python criador_de_romaneio.py`
   *(O projeto acompanha o arquivo `dados_exemplo.xlsx` com registros fictícios para teste imediato).*