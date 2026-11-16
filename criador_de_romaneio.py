import os
import sys
import re
import pandas as pd
import openpyxl
import tkinter as tk
from tkinter import ttk, messagebox

# ==============================================================================
# CONFIGURAÇÃO DE ÁUDIO E MICROFONE
# Redireciona o SpeechRecognition para utilizar o PyAudioWPatch como driver de áudio
# ==============================================================================
try:
    import pyaudiowpatch as pyaudio
    sys.modules['pyaudio'] = pyaudio
except ImportError:
    pass

import speech_recognition as sr

# ==============================================================================
# CAMINHO DO ARQUIVO DA BASE DE DADOS (Caminho relativo para portfólio)
# ==============================================================================
DIRETORIO_BASE = os.path.dirname(os.path.abspath(__file__))
CAMINHO_EXCEL = os.path.join(DIRETORIO_BASE, "dados_exemplo.xlsx")


class Tooltip:
    """
    CLASSE TOOLTIP
    Cria um pequeno balão de texto explicativo (dica) quando o usuário
    passa o cursor do mouse por cima de um botão ou componente.
    """
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event=None):
        """Exibe a janelinha flutuante com o texto da dica"""
        if self.tooltip_window or not self.text:
            return
        x, y, _, _ = self.widget.bbox("insert") if self.widget.bbox("insert") else (0, 0, 0, 0)
        x = x + self.widget.winfo_rootx() + 25
        y = y + self.widget.winfo_rooty() + 25

        self.tooltip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")

        label = tk.Label(
            tw,
            text=self.text,
            justify="left",
            bg="#1C1C1E",
            fg="white",
            relief="solid",
            borderwidth=0,
            font=("Segoe UI", 9),
            padx=8,
            pady=4
        )
        label.pack(ipadx=1)

    def hide_tooltip(self, event=None):
        """Esconde a dica quando o mouse sai do elemento"""
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None


class CriadorRomaneioApp:
    """
    CLASSE PRINCIPAL DO APLICATIVO
    Gerencia a interface gráfica, leitura da planilha, atalhos, busca e geração do romaneio.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("Criador de Romaneio - Automação Logística")
        self.root.geometry("960x650")

        self.style = ttk.Style()
        self.style.theme_use("clam")

        # Variáveis de Estado da Aplicação
        self.clientes_romaneio = []
        self.clientes_filtrados_atuais = []
        self.tamanho_fonte_atual = 12
        self.modo_escuro = False

        # ----------------------------------------------------------------------
        # 1. CABEÇALHO SUPERIOR (Barra Azul)
        # ----------------------------------------------------------------------
        self.frame_topo = tk.Frame(self.root, bg="#007AFF", height=65)
        self.frame_topo.pack(fill="x", side="top")

        self.label_titulo = tk.Label(
            self.frame_topo,
            text="📦 Criador de Romaneio - Automação Logística",
            font=("Segoe UI", 14, "bold"),
            fg="white",
            bg="#007AFF"
        )
        self.label_titulo.pack(side="left", pady=16, padx=20)

        frame_controles_topo = tk.Frame(self.frame_topo, bg="#007AFF")
        frame_controles_topo.pack(side="right", padx=20, pady=16)

        self.btn_tema = tk.Button(
            frame_controles_topo,
            text="🌙",
            font=("Segoe UI Emoji", 10, "bold"),
            bg="#0056B3",
            fg="white",
            relief="flat",
            activebackground="#003D80",
            activeforeground="white",
            cursor="hand2",
            width=3,
            command=self.alternar_tema
        )
        self.btn_tema.pack(side="left", padx=(0, 15), ipady=1)
        Tooltip(self.btn_tema, "Alternar entre Modo Claro e Escuro")

        lbl_fonte = tk.Label(
            frame_controles_topo,
            text="Fonte:",
            font=("Segoe UI", 10, "bold"),
            fg="white",
            bg="#007AFF"
        )
        lbl_fonte.pack(side="left", padx=(0, 6))

        btn_fonte_menos = tk.Button(
            frame_controles_topo,
            text="A-",
            font=("Segoe UI", 10, "bold"),
            bg="#0056B3",
            fg="white",
            relief="flat",
            activebackground="#003D80",
            activeforeground="white",
            cursor="hand2",
            width=3,
            command=self.diminuir_fonte
        )
        btn_fonte_menos.pack(side="left", padx=2, ipady=1)
        Tooltip(btn_fonte_menos, "Diminuir o tamanho da letra das listas")

        btn_fonte_mais = tk.Button(
            frame_controles_topo,
            text="A+",
            font=("Segoe UI", 10, "bold"),
            bg="#0056B3",
            fg="white",
            relief="flat",
            activebackground="#003D80",
            activeforeground="white",
            cursor="hand2",
            width=3,
            command=self.aumentar_fonte
        )
        btn_fonte_mais.pack(side="left", padx=2, ipady=1)
        Tooltip(btn_fonte_mais, "Aumentar o tamanho da letra das listas")

        # ----------------------------------------------------------------------
        # 2. CONTAINER PRINCIPAL (Divide a tela em 2 colunas: Esquerda e Direita)
        # ----------------------------------------------------------------------
        self.container = tk.Frame(self.root)
        self.container.pack(fill="both", expand=True, padx=20, pady=15)

        self.container.columnconfigure(0, weight=1)
        self.container.columnconfigure(1, weight=1)
        self.container.rowconfigure(0, weight=1)

        # ======================================================================
        # PAINEL ESQUERDO: BUSCA E LISTA DA BASE DE CLIENTES
        # ======================================================================
        self.frame_esquerdo = tk.Frame(self.container)
        self.frame_esquerdo.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self.lbl_busca = tk.Label(
            self.frame_esquerdo,
            text="🔍 Buscar na Base de Clientes (0):",
            font=("Segoe UI", 10, "bold")
        )
        self.lbl_busca.pack(anchor="w", pady=(0, 5))

        self.frame_busca_linha = tk.Frame(self.frame_esquerdo)
        self.frame_busca_linha.pack(fill="x", pady=(0, 10))

        self.var_busca = tk.StringVar()
        self.var_busca.trace_add("write", self.filtrar_clientes)

        self.entry_busca = ttk.Entry(self.frame_busca_linha, textvariable=self.var_busca, font=("Segoe UI", 12))
        self.entry_busca.pack(side="left", fill="x", expand=True, ipady=5, padx=(0, 5))
        self.entry_busca.focus()

        self.entry_busca.bind("<Return>", lambda event: self.adicionar_texto_digitado())

        btn_mic = tk.Button(
            self.frame_busca_linha,
            text="🎙️",
            font=("Segoe UI Emoji", 11, "bold"),
            bg="#FF9500",
            fg="white",
            relief="flat",
            activebackground="#D97E00",
            activeforeground="white",
            cursor="hand2",
            width=3,
            command=self.ouvir_microfone
        )
        btn_mic.pack(side="right", ipady=3)
        Tooltip(btn_mic, "Pesquisar por voz (Atalho: F11)")

        self.frame_lista_base = tk.Frame(self.frame_esquerdo, highlightthickness=1)
        self.frame_lista_base.pack(fill="both", expand=True)

        scroll_base = ttk.Scrollbar(self.frame_lista_base)
        scroll_base.pack(side="right", fill="y")

        self.listbox_base = tk.Listbox(
            self.frame_lista_base,
            yscrollcommand=scroll_base.set,
            font=("Segoe UI", self.tamanho_fonte_atual),
            selectbackground="#007AFF",
            selectforeground="white",
            borderwidth=0,
            highlightthickness=0,
            activestyle="none"
        )
        self.listbox_base.pack(side="left", fill="both", expand=True)
        scroll_base.config(command=self.listbox_base.yview)

        self.listbox_base.bind("<Double-1>", self.adicionar_cliente_selecionado)

        self.frame_botoes_esq = tk.Frame(self.frame_esquerdo)
        self.frame_botoes_esq.pack(fill="x", pady=(8, 0))

        btn_adicionar = tk.Button(
            self.frame_botoes_esq,
            text="➔",
            font=("Segoe UI Emoji", 11, "bold"),
            bg="#34C759",
            fg="white",
            relief="flat",
            activebackground="#2FB34F",
            activeforeground="white",
            cursor="hand2",
            width=3,
            command=self.adicionar_cliente_selecionado
        )
        btn_adicionar.pack(side="left", padx=(0, 4), ipady=2)
        Tooltip(btn_adicionar, "Adicionar cliente selecionado na lista do romaneio")

        btn_adicionar_digitado = tk.Button(
            self.frame_botoes_esq,
            text="➕",
            font=("Segoe UI Emoji", 11, "bold"),
            bg="#007AFF",
            fg="white",
            relief="flat",
            activebackground="#0062CC",
            activeforeground="white",
            cursor="hand2",
            width=3,
            command=self.adicionar_texto_digitado
        )
        btn_adicionar_digitado.pack(side="left", padx=2, ipady=2)
        Tooltip(btn_adicionar_digitado, "Adicionar o nome/texto digitado na busca diretamente")

        # ======================================================================
        # PAINEL DIREITO: LISTA DO ROMANEIO CRIADO
        # ======================================================================
        self.frame_direito = tk.Frame(self.container)
        self.frame_direito.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        self.lbl_romaneio = tk.Label(
            self.frame_direito,
            text="📋 Clientes no Romaneio (0):",
            font=("Segoe UI", 10, "bold")
        )
        self.lbl_romaneio.pack(anchor="w", pady=(0, 5))

        self.frame_lista_romaneio = tk.Frame(self.frame_direito, highlightthickness=1)
        self.frame_lista_romaneio.pack(fill="both", expand=True)

        scroll_romaneio = ttk.Scrollbar(self.frame_lista_romaneio)
        scroll_romaneio.pack(side="right", fill="y")

        self.listbox_romaneio = tk.Listbox(
            self.frame_lista_romaneio,
            yscrollcommand=scroll_romaneio.set,
            font=("Segoe UI", self.tamanho_fonte_atual),
            selectbackground="#FF3B30",
            selectforeground="white",
            borderwidth=0,
            highlightthickness=0,
            activestyle="none"
        )
        self.listbox_romaneio.pack(side="left", fill="both", expand=True)
        scroll_romaneio.config(command=self.listbox_romaneio.yview)

        self.listbox_romaneio.bind("<Double-1>", self.remover_cliente_selecionado)

        self.frame_botoes_dir = tk.Frame(self.frame_direito)
        self.frame_botoes_dir.pack(fill="x", pady=(8, 0))

        btn_ordenar = tk.Button(
            self.frame_botoes_dir,
            text="🔤",
            font=("Segoe UI Emoji", 11, "bold"),
            bg="#5856D6",
            fg="white",
            relief="flat",
            activebackground="#4B49B6",
            activeforeground="white",
            cursor="hand2",
            width=3,
            command=self.ordenar_alfabeticamente
        )
        btn_ordenar.pack(side="left", padx=(0, 3), ipady=2)
        Tooltip(btn_ordenar, "Ordenar a lista em ordem alfabética (A-Z)")

        btn_copiar_tudo = tk.Button(
            self.frame_botoes_dir,
            text="📋",
            font=("Segoe UI Emoji", 11, "bold"),
            bg="#007AFF",
            fg="white",
            relief="flat",
            activebackground="#0062CC",
            activeforeground="white",
            cursor="hand2",
            width=3,
            command=self.copiar_para_clipboard
        )
        btn_copiar_tudo.pack(side="left", padx=3, ipady=2)
        Tooltip(btn_copiar_tudo, "Copiar tudo (Nome e Telefone)")

        btn_copiar_nomes = tk.Button(
            self.frame_botoes_dir,
            text="👤",
            font=("Segoe UI Emoji", 11, "bold"),
            bg="#30B0C7",
            fg="white",
            relief="flat",
            activebackground="#2894A8",
            activeforeground="white",
            cursor="hand2",
            width=3,
            command=self.copiar_apenas_nomes
        )
        btn_copiar_nomes.pack(side="left", padx=3, ipady=2)
        Tooltip(btn_copiar_nomes, "Copiar APENAS os nomes dos clientes")

        btn_excluir = tk.Button(
            self.frame_botoes_dir,
            text="❌",
            font=("Segoe UI Emoji", 11, "bold"),
            bg="#FF9500",
            fg="white",
            relief="flat",
            activebackground="#D97E00",
            activeforeground="white",
            cursor="hand2",
            width=3,
            command=self.remover_cliente_selecionado
        )
        btn_excluir.pack(side="left", padx=3, ipady=2)
        Tooltip(btn_excluir, "Excluir o cliente selecionado na lista (Atalho: Delete)")

        btn_limpar = tk.Button(
            self.frame_botoes_dir,
            text="🧹",
            font=("Segoe UI Emoji", 11, "bold"),
            bg="#FF3B30",
            fg="white",
            relief="flat",
            activebackground="#D73027",
            activeforeground="white",
            cursor="hand2",
            width=3,
            command=self.limpar_lista_toda
        )
        btn_limpar.pack(side="left", padx=(3, 0), ipady=2)
        Tooltip(btn_limpar, "Limpar toda a lista de romaneio de uma vez")

        # ----------------------------------------------------------------------
        # 3. ATALHOS DE TECLADO
        # ----------------------------------------------------------------------
        self.root.bind("<F11>", lambda event: self.ouvir_microfone())
        self.listbox_romaneio.bind("<Delete>", lambda event: self.remover_cliente_selecionado())
        self.root.bind("<Delete>", self.tratar_tecla_delete_global)

        self.todos_os_clientes = []
        self.carregar_dados()
        self.aplicar_tema()
        self.filtrar_clientes()

    # ==========================================================================
    # MÉTODOS DE CONTROLE VISUAL E MODO ESCURO
    # ==========================================================================
    def tratar_tecla_delete_global(self, event):
        if self.root.focus_get() == self.listbox_romaneio or self.listbox_romaneio.curselection():
            self.remover_cliente_selecionado()

    def alternar_tema(self):
        self.modo_escuro = not self.modo_escuro
        self.btn_tema.config(text="☀️" if self.modo_escuro else "🌙")
        self.aplicar_tema()

    def aplicar_tema(self):
        if self.modo_escuro:
            bg_geral = "#1C1C1E"
            bg_card = "#2C2C2E"
            fg_texto = "#FFFFFF"
            border_col = "#3A3A3C"
        else:
            bg_geral = "#F5F5F7"
            bg_card = "#FFFFFF"
            fg_texto = "#1C1C1E"
            border_col = "#D1D1D6"

        self.style.configure("TEntry", fieldbackground=bg_card, foreground=fg_texto, insertcolor=fg_texto)

        self.root.configure(bg=bg_geral)
        self.container.configure(bg=bg_geral)
        self.frame_esquerdo.configure(bg=bg_geral)
        self.frame_direito.configure(bg=bg_geral)
        self.frame_busca_linha.configure(bg=bg_geral)
        self.frame_botoes_esq.configure(bg=bg_geral)
        self.frame_botoes_dir.configure(bg=bg_geral)

        self.lbl_busca.configure(bg=bg_geral, fg=fg_texto)
        self.lbl_romaneio.configure(bg=bg_geral, fg=fg_texto)

        self.frame_lista_base.configure(bg=bg_card, highlightbackground=border_col)
        self.frame_lista_romaneio.configure(bg=bg_card, highlightbackground="#007AFF")

        self.listbox_base.configure(bg=bg_card, fg=fg_texto)
        self.listbox_romaneio.configure(bg=bg_card, fg=fg_texto)

        self.aplicar_cores_alternadas(self.listbox_base)
        self.aplicar_cores_alternadas(self.listbox_romaneio)

    def aplicar_cores_alternadas(self, listbox):
        if self.modo_escuro:
            cor_par = "#2C2C2E"
            cor_impar = "#1C1C1E"
        else:
            cor_par = "#FFFFFF"
            cor_impar = "#F2F2F7"

        for i in range(listbox.size()):
            cor_fundo = cor_par if i % 2 == 0 else cor_impar
            listbox.itemconfig(i, bg=cor_fundo)

    def aumentar_fonte(self):
        if self.tamanho_fonte_atual < 18:
            self.tamanho_fonte_atual += 1
            self.atualizar_fontes()

    def diminuir_fonte(self):
        if self.tamanho_fonte_atual > 10:
            self.tamanho_fonte_atual -= 1
            self.atualizar_fontes()

    def atualizar_fontes(self):
        nova_fonte = ("Segoe UI", self.tamanho_fonte_atual)
        self.listbox_base.config(font=nova_fonte)
        self.listbox_romaneio.config(font=nova_fonte)
        self.atualizar_painel_direito()

    # ==========================================================================
    # RECONHECIMENTO DE VOZ (MICROFONE)
    # ==========================================================================
    def ouvir_microfone(self):
        rec = sr.Recognizer()
        indices_microfone = [1, 2, None]
        total_atuais = len(self.clientes_filtrados_atuais)

        for idx in indices_microfone:
            try:
                mic = sr.Microphone(device_index=idx) if idx is not None else sr.Microphone()
                with mic as source:
                    self.lbl_busca.config(text=f"🎙️ Ouvindo voz... Fale o comando! ({total_atuais})")
                    self.root.update()

                    rec.adjust_for_ambient_noise(source, duration=0.6)
                    audio = rec.listen(source, timeout=4, phrase_time_limit=5)

                    self.lbl_busca.config(text=f"⏳ Processando fala... ({total_atuais})")
                    self.root.update()

                    texto_falado = rec.recognize_google(audio, language="pt-BR").lower()

                    self.lbl_busca.config(text=f"🔍 Buscar na Base de Clientes ({total_atuais}):")

                    palavras_acao = ["acrescentar", "adicionar", "colocar", "inserir", "linha"]
                    tem_acao = any(p in texto_falado for p in palavras_acao)
                    match_numero = re.search(r'\b(\d+)\b', texto_falado)

                    if tem_acao and match_numero:
                        num_linha = int(match_numero.group(1)) - 1
                        if 0 <= num_linha < len(self.clientes_filtrados_atuais):
                            cliente_escolhido = self.clientes_filtrados_atuais[num_linha]
                            self.clientes_romaneio.append(cliente_escolhido)
                            self.atualizar_painel_direito()

                            self.var_busca.set("")
                            self.entry_busca.focus()
                            return
                        else:
                            messagebox.showwarning("Linha Inválida", f"A linha {num_linha + 1} não existe na busca atual.", parent=self.root)
                            return

                    self.var_busca.set(texto_falado)
                    self.entry_busca.focus()
                    return

            except sr.WaitTimeoutError:
                self.lbl_busca.config(text=f"🔍 Buscar na Base de Clientes ({total_atuais}):")
                messagebox.showwarning("Microfone", "Tempo esgotado! Nenhum som foi detectado.", parent=self.root)
                return
            except sr.UnknownValueError:
                self.lbl_busca.config(text=f"🔍 Buscar na Base de Clientes ({total_atuais}):")
                messagebox.showwarning("Microfone", "Não foi possível entender a fala. Tente novamente.", parent=self.root)
                return
            except Exception:
                continue

        self.lbl_busca.config(text=f"🔍 Buscar na Base de Clientes ({total_atuais}):")
        messagebox.showerror("Erro no Microfone", "Não foi possível conectar ao microfone do computador.", parent=self.root)

    # ==========================================================================
    # CARREGAMENTO DE DADOS DA PLANILHA EXCEL
    # ==========================================================================
    def tratar_telefone(self, valor):
        if pd.isna(valor) or str(valor).strip() == "":
            return ""
        texto = str(valor).strip()
        if texto.endswith(".0"):
            texto = texto[:-2]
        return texto

    def carregar_dados(self):
        if os.path.exists(CAMINHO_EXCEL):
            try:
                excel_file = pd.ExcelFile(CAMINHO_EXCEL)
                abas_disponiveis = excel_file.sheet_names

                aba_alvo = next((aba for aba in abas_disponiveis if "contato" in aba.lower()), None)
                if not aba_alvo:
                    aba_alvo = abas_disponiveis[0]

                df = pd.read_excel(CAMINHO_EXCEL, sheet_name=aba_alvo)
                df.columns = [str(col).strip().lower() for col in df.columns]

                col_cliente = next((c for c in df.columns if "destinat" in c or "nome" in c or "client" in c), df.columns[0])
                col_celular = next((c for c in df.columns if "cel" in c or "tel" in c or "fone" in c), None)

                df[col_cliente] = df[col_cliente].fillna("").astype(str)

                self.todos_os_clientes = []
                for _, row in df.iterrows():
                    nome = row[col_cliente].strip()
                    if not nome or nome.lower() in ["none", "nan"]:
                        continue
                    
                    celular_raw = row[col_celular] if col_celular else ""
                    celular = self.tratar_telefone(celular_raw)

                    self.todos_os_clientes.append({
                        "cliente": nome,
                        "celular": celular
                    })

            except Exception as e:
                messagebox.showerror("Erro de Leitura", f"Erro ao ler a planilha:\n{e}", parent=self.root)
        else:
            messagebox.showerror("Erro", f"Planilha não encontrada no caminho:\n{CAMINHO_EXCEL}", parent=self.root)

    # ==========================================================================
    # FILTRAGEM E MANIPULAÇÃO DE LISTAS
    # ==========================================================================
    def filtrar_clientes(self, *args):
        termo = self.var_busca.get().lower().strip()
        self.listbox_base.delete(0, tk.END)
        self.clientes_filtrados_atuais = []

        contador_linha = 1
        for item in self.todos_os_clientes:
            nome = item.get("cliente", "")
            celular = item.get("celular", "")

            if termo in nome.lower() or termo in celular.lower():
                self.clientes_filtrados_atuais.append(item)
                texto_exibicao = f"  {contador_linha} - {nome},{celular}" if celular else f"  {contador_linha} - {nome}"
                self.listbox_base.insert(tk.END, texto_exibicao)
                contador_linha += 1

        self.aplicar_cores_alternadas(self.listbox_base)

        total_base = len(self.clientes_filtrados_atuais)
        self.lbl_busca.config(text=f"🔍 Buscar na Base de Clientes ({total_base}):")

    def adicionar_cliente_selecionado(self, event=None):
        selecao = self.listbox_base.curselection()
        if selecao:
            indice = selecao[0]
            if indice < len(self.clientes_filtrados_atuais):
                cliente_escolhido = self.clientes_filtrados_atuais[indice]
                self.clientes_romaneio.append(cliente_escolhido)

            self.atualizar_painel_direito()
            self.var_busca.set("")
            self.entry_busca.focus()

    def adicionar_texto_digitado(self):
        texto_digitado = self.var_busca.get().strip()
        
        if not texto_digitado:
            messagebox.showwarning("Aviso", "O campo de busca está vazio!", parent=self.root)
            return

        if "," in texto_digitado:
            partes = texto_digitado.split(",", 1)
            nome = partes[0].strip()
            celular = partes[1].strip()
        else:
            nome = texto_digitado
            celular = ""

        self.clientes_romaneio.append({
            "cliente": nome,
            "celular": celular
        })

        self.atualizar_painel_direito()
        self.var_busca.set("")
        self.entry_busca.focus()

    def remover_cliente_selecionado(self, event=None):
        selecao = self.listbox_romaneio.curselection()
        if selecao:
            indice = selecao[0]
            del self.clientes_romaneio[indice]
            self.atualizar_painel_direito()

    def limpar_lista_toda(self):
        if self.clientes_romaneio:
            if messagebox.askyesno("Confirmar Limpeza", "Deseja realmente apagar todos os clientes da lista de romaneio?", parent=self.root):
                self.clientes_romaneio.clear()
                self.atualizar_painel_direito()

    def rolar_para_o_final(self, ultimo_indice):
        self.listbox_romaneio.see(ultimo_indice)
        self.listbox_romaneio.yview_moveto(1.0)

    def atualizar_painel_direito(self):
        self.listbox_romaneio.delete(0, tk.END)

        for item in self.clientes_romaneio:
            nome = item.get("cliente", "")
            celular = item.get("celular", "")
            
            linha_formatada = f"  {nome},{celular}" if celular else f"  {nome}"
            self.listbox_romaneio.insert(tk.END, linha_formatada)

        self.aplicar_cores_alternadas(self.listbox_romaneio)

        total = len(self.clientes_romaneio)
        self.lbl_romaneio.config(text=f"📋 Clientes no Romaneio ({total}):")

        if total > 0:
            ultimo = total - 1
            self.listbox_romaneio.see(ultimo)
            self.root.update_idletasks()
            self.root.after(50, lambda: self.rolar_para_o_final(ultimo))

    def ordenar_alfabeticamente(self):
        if self.clientes_romaneio:
            self.clientes_romaneio.sort(key=lambda x: x.get("cliente", "").lower())
            self.atualizar_painel_direito()

    # ==========================================================================
    # CÓPIA PARA A ÁREA DE TRANSFERÊNCIA (CLIPBOARD)
    # ==========================================================================
    def copiar_para_clipboard(self):
        if self.clientes_romaneio:
            linhas = []
            for item in self.clientes_romaneio:
                nome = item.get("cliente", "")
                celular = item.get("celular", "")
                linhas.append(f"{nome},{celular}" if celular else nome)

            texto_copia = "\n".join(linhas)
            self.root.clipboard_clear()
            self.root.clipboard_append(texto_copia)
            self.root.update()
            messagebox.showinfo("Copiado!", "Lista completa (Nome,Telefone) copiada com sucesso!", parent=self.root)
        else:
            messagebox.showwarning("Aviso", "A lista do romaneio está vazia.", parent=self.root)

    def copiar_apenas_nomes(self):
        if self.clientes_romaneio:
            nomes = [item.get("cliente", "") for item in self.clientes_romaneio if item.get("cliente", "")]
            texto_copia = "\n".join(nomes)
            
            self.root.clipboard_clear()
            self.root.clipboard_append(texto_copia)
            self.root.update()
            messagebox.showinfo("Copiado!", "Apenas os NOMES dos clientes foram copiados!", parent=self.root)
        else:
            messagebox.showwarning("Aviso", "A lista do romaneio está vazia.", parent=self.root)


# ==============================================================================
# PONTO DE ENTRADA DO PROGRAMA
# ==============================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = CriadorRomaneioApp(root)
    root.mainloop()