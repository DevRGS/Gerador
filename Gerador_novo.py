import sys
import subprocess
import importlib

# Dicionário: "nome para importar" : "nome do pacote no PyPI"
dependencias = {
    "ttkbootstrap": "ttkbootstrap",
    "pptx": "python-pptx",
    "googleapiclient": "google-api-python-client",
    "google.auth.transport": "google-auth-httplib2",
    "google_auth_oauthlib": "google-auth-oauthlib",
    "requests": "requests"
}

for modulo, pacote in dependencias.items():
    try:
        importlib.import_module(modulo)
    except ImportError:
        print(f"Instalando '{pacote}' pois o módulo '{modulo}' não foi encontrado...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pacote])
        # Atualiza o cache de importação
        importlib.invalidate_caches()
        try:
            importlib.import_module(modulo)
        except ImportError:
            print(f"Falha ao importar '{modulo}' mesmo após a instalação.")


import os
import requests
import io
import pickle
import json
from datetime import date

# Tkinter e ttkbootstrap
import tkinter as tk
from tkinter import ttk
import ttkbootstrap as ttkb
from tkinter.messagebox import showerror, showinfo

# python-pptx
from pptx import Presentation

# Google Drive / Auth
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload




def baixar_arquivo_if_needed(nome_arquivo, url):
    if not os.path.exists(nome_arquivo):
        print(f"Baixando {nome_arquivo}...")
        r = requests.get(url)
        with open(nome_arquivo, "wb") as f:
            f.write(r.content)


# ---------------------------------------------------------
# Ajustes Globais
# ---------------------------------------------------------
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)
print("Novo diretório de trabalho:", os.getcwd())

CONFIG_FILE = "config_vendedor.json"
MAX_ABAS = 10


# ---------------------------------------------------------
# Configurações de vendedor (salvar/carregar)
# ---------------------------------------------------------
def carregar_config(nome_closer_var, celular_closer_var, email_closer_var):
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            nome_closer_var.set(data.get("nome_vendedor", ""))
            celular_closer_var.set(data.get("celular_vendedor", ""))
            email_closer_var.set(data.get("email_vendedor", ""))
        except (json.JSONDecodeError, FileNotFoundError):
            pass

def salvar_config(nome_closer, celular_closer, email_closer):
    dados = {
        "nome_vendedor": nome_closer,
        "celular_vendedor": celular_closer,
        "email_vendedor": email_closer
    }
    # Attempt to make it writable, handle error if not possible
    if os.path.exists(CONFIG_FILE):
        try:
            os.chmod(CONFIG_FILE, 0o666)
        except PermissionError:
            print(f"Warning: Could not change permissions for {CONFIG_FILE}")
            pass # Continue trying to write
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(dados, f, indent=4, ensure_ascii=False)
    except PermissionError:
        print(f"Error: Permission denied to write to {CONFIG_FILE}")
        pass


# ---------------------------------------------------------
# Dados de Planos e Tabelas de Preço (Ajustado conforme solicitação anterior)
# ---------------------------------------------------------
PLAN_INFO = {
    "Plano PDV": {
        "base_mensal": 99.00,
        "base_anual": 0.0,  # Valor anual não fornecido nos parâmetros
        "min_pdv": 1,
        "min_users": 2,
        "mandatory": [
            "Suporte Técnico - Via chamados",
            "Relatório Básico",
            "PDV - Frente de Caixa",
            "30 Notas Fiscais"
        ]
    },
    "Plano Gestão": {
        "base_mensal": 199.00,
        "base_anual": 0.0,  # Valor anual não fornecido nos parâmetros
        "min_pdv": 2,
        "min_users": 3,
        "mandatory": [
            "Notas Fiscais Ilimitadas",
            "Importação de XML",
            "PDV - Frente de Caixa",
            "Painel Senha TV",
            "Estoque em Grade",
            "Relatórios",
            "Suporte Técnico - Via chamados",
            "Suporte Técnico - Via chat",
            "Delivery",
            "Relatório KDS"
        ]
    },
    "Plano Performance": {
        "base_mensal": 499.00,
        "base_anual": 0.0,  # Valor anual não fornecido nos parâmetros
        "min_pdv": 3,
        "min_users": 5,
        "mandatory": [
            "Produção",
            "Promoções",
            "Notas Fiscais Ilimitadas",
            "Importação de XML",
            "Hub de Delivery",
            "Ordem de Serviço",
            "Delivery",
            "App Gestão CPlug",
            "Relatório KDS",
            "Painel Senha TV",
            "Painel Senha Mobile",
            "Controle de Mesas",
            "Estoque em Grade",
            "Marketing",
            "Relatórios",
            "Relatório Dinâmico",
            "Atualização em tempo real",
            "Facilita NFE",
            "Conciliação Bancária",
            "Contratos de cartões e outros",
            "Suporte Técnico - Via chamados",
            "Suporte Técnico - Via chat",
            "Suporte Técnico - Estendido",
            "PDV - Frente de Caixa", # Incluído como feature, quantidade base em min_pdv
            "Smart TEF" # Incluído como feature, quantidade base (3x) no plano Performance
        ]
    },
     # Mantendo planos antigos se ainda forem relevantes na interface
     # Remova-os se os novos planos os substituírem totalmente
    "Autoatendimento": {
        "base_mensal": 0.0,
        "base_anual": 419.90,
        "min_pdv": 0,
        "min_users": 1,
        "mandatory": [
            "Contratos de cartões e outros","Estoque em Grade","Notas Fiscais Ilimitadas",
            "Produção","Vendas - Estoque - Financeiro" # Assumindo Vendas... é incluído
        ]
    },
    "Bling": {
        "base_mensal": 369.80,
        "base_anual": 189.90,
        "min_pdv": 1,
        "min_users": 5,
        "mandatory": [
            "Relatórios",
            "Vendas - Estoque - Financeiro", # Assumindo Vendas... é incluído
            "Notas Fiscais Ilimitadas"
        ]
    },
    "Em Branco": {
        "base_mensal": 0.0,
        "base_anual": 0.0,
        "min_pdv": 0,
        "min_users": 0,
        "mandatory": []
    }
}

# Módulos que não recebem desconto (mantido do código original + Terminais Autoatendimento)
SEM_DESCONTO = {
    "TEF",
    "Terminais Autoatendimento", # Adicionado com base na lógica de preço da nova lista
    "Smart TEF",
    "Domínio Próprio",
    "Gestão de Entregadores",
    "Robô de WhatsApp + Recuperador de Pedido",
    "Gestão de Redes Sociais",
    "Combo de Logística",
    "Painel MultiLojas",
    "Programa de Fidelidade",
    "Integração API",
    "Integração TAP",
    "Central Telefônica (Base)",
    "Central Telefônica (Por Loja)"
    # Tipos de Notas (ex: "60 Notas Fiscais") não estão no SEM_DESCONTO com base nos dados originais
}

precos_mensais = {
    "Conciliação Bancária": 50.00,
    "Contratos de cartões e outros": 50.00,
    "Controle de Mesas": 49.00,
    "Delivery": 30.00,
    "Estoque em Grade": 30.00,
    "Importação de XML": 29.00,
    "Ordem de Serviço": 20.00,
    "Produção": 30.00,
    "Relatório Dinâmico": 50.00,
    "Notas Fiscais Ilimitadas": 119.90,
    "3000 Notas Fiscais": 0.0, # Preço 0.0, provavelmente incluído ou opção de nível gratuito
    "30 Notas Fiscais": 0.0,   # Preço 0.0 com base na inclusão fixa do PDV

    "60 Notas Fiscais": 40.00,
    "120 Notas Fiscais": 70.00,
    "250 Notas Fiscais": 90.00,

    "TEF": 99.90,
    "Smart TEF": 49.90,
    "Backup Realtime": 199.90,
    "Atualização em tempo real": 49.00,
    "Business Intelligence (BI)": 199.00,
    "Hub de Delivery": 79.00,
    "Facilita NFE": 99.00,
    "Smart Menu": 99.00,
    "Cardápio digital": 99.00, # Renomeado de Cardápio Digital, preço ajustado
    "Programa de Fidelidade": 299.90,
    "Terminais Autoatendimento": 199.00, # Nome e preço ajustados de Autoatendimento
    "Delivery Direto Básico": 247.00,
    "Delivery Direto Profissional": 200.00,
    "Delivery Direto VIP": 300.00,
    "Promoções": 24.50,
    "Marketing": 24.50,
    "Painel de Senha": 49.90, # Mantendo o original se ainda for relevante, mas novas listas têm tipos específicos
    "Integração TAP": 299.00,
    "Integração API": 299.00,
    "Relatório KDS": 29.90, # Preço 29.90, também listado como fixo.
    "App Gestão CPlug": 20.00,
    "Domínio Próprio": 19.90,
    "Gestão de Entregadores": 19.90,
    "Robô de WhatsApp + Recuperador de Pedido": 99.90,
    "Gestão de Redes Sociais": 9.90,
    "Combo de Logística": 74.90,
    "Painel MultiLojas": 199.00,
    "Central Telefônica (Base)": 399.90,
    "Central Telefônica (Por Loja)": 49.90,
    "Painel Senha Mobile": 49.00, # Novo módulo da lista
    "Suporte Técnico - Estendido": 99.00 # Novo módulo da lista
    # Módulos fixos sem preço listado e não nos precos_mensais originais são assumidos como incluídos no preço base
    # ex: "Suporte Técnico - Via chamados", "Relatório Básico", "Painel Senha TV", "Relatórios", "Suporte Técnico - Via chat", "Vendas - Estoque - Financeiro"
}


# Função utilitária para substituir placeholders no Slide (Mantida)
def substituir_placeholders_no_slide(slide, dados):
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    txt = run.text
                    # Substituir placeholders (ex: {{NOME_CLIENTE}} ou {NOME_CLIENTE})
                    import re
                    for k, v in dados.items():
                         # Substitui {{KEY}} e {KEY}
                         txt = txt.replace("{{" + k + "}}", str(v)).replace("{" + k + "}", str(v))

                    run.text = txt


# ---------------------------------------------------------
# Classe PlanoFrame (Aba)
#   => contém toda a lógica de cálculo do plano
# ---------------------------------------------------------
class PlanoFrame(ttkb.Frame):
    def __init__(
        self,
        parent,
        aba_index,
        nome_cliente_var_shared,
        validade_proposta_var_shared,
        on_close_callback=None
    ):
        super().__init__(parent)
        self.aba_index = aba_index
        self.on_close_callback = on_close_callback

        # Variáveis compartilhadas para Nome do Cliente, Validade e Plano
        self.nome_cliente_var = nome_cliente_var_shared
        self.validade_proposta_var = validade_proposta_var_shared
        self.nome_plano_var = tk.StringVar(value="") # ← Nome do plano

        # Plano atual (padrão será definido em configure_plano)
        self.current_plan = "Plano PDV"

        # Variáveis para spinboxes de incremento dedicadas (mantidas)
        self.spin_pdv_var = tk.IntVar(value=0)
        self.spin_users_var = tk.IntVar(value=0)
        self.spin_terminais_auto_var = tk.IntVar(value=0) # Renomeado para clareza
        self.spin_cardapio_var = tk.IntVar(value=0)
        self.spin_tef_var = tk.IntVar(value=0)
        self.spin_smart_tef_var = tk.IntVar(value=0)
        self.spin_app_cplug_var = tk.IntVar(value=0)
        self.spin_delivery_direto_basico_var = tk.IntVar(value=0)

        # Variáveis e widgets para os tipos de Notas Fiscais (novo layout com checkboxes)
        self.notes_vars = {
            "30 Notas Fiscais": tk.IntVar(value=0), # Usar IntVar 0 ou 1
            "60 Notas Fiscais": tk.IntVar(value=0),
            "120 Notas Fiscais": tk.IntVar(value=0),
            "250 Notas Fiscais": tk.IntVar(value=0),
            "3000 Notas Fiscais": tk.IntVar(value=0),
            "Notas Fiscais Ilimitadas": tk.IntVar(value=0)
        }
        self.notes_checkbuttons = {} # Para armazenar referências dos widgets

        # Módulos Dinâmicos (Checkbox + Spinbox)
        # Armazenar variáveis de quantidade (IntVars)
        self.module_quantities = {}
        # Armazenar referências dos widgets {'nome_modulo': {'check': cb, 'spin': sp, 'check_var': tk.BooleanVar, 'spin_var': tk.IntVar}}
        self.module_widgets = {}

        # Definir a lista de módulos que receberão o tratamento Checkbox+Spinbox
        # Esta lista é derivada de precos_mensais, excluindo spinboxes dedicadas e itens fixos sem preço
        dedicated_spinbox_module_names = [
             "Terminais Autoatendimento", # Nome em precos_mensais
             "Cardápio digital",       # Nome em precos_mensais
             "TEF",
             "Smart TEF",
             "App Gestão CPlug",
             "Delivery Direto Básico"
        ]
        notes_module_names = list(self.notes_vars.keys())

        self.quantifiable_modules = [
             m for m in precos_mensais.keys()
             if m not in dedicated_spinbox_module_names and m not in notes_module_names
        ]


        # Inicializar variáveis de quantidade para módulos dinâmicos
        for module_name in self.quantifiable_modules:
            self.module_quantities[module_name] = tk.IntVar(value=0)


        # Overrides de cálculo (mantidos)
        self.user_override_anual_active = tk.BooleanVar(value=False)
        self.user_override_discount_active = tk.BooleanVar(value=False)
        self.valor_anual_editavel = tk.StringVar(value="0,00") # Usar vírgula por padrão
        self.desconto_personalizado = tk.StringVar(value="0")


        # Armazenar valores calculados (mantidos)
        self.computed_mensal = 0.0
        self.computed_anual = 0.0
        self.computed_desconto_percent = 0.0

        # Layout com scrollbar (mantido)
        self.canvas = tk.Canvas(self)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar = ttkb.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.container = ttkb.Frame(self.canvas)
        self.canvas.create_window((0,0), window=self.container, anchor="nw")
        self.container.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        self.frame_main = ttkb.Frame(self.container)
        self.frame_main.pack(fill="both", expand=True)

        self.frame_left = ttkb.Frame(self.frame_main)
        self.frame_left.pack(side="left", fill="both", expand=True, padx=5, pady=5)

        self.frame_right = ttkb.Frame(self.frame_main)
        self.frame_right.pack(side="left", fill="y", padx=5, pady=5)

        self._montar_layout_esquerda()
        self._montar_layout_direita()

        # Configurar plano inicial (deve acontecer depois que os widgets são criados)
        self.configurar_plano("Plano PDV") # Padrão para o primeiro novo plano


    def fechar_aba(self):
        if self.on_close_callback:
            self.on_close_callback(self.aba_index)

    def _montar_layout_esquerda(self):
        # Barra superior (mantida)
        top_bar = ttkb.Frame(self.frame_left)
        top_bar.pack(fill="x", pady=5)
        ttkb.Label(top_bar, text=f"Aba Plano {self.aba_index}", font="-size 12 -weight bold").pack(side="left")
        btn_close = ttkb.Button(top_bar, text="Fechar Aba", command=self.fechar_aba)
        btn_close.pack(side="right")

        # Seleção de Plano (nomes de botão atualizados)
        frame_planos = ttkb.Labelframe(self.frame_left, text="Planos")
        frame_planos.pack(fill="x", pady=5)
        # Usar chaves do PLAN_INFO atualizado
        for p in PLAN_INFO.keys():
             ttkb.Button(frame_planos, text=p,
                         command=lambda pl=p: self.configurar_plano(pl)
                       ).pack(side="left", padx=5)

        # Seção de Notas Fiscais (Novo layout com checkboxes)
        frame_notas = ttkb.Labelframe(self.frame_left, text="Notas Fiscais")
        frame_notas.pack(fill="x", pady=5)
        f_nf_cols = ttkb.Frame(frame_notas)
        f_nf_cols.pack(fill="x", padx=5, pady=5)

        notes_modules = sorted(self.notes_vars.keys()) # Ordenar alfabeticamente

        for note_m in notes_modules:
             var = self.notes_vars[note_m]
             # Usar IntVar 0/1 e text="0/1" para refletir a seleção
             cb = ttk.Checkbutton(f_nf_cols, text=note_m, variable=var)
             cb.pack(anchor="w", pady=1)
             self.notes_checkbuttons[note_m] = cb
             # Ligar comando APÓS criar todos os botões para garantir a lógica de exclusividade
             cb.config(command=lambda m=note_m: self.handle_notes_exclusivity(m))


        # Módulos Dinâmicos (Layout Checkbox + Spinbox)
        frame_mod = ttkb.Labelframe(self.frame_left, text="Outros Módulos")
        frame_mod.pack(fill="both", expand=True, pady=5)

        # FIX: Use pack for the main grid frame, and pack for column frames
        f_mod_grid = ttkb.Frame(frame_mod)
        f_mod_grid.pack(fill="both", expand=True, padx=5, pady=5) # Use pack here

        # Create column frames and pack them side-by-side within f_mod_grid
        num_columns = 2 # Or 3, depending on space
        column_frames = []
        # Reinitialize col_widgets as it's not used after creating frames anymore
        # col_widgets = [[] for _ in range(num_columns)] # List of lists to hold widgets per column

        for col_idx in range(num_columns):
            col_frame = ttkb.Frame(f_mod_grid)
            col_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5) # Pack column frames
            column_frames.append(col_frame)

        # Create Checkbox and Spinbox for each quantifiable module
        # Arrange in columns
        for i, module_name in enumerate(sorted(self.quantifiable_modules)): # Sort for consistent order
            # Create a small frame for each module row, parented by the correct column frame
            col_idx = i % num_columns
            mod_frame = ttkb.Frame(column_frames[col_idx]) # Parented by the column frame

            # Module Name Label
            ttkb.Label(mod_frame, text=module_name).pack(side="left", padx=(0, 5))

            # Checkbox (reflects spinbox > 0, allows toggling)
            check_var = tk.BooleanVar(value=False) # New boolean var for check state
            cb = ttk.Checkbutton(mod_frame, variable=check_var)
            cb.pack(side="left")

            # Spinbox (controls quantity)
            spin_var = self.module_quantities[module_name]
            sp = ttkb.Spinbox(mod_frame, from_=0, to=999, textvariable=spin_var, width=5) # Max quantity 999? Ajustar se necessário
            sp.pack(side="left", padx=(5, 0))

            # Store widgets and link checkbox/spinbox
            self.module_widgets[module_name] = {'check': cb, 'spin': sp, 'check_var': check_var, 'spin_var': spin_var, 'frame': mod_frame}

            # Link spinbox change to update checkbox and overall values
            spin_var.trace_add("write", lambda name, index, mode, m=module_name: self.on_module_quantity_change(m))
            # Link checkbox click to update spinbox and overall values
            # Usar trace 'write' na variável booleana
            check_var.trace_add("write", lambda name, index, mode, m=module_name: self.on_module_check_change(m))

            # Pack the module frame within its column frame
            mod_frame.pack(anchor="w", pady=1) # Pack within the column frame, stacked vertically


        # Seção de Entrada de Dados (mantida)
        frame_dados = ttkb.Labelframe(self.frame_left, text="Dados do Cliente")
        frame_dados.pack(fill="x", pady=5)
        ttkb.Label(frame_dados, text="Nome do Cliente:").grid(row=0, column=0, sticky="w")
        ttkb.Entry(frame_dados, textvariable=self.nome_cliente_var).grid(row=0, column=1, padx=5, pady=2)
        ttkb.Label(frame_dados, text="Validade Proposta:").grid(row=1, column=0, sticky="w")
        ttkb.Entry(frame_dados, textvariable=self.validade_proposta_var).grid(row=1, column=1, padx=5, pady=2)

        ttkb.Label(frame_dados, text="Nome do Plano:").grid(row=2, column=0, sticky="w")
        ttkb.Entry(frame_dados, textvariable=self.nome_plano_var).grid(row=2, column=1, padx=5, pady=2)


    def _montar_layout_direita(self):
        # Seção de Incrementos (Atualizar labels se necessário, manter estrutura)
        frame_inc = ttkb.Labelframe(self.frame_right, text="Incrementos")
        frame_inc.pack(fill="x", pady=5)

        # PDVs (mantido)
        ttkb.Label(frame_inc, text="PDVs").grid(row=0, column=0, sticky="w")
        sp_pdv = ttkb.Spinbox(frame_inc, from_=0, to=99,
                              textvariable=self.spin_pdv_var,
                              command=self.atualizar_valores)
        sp_pdv.grid(row=0, column=1, padx=5, pady=2)

        # Usuários (mantido)
        ttkb.Label(frame_inc, text="Usuários").grid(row=1, column=0, sticky="w")
        sp_usr = ttkb.Spinbox(frame_inc, from_=0, to=999,
                              textvariable=self.spin_users_var,
                              command=self.atualizar_valores)
        sp_usr.grid(row=1, column=1, padx=5, pady=2)

        # Terminais Autoatendimento (Label e nome da variável atualizados)
        ttkb.Label(frame_inc, text="Terminais Autoatendimento").grid(row=2, column=0, sticky="w")
        sp_at = ttkb.Spinbox(frame_inc, from_=0, to=999,
                             textvariable=self.spin_terminais_auto_var,
                             command=self.atualizar_valores)
        sp_at.grid(row=2, column=1, padx=5, pady=2)

        # Cardápio Digital (mantido)
        ttkb.Label(frame_inc, text="Cardápio Digital").grid(row=3, column=0, sticky="w")
        sp_cd = ttkb.Spinbox(frame_inc, from_=0, to=999,
                             textvariable=self.spin_cardapio_var,
                             command=self.atualizar_valores)
        sp_cd.grid(row=3, column=1, padx=5, pady=2)

        # TEF (mantido)
        ttkb.Label(frame_inc, text="TEF").grid(row=4, column=0, sticky="w")
        sp_tef = ttkb.Spinbox(frame_inc, from_=0, to=99,
                              textvariable=self.spin_tef_var,
                              command=self.atualizar_valores)
        sp_tef.grid(row=4, column=1, padx=5, pady=2)

        # Smart TEF (mantido)
        ttkb.Label(frame_inc, text="Smart TEF").grid(row=5, column=0, sticky="w")
        sp_smf = ttkb.Spinbox(frame_inc, from_=0, to=99,
                              textvariable=self.spin_smart_tef_var,
                              command=self.atualizar_valores)
        sp_smf.grid(row=5, column=1, padx=5, pady=2)

        # App Gestão CPlug (mantido)
        ttkb.Label(frame_inc, text="App Gestão CPlug").grid(row=6, column=0, sticky="w")
        sp_app = ttkb.Spinbox(frame_inc, from_=0, to=999,
                              textvariable=self.spin_app_cplug_var,
                              command=self.atualizar_valores)
        sp_app.grid(row=6, column=1, padx=5, pady=2)

        # Delivery Direto Básico (mantido)
        ttkb.Label(frame_inc, text="Delivery Direto Básico").grid(row=7, column=0, sticky="w")
        sp_ddb = ttkb.Spinbox(frame_inc, from_=0, to=999,
                              textvariable=self.spin_delivery_direto_basico_var,
                              command=self.atualizar_valores)
        sp_ddb.grid(row=7, column=1, padx=5, pady=2)

        # Valores Finais e Overrides (mantidos)
        frame_valores = ttkb.Labelframe(self.frame_right, text="Valores Finais")
        frame_valores.pack(fill="x", pady=5)

        self.lbl_plano_mensal = ttkb.Label(frame_valores, text="Plano (Mensal): R$ 0,00", font="-size 12 -weight bold")
        self.lbl_plano_mensal.pack()
        self.lbl_plano_anual = ttkb.Label(frame_valores, text="Plano (Anual): R$ 0,00", font="-size 12 -weight bold")
        self.lbl_plano_anual.pack()
        self.lbl_treinamento = ttkb.Label(frame_valores, text="Custo Treinamento (Mensal): R$ 0,00", font="-size 12 -weight bold")
        self.lbl_treinamento.pack()
        self.lbl_desconto = ttkb.Label(frame_valores, text="Desconto: 0%", font="-size 12 -weight bold")
        self.lbl_desconto.pack()

        frame_edit_anual = ttkb.Labelframe(self.frame_right, text="Plano (Anual) (editável)")
        frame_edit_anual.pack(pady=5, fill="x")
        e_anual = ttkb.Entry(frame_edit_anual, textvariable=self.valor_anual_editavel, width=10)
        e_anual.pack(side="left", padx=5)
        e_anual.bind("<KeyRelease>", self.on_user_edit_valor_anual)
        b_reset_anual = ttkb.Button(frame_edit_anual, text="Reset Anual", command=self.on_reset_anual)
        b_reset_anual.pack(side="left", padx=5)

        frame_edit_desc = ttkb.Labelframe(self.frame_right, text="Desconto (%) (editável)")
        frame_edit_desc.pack(pady=5, fill="x")
        e_desc = ttkb.Entry(frame_edit_desc, textvariable=self.desconto_personalizado, width=10)
        e_desc.pack(side="left", padx=5)
        e_desc.bind("<KeyRelease>", self.on_user_edit_desconto)
        b_reset_desc = ttkb.Button(frame_edit_desc, text="Reset Desconto", command=self.on_reset_desconto)
        b_reset_desc.pack(side="left", padx=5)

    # Handlers para overrides (mantidos)
    def on_user_edit_valor_anual(self, *args):
        self.user_override_anual_active.set(True)
        self.user_override_discount_active.set(False)
        self.atualizar_valores()

    def on_reset_anual(self):
        self.user_override_anual_active.set(False)
        self.valor_anual_editavel.set("0,00")
        self.atualizar_valores()

    def on_user_edit_desconto(self, *args):
        self.user_override_discount_active.set(True)
        self.user_override_anual_active.set(False)
        self.desconto_personalizado.set("0")
        self.atualizar_valores()

    def on_reset_desconto(self):
        self.user_override_discount_active.set(False)
        self.desconto_personalizado.set("0")
        self.atualizar_valores()

    # Novos handlers para interação checkbox/spinbox de módulos dinâmicos
    def on_module_quantity_change(self, module_name):
        """Atualiza o estado do checkbox quando o valor do spinbox muda e dispara recálculo."""
        widgets = self.module_widgets.get(module_name)
        if not widgets:
             return
        spin_var = widgets['spin_var']
        check_var = widgets['check_var']
        qty = spin_var.get()

        # Atualizar estado do checkbox baseado na quantidade
        check_var.set(qty > 0)

        self.atualizar_valores()

    def on_module_check_change(self, module_name):
        """Atualiza o valor do spinbox quando o estado do checkbox muda e dispara recálculo."""
        widgets = self.module_widgets.get(module_name)
        if not widgets:
             return
        spin_var = widgets['spin_var']
        check_var = widgets['check_var']
        is_checked = check_var.get()

        # Obter a quantidade atual sem disparar o trace novamente
        # Remover temporariamente o trace, definir o valor, readicionar o trace
        trace_id = spin_var.trace_add("write", lambda name, index, mode, m=module_name: self.on_module_quantity_change(m))

        if is_checked and spin_var.get() == 0:
            # Se marcado e quantidade é 0, definir para 1 (ou mínimo permitido se diferente)
            spin_var.set(1)
        elif not is_checked and spin_var.get() > 0:
            # Se desmarcado e quantidade é > 0, definir para 0
            spin_var.set(0)
        # Se marcado e quantidade > 0, ou desmarcado e quantidade 0, não fazer nada no valor do spinbox

        spin_var.trace_remove("write", trace_id) # Remover o trace temporário
        self.atualizar_valores()


    # Novo handler para exclusividade de Notas Fiscais
    def handle_notes_exclusivity(self, selected_module):
        """Garante que apenas uma opção de Notas Fiscais seja selecionada por vez."""
        selected_var = self.notes_vars[selected_module]
        if selected_var.get() == 1: # Se a opção selecionada foi marcada
            for module_name, var in self.notes_vars.items():
                if module_name != selected_module and var.get() == 1:
                    var.set(0) # Desmarcar as outras

        self.atualizar_valores() # Sempre atualizar valores após uma mudança

    def configurar_plano(self, plano):
        info = PLAN_INFO.get(plano, {}) # Usar .get() para segurança
        self.current_plan = plano
        self.nome_plano_var.set(plano) # Definir a label/entrada do nome do plano

        # Resetar spinboxes dedicadas com base nos mínimos do plano
        self.spin_pdv_var.set(info.get("min_pdv", 0))
        self.spin_users_var.set(info.get("min_users", 0))
        # Resetar outras spinboxes dedicadas para 0
        self.spin_terminais_auto_var.set(0)
        self.spin_cardapio_var.set(0)
        self.spin_tef_var.set(0)
        self.spin_smart_tef_var.set(0)
        self.spin_app_cplug_var.set(0)
        self.spin_delivery_direto_basico_var.set(0)


        # Resetar checkboxes de Notas Fiscais para 0 e habilitá-los
        for var in self.notes_vars.values():
            var.set(0)
        for cb in self.notes_checkbuttons.values():
             cb.config(state='normal') # Habilitar todos inicialmente

        # Resetar quantidades de módulos dinâmicos para 0 e habilitar widgets
        for module_name in self.quantifiable_modules:
            self.module_quantities[module_name].set(0)
            widgets = self.module_widgets.get(module_name)
            if widgets:
                widgets['check'].config(state='normal')
                widgets['spin'].config(state='normal')


        # Definir módulos obrigatórios com base nas informações do novo plano
        mandatory_list = info.get("mandatory", [])

        # Lidar com Notas Fiscais obrigatórias
        for oblig in mandatory_list:
            if oblig in self.notes_vars:
                self.notes_vars[oblig].set(1)
                if oblig in self.notes_checkbuttons:
                    self.notes_checkbuttons[oblig].config(state='disabled') # Desabilitar o checkbox

        # Lidar com módulos dinâmicos obrigatórios (definir quantidade para 1 e desabilitar widgets)
        for oblig in mandatory_list:
             if oblig in self.module_quantities:
                 # Para módulos obrigatórios, definir a quantidade mínima como 1
                 self.module_quantities[oblig].set(1)
                 widgets = self.module_widgets.get(oblig)
                 if widgets:
                     # Checkbox já deve refletir a quantidade > 0 devido ao trace
                     widgets['check'].config(state='disabled')
                     widgets['spin'].config(state='disabled')

        # Não é necessário lidar com o plano Autoatendimento de forma especial aqui com base na nova estrutura

        self.user_override_anual_active.set(False)
        self.user_override_discount_active.set(False)
        self.valor_anual_editavel.set("0,00")
        self.desconto_personalizado.set("0")

        # Call atualizar_valores at the end to update calculations and UI state
        self.atualizar_valores()

    def atualizar_valores(self, *args):
        try:
            info = PLAN_INFO.get(self.current_plan, {}) # Usar .get() para segurança
        except KeyError: # Não deveria acontecer com .get(), mas por precaução
            return

        base_mensal = info.get("base_mensal", 0.0)
        mandatory = info.get("mandatory", []) # Usar get com default

        # Começar com o base do plano
        parte_descontavel = base_mensal
        parte_sem_desc = 0.0

        # Custo de PDVs e Usuários EXTRAS (assumindo que o custo da contagem base está no base_mensal)
        # PDVs
        min_pdv = info.get("min_pdv", 0)
        selected_pdv = self.spin_pdv_var.get()
        pdv_extras = max(0, selected_pdv - min_pdv)
        # Usar 59.90 como preço extra por PDV com base no preço antigo "outros" e contexto da nova lista
        parte_descontavel += pdv_extras * 59.90

        # Usuários
        min_users = info.get("min_users", 0)
        selected_users = self.spin_users_var.get()
        user_extras = max(0, selected_users - min_users)
        # Usar 19.00 como preço extra por Usuário com base na nova lista
        parte_descontavel += user_extras * 19.00

        # Custo de módulos de spinbox dedicadas (quantidade total * preço)
        # Precisa usar os nomes corretos dos módulos como chaves para precos_mensais
        dedicated_spinbox_costs = {
            "Terminais Autoatendimento": {"qty": self.spin_terminais_auto_var.get(), "price": precos_mensais.get("Terminais Autoatendimento", 0.0), "sem_desconto": "Terminais Autoatendimento" in SEM_DESCONTO},
            "Cardápio digital": {"qty": self.spin_cardapio_var.get(), "price": precos_mensais.get("Cardápio digital", 0.0), "sem_desconto": "Cardápio digital" in SEM_DESCONTO},
            "TEF": {"qty": self.spin_tef_var.get(), "price": precos_mensais.get("TEF", 0.0), "sem_desconto": "TEF" in SEM_DESCONTO},
            "Smart TEF": {"qty": self.spin_smart_tef_var.get(), "price": precos_mensais.get("Smart TEF", 0.0), "sem_desconto": "Smart TEF" in SEM_DESCONTO},
            "App Gestão CPlug": {"qty": self.spin_app_cplug_var.get(), "price": precos_mensais.get("App Gestão CPlug", 0.0), "sem_desconto": "App Gestão CPlug" in SEM_DESCONTO},
            "Delivery Direto Básico": {"qty": self.spin_delivery_direto_basico_var.get(), "price": precos_mensais.get("Delivery Direto Básico", 0.0), "sem_desconto": "Delivery Direto Básico" in SEM_DESCONTO},
        }

        for module_name, data in dedicated_spinbox_costs.items():
             qty = data["qty"]
             price = data["price"]
             is_sem_desconto = data["sem_desconto"]

             # Lidar com módulos de spinbox dedicadas obrigatórios (como Smart TEF no Performance)
             # Assumir que o base_mensal cobre a *funcionalidade*. O custo se aplica às unidades totais selecionadas.
             if qty > 0:
                 if is_sem_desconto:
                     parte_sem_desc += qty * price
                 else:
                     parte_descontavel += qty * price


        # Custo dos checkboxes de Notas Fiscais (quantidade é 1 se marcado)
        for note_m, var in self.notes_vars.items():
            if var.get() == 1: # Se marcado (quantidade é 1)
                price = precos_mensais.get(note_m, 0.0)
                # Tipos de Notas geralmente não estão no SEM_DESCONTO com base no código antigo
                parte_descontavel += price


        # Custo de módulos dinâmicos (quantidade total * preço)
        for module_name in self.quantifiable_modules:
            # Obter a quantidade usando .get() da IntVar
            qty = self.module_quantities[module_name].get()
            if qty > 0:
                price = precos_mensais.get(module_name, 0.0)
                if module_name in SEM_DESCONTO:
                     parte_sem_desc += qty * price
                else:
                     parte_descontavel += qty * price


        # --- Cálculo do Valor Mensal Final ---
        valor_mensal_automatico = parte_descontavel + parte_sem_desc

        # --- Cálculo do Valor Anual Final ---
        # Aplicar desconto apenas à parte_descontavel
        if self.user_override_anual_active.get():
            try:
                # Lidar com entrada com vírgula
                final_anual = float(self.valor_anual_editavel.get().replace(',', '.'))
            except ValueError:
                # Fallback para anualizar o mensal se a entrada for inválvida
                final_anual = valor_mensal_automatico * 12
                # FIX: Formatar o float para string ANTES de substituir o ponto pela vírgula
                self.valor_anual_editavel.set(f"{final_anual:.2f}".replace('.', ','))
        elif self.user_override_discount_active.get():
            try:
                # Lidar com entrada com vírgula
                desc_custom = float(self.desconto_personalizado.get().replace(',', '.'))
            except ValueError:
                desc_custom = 0.0
            desc_dec = desc_custom / 100.0
            # Aplicar desconto à porção *descontável*
            descontavel_anual = parte_descontavel * 12 * (1 - desc_dec)
            sem_desc_anual = parte_sem_desc * 12
            final_anual = descontavel_anual + sem_desc_anual

            # Garantir que o preço anual não seja menor que a parte não descontável anualizada
            final_anual = max(final_anual, parte_sem_desc * 12)

            # FIX: Formatar o float para string ANTES de substituir o ponto pela vírgula
            self.valor_anual_editavel.set(f"{final_anual:.2f}".replace('.', ','))
        else:
            # Desconto anual padrão (10%) na parte descontável
            desc_padrao = 0.10
            descontavel_anual = parte_descontavel * 12 * (1 - desc_padrao)
            sem_desc_anual = parte_sem_desc * 12
            final_anual = descontavel_anual + sem_desc_anual

            # Garantir que o preço anual não seja menor que a parte não descontável anualizada
            final_anual = max(final_anual, parte_sem_desc * 12)

            # FIX: Formatar o float para string ANTES de substituir o ponto pela vírgula
            self.valor_anual_editavel.set(f"{final_anual:.2f}".replace('.', ','))


        # Custo treinamento (simplificado para 0 com base na falta de novas regras)
        training_cost = 0.0
        # FIX: Formatar o float para string ANTES de substituir o ponto pela vírgula
        self.lbl_treinamento.config(text=f"Custo Treinamento (Mensal): R$ {training_cost:.2f}".replace('.', ','))


        # Atualização das labels
        # FIX: Formatar o float para string ANTES de substituir o ponto pela vírgula
        self.lbl_plano_mensal.config(text=f"Plano (Mensal): R$ {valor_mensal_automatico:.2f}".replace('.', ','))
        # FIX: Formatar o float para string ANTES de substituir o ponto pela vírgula
        self.lbl_plano_anual.config(text=f"Plano (Anual): R$ {final_anual:.2f}".replace('.', ','))


        # Calcular e exibir porcentagem de desconto
        total_mensal_anualizado = valor_mensal_automatico * 12
        if total_mensal_anualizado > 0:
            # O desconto é a diferença entre o custo mensal total anualizado e o custo anual total
            desconto_val = total_mensal_anualizado - final_anual
            desconto_percent = (desconto_val / total_mensal_anualizado) * 100
        else:
            desconto_percent = 0.0

        self.lbl_desconto.config(text=f"Desconto: {max(0, round(desconto_percent))}%") # Garantir porcentagem não negativa

        self.computed_mensal = valor_mensal_automatico
        self.computed_anual = final_anual
        self.computed_desconto_percent = max(0, round(desconto_percent)) # Garantir porcentagem não negativa

        # Atualizar estado dos checkboxes com base nas quantidades dos spinboxes após o cálculo
        for module_name in self.quantifiable_modules:
            widgets = self.module_widgets.get(module_name)
            if widgets:
                widgets['check_var'].set(widgets['spin_var'].get() > 0)

        # Atualizar estado dos checkboxes de Notas Fiscais após o cálculo (já feito no handler de exclusividade, mas garantir)
        # for note_m, var in self.notes_vars.items():
        #     if var.get() == 1:
        #         pass # Já está marcado
        #     else:
        #         pass # Já está desmarcado


    def montar_lista_modulos(self):
        linhas = []

        # Add itens com base em spinboxes dedicadas se a quantidade > 0
        selected_pdv = self.spin_pdv_var.get()
        if selected_pdv > 0:
             linhas.append(f"{selected_pdv} PDVs")

        selected_users = self.spin_users_var.get()
        if selected_users > 0:
             linhas.append(f"{selected_users} Usuário{'s' if selected_users > 1 else ''}")

        auto_qty = self.spin_terminais_auto_var.get() # Nome correto da variável
        if auto_qty > 0:
             linhas.append(f"{auto_qty} Terminal{'is' if auto_qty > 1 else ''} Autoatendimento")

        card_qty = self.spin_cardapio_var.get()
        if card_qty > 0:
            # Usar o nome exato de precos_mensais
            linhas.append(f"{card_qty} Cardápio(s) Digital(is)")

        tef_qty = self.spin_tef_var.get()
        if tef_qty > 0:
             linhas.append(f"{tef_qty} TEF")

        smart_tef_qty = self.spin_smart_tef_var.get()
        if smart_tef_qty > 0:
             linhas.append(f"{smart_tef_qty} Smart TEF")

        app_cplug_qty = self.spin_app_cplug_var.get()
        if app_cplug_qty > 0:
             linhas.append(f"{app_cplug_qty} App Gestão CPlug")

        ddb_qty = self.spin_delivery_direto_basico_var.get()
        if ddb_qty > 0:
             linhas.append(f"{ddb_qty} Delivery Direto Básico")


        # Add Notas Fiscais selecionadas (apenas se quantidade > 0)
        for note_m, var in self.notes_vars.items():
            if var.get() == 1: # Quantidade é 1 se marcado
                linhas.append(note_m)


        # Add módulos dinâmicos se a quantidade > 0
        for module_name in self.quantifiable_modules:
            qty = self.module_quantities[module_name].get()
            if qty > 0:
                 # Caso especial: Delivery Direto Profissional e VIP geralmente são itens únicos
                 if module_name in ["Delivery Direto Profissional", "Delivery Direto VIP"]:
                     linhas.append(module_name) # Não mostrar quantidade
                 else:
                    linhas.append(f"{qty} {module_name}{'s' if qty > 1 else ''}")


        # Add módulos obrigatórios fixos que não têm preço/spinbox se incluídos no plano
        mandatory = PLAN_INFO.get(self.current_plan, {}).get("mandatory", [])
        fixed_no_price_modules = [
            "Suporte Técnico - Via chamados",
            "Relatório Básico",
            "Painel Senha TV", # Listado como fixo em Gestão/Performance
            "Relatórios",     # Listado como fixo em Gestão/Performance
            "Suporte Técnico - Via chat", # Listado como fixo em Gestão/Performance
            # Relatório KDS tem preço e spinbox agora - lidado acima como quantificável
            "PDV - Frente de Caixa", # Nome da funcionalidade (quantidade base lidada pelo spinbox de PDVs)
            "Usuários", # Nome da funcionalidade (quantidade base lidada pelo spinbox de Usuários)
            "Smart TEF", # Nome da funcionalidade (quantidade base lidada pelo spinbox de Smart TEF)
            "Vendas - Estoque - Financeiro" # Fixo em planos antigos, mantido no mandatory do Bling
        ]
        for fixed_mod in fixed_no_price_modules:
             if fixed_mod in mandatory and fixed_mod not in linhas: # Evitar adicionar duplicatas
                 linhas.append(fixed_mod)


        unique_mods = []
        for mod in linhas:
            if mod not in unique_mods:
                unique_mods.append(mod)

        # Ordenar a lista final alfabeticamente para consistência
        unique_mods.sort()

        return unique_mods


    def gerar_dados_proposta(self, nome_closer, cel_closer, email_closer):
        # Manter esta função em grande parte como está, garantir que use os valores calculados
        # e a montar_lista_modulos atualizada
        nome_plano = self.nome_plano_var.get().strip() or "Plano"

        valor_mensal = self.computed_mensal
        valor_anual = self.computed_anual
        desconto_percent = self.computed_desconto_percent

        # Cálculo do custo de treinamento (atualmente 0.0)
        training_cost = 0.0 # Simplificado conforme decisão em atualizar_valores

        # FIX: Formatar o float para string ANTES de substituir o ponto pela vírgula
        plano_mensal_str = f"R$ {valor_mensal:.2f}".replace(".", ",")
        # Add custo de treinamento à exibição se > 0
        if training_cost > 0.01: # Usar um pequeno limiar
             # FIX: Formatar o float para string ANTES de substituir o ponto pela vírgula
             part_mensal_formatted = f"{valor_mensal:.2f}".replace(".", ",")
             # FIX: Formatar o float para string ANTES de substituir o ponto pela vírgula
             part_training_formatted = f"{training_cost:.2f}".replace(".", ",")
             plano_mensal_str = f"R$ {part_mensal_formatted} + R$ {part_training_formatted} (Treinamento)"


        # FIX: Formatar o float para string ANTES de substituir o ponto pela vírgula
        plano_anual_str = f"R$ {valor_anual:.2f}".replace(".", ",")

        # Lógica de suporte baseada no valor anual (mantida)
        if valor_anual >= 269.90: # Este limiar pode precisar de ajuste com base nos novos níveis de preços
            tipo_suporte = "Estendido"
            horario_suporte = "09:00 às 22:00 de Segunda a Sexta-feira & Sábado e Domingo das 11:00 às 21:00"
        else:
            tipo_suporte = "Regular"
            horario_suporte = "09:00 às 17:00 de Segunda a Sexta-feira"

        lista_mods = self.montar_lista_modulos()
        montagem = "\n".join(f"•    {m}" for m in lista_mods)

        # Cálculo de economia anual (mantido)
        # Recalcular com base nos valores atuais
        custo_anual_mensalizado = valor_mensal * 12 + training_cost # Incluir custo de treinamento
        custo_anual_plano = valor_anual * 12
        economia_val = custo_anual_mensalizado - custo_anual_plano

        if economia_val > 0.1: # Usar um pequeno limiar para comparação de ponto flutuante
             # FIX: Formatar o float para string ANTES de substituir o ponto pela vírgula
             econ = f"{economia_val:.2f}".replace(".", ",")
             economia_str = f"Economia de R$ {econ} ao ano"
        else:
             economia_str = ""


        dados = {
            "montagem_do_plano": montagem,
            "plano_mensal": plano_mensal_str,
            "plano_anual": plano_anual_str,
            "desconto_total": f"{desconto_percent}%",
            "nome_do_plano": nome_plano,
            "tipo_de_suporte": tipo_de_suporte,
            "horario_de_suporte": horario_de_suporte,
            "validade_proposta": self.validade_proposta_var.get(),
            "nome_closer": nome_closer,
            "celular_closer": cel_closer,
            "email_closer": email_closer,
            "nome_cliente": self.nome_cliente_var.get(),
            "economia_anual": economia_str
        }

        return dados


# ---------------------------------------------------------
# Funções que geram .pptx (Proposta e Material) (Mantidas)
# ---------------------------------------------------------
def gerar_proposta(lista_abas, nome_closer, celular_closer, email_closer):
    ppt_file = "Proposta Comercial ConnectPlug.pptx"
    if not os.path.exists(ppt_file):
        showerror("Erro", f"Arquivo '{ppt_file}' não encontrado!")
        return None

    try:
        prs = Presentation(ppt_file)
    except Exception as e:
        showerror("Erro", f"Falha ao abrir '{ppt_file}': {e}")
        return None

    if not lista_abas:
        showerror("Erro", "Não há abas para gerar Proposta.")
        return None

    # 1) Descobrir quais slides manter (opcional) - Manter lógica
    abas_indices = sorted([aba.aba_index for aba in lista_abas])
    used_plans = {aba.current_plan for aba in lista_abas}

    keep_slides = set()
    slide_map_aba = {}

    for i, slide in enumerate(prs.slides):
        texts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        texts.append(run.text)
        full_txt = " ".join(texts)

        # Exemplo: se tiver "slide_bling", só mantém se "Bling" estiver em used_plans
        if "slide_bling" in full_txt:
            if "Bling" not in used_plans:
                continue

        found_aba = None
        if "aba_plano_" not in full_txt:
            # Slide genérico => mantém
            keep_slides.add(i)
            slide_map_aba[i] = None
        else:
            # Slide para aba específica
            for x in abas_indices:
                marker = f"aba_plano_{x}"
                if marker in full_txt:
                    found_aba = x
                    break
            if found_aba is not None:
                keep_slides.add(i)
                slide_map_aba[i] = found_aba

    # 2) Remover slides não mantidos - Manter lógica
    for idx in reversed(range(len(prs.slides))):
        if idx not in keep_slides:
            rid = prs.slides._sldIdLst[idx].rId
            prs.part.drop_rel(rid)
            del prs.slides._sldIdLst[idx]

    # 3) Re-mapear índices - Manter lógica
    sorted_kept = sorted(list(keep_slides)) # Converter para lista antes de ordenar
    new_order_map = {}
    # Ajustar a iteração sobre slides após remoção
    for new_idx, slide in enumerate(prs.slides):
        # Encontrar o índice original deste slide
        # Isso requer comparar o slide atual com os slides originais antes da remoção
        # Ou, mais simples, usar o new_idx e sorted_kept para encontrar o old_idx correspondente
        if new_idx < len(sorted_kept): # Verificar limite para evitar IndexError
             old_idx = sorted_kept[new_idx]
             if old_idx in slide_map_aba: # Garantir que o old_idx estava no mapa
                 new_order_map[new_idx] = slide_map_aba[old_idx]
             else:
                 new_order_map[new_idx] = None # Slide genérico ou não mapeado


    # 4) Substituir placeholders - Manter lógica
    dados_de_aba = {}
    for aba in lista_abas:
        d = aba.gerar_dados_proposta(nome_closer, celular_closer, email_closer)
        dados_de_aba[aba.aba_index] = d

    # Fallback: se não tiver slides específicos, use dados da primeira aba
    fallback_aba = lista_abas[0]
    d_fallback = dados_de_aba[fallback_aba.aba_index]

    for new_idx, slide in enumerate(prs.slides):
        aba_num = new_order_map.get(new_idx, None) # Usar .get para segurança
        if aba_num is None:
            substituir_placeholders_no_slide(slide, d_fallback)
        else:
            d_aba = dados_de_aba.get(aba_num, d_fallback) # Usar dados da aba específica, fallback para a primeira
            substituir_placeholders_no_slide(slide, d_aba)

    # 5) Salvar - Manter lógica
    nome_cliente_primeira = d_fallback.get("nome_cliente", "SemNome")
    hoje_str = date.today().strftime("%d-%m-%Y")
    nome_arquivo = f"Proposta ConnectPlug - {nome_cliente_primeira} - {hoje_str}.pptx"

    try:
        prs.save(nome_arquivo)
        showinfo("Sucesso", f"Proposta gerada: '{nome_arquivo}'")
        return nome_arquivo
    except Exception as e:
        showerror("Erro", f"Falha ao salvar: {e}")
        return None

def gerar_material(lista_abas, nome_closer, celular_closer, email_closer):
    mat_file = "Material Tecnico ConnectPlug.pptx"
    if not os.path.exists(mat_file):
        showerror("Erro", f"Arquivo '{mat_file}' não encontrado!")
        return None

    try:
        prs = Presentation(mat_file)
    except Exception as e:
        showerror("Erro", f"Falha ao abrir '{mat_file}': {e}")
        return None

    if not lista_abas:
        showerror("Erro", "Não há abas para gerar Material Técnico.")
        return None

    # ---------------------------------------------------
    # 1) Descobrir módulos ativos e planos usados - ATUALIZAR para usar novas variáveis
    # ---------------------------------------------------
    modulos_ativos = set()
    planos_usados = set()

    for aba in lista_abas:
        planos_usados.add(aba.current_plan)

        # Módulos de checkbox+spinbox dinâmicos (se quantidade > 0)
        for nome_mod, var_mod in aba.module_quantities.items():
            if var_mod.get() > 0:
                 modulos_ativos.add(nome_mod)

        # Módulos de checkboxes de Notas Fiscais (se quantidade for 1)
        for nome_mod, var_mod in aba.notes_vars.items():
            if var_mod.get() == 1:
                 modulos_ativos.add(nome_mod)


        # Módulos de spinboxes dedicadas (se quantidade > 0)
        # Precisa adicionar o *nome* do módulo/funcionalidade, não apenas a contagem
        if aba.spin_pdv_var.get() > 0:
            modulos_ativos.add("PDV") # Adicionar nome da funcionalidade
        if aba.spin_users_var.get() > 0:
            modulos_ativos.add("Usuários") # Adicionar nome da funcionalidade
        if aba.spin_terminais_auto_var.get() > 0: # Nome correto da variável
            modulos_ativos.add("Terminais Autoatendimento") # Adicionar nome da funcionalidade
        if aba.spin_cardapio_var.get() > 0:
            modulos_ativos.add("Cardápio Digital") # Adicionar nome da funcionalidade
        if aba.spin_tef_var.get() > 0:
            modulos_ativos.add("TEF") # Adicionar nome da funcionalidade
        if aba.spin_smart_tef_var.get() > 0:
            modulos_ativos.add("Smart TEF") # Adicionar nome da funcionalidade
        if aba.spin_app_cplug_var.get() > 0:
            modulos_ativos.add("App Gestão CPlug") # Nome exato
        if aba.spin_delivery_direto_basico_var.get() > 0:
            modulos_ativos.add("Delivery Direto Básico") # Adicionar nome da funcionalidade


        # Adicionar módulos obrigatórios que são sempre incluídos como funcionalidades, mesmo sem preço/spinbox
        mandatory_list = PLAN_INFO.get(aba.current_plan, {}).get("mandatory", [])
        fixed_no_price_modules_in_list = [
            "Suporte Técnico - Via chamados",
            "Relatório Básico",
            "Painel Senha TV",
            "Relatórios",
            "Suporte Técnico - Via chat",
            # Relatório KDS está incluído como quantificável agora, mas também pode ser fixo
            "PDV - Frente de Caixa", # Nome da funcionalidade (quantidade base lidada pelo spinbox de PDVs)
            "Usuários", # Nome da funcionalidade (quantidade base lidada pelo spinbox de Usuários)
            "Smart TEF", # Nome da funcionalidade (quantidade base lidada pelo spinbox de Smart TEF)
            "Vendas - Estoque - Financeiro" # Fixo em planos antigos, mantido no mandatory do Bling
        ]
        for fixed_mod in fixed_no_price_modules_in_list:
             if fixed_mod in mandatory_list:
                  modulos_ativos.add(fixed_mod)

        # Lidar com Relatório KDS: se está no mandatory, adicionar como ativo.
        if "Relatório KDS" in mandatory_list:
             modulos_ativos.add("Relatório KDS")


    # ---------------------------------------------------
    # 2) Mapeamento de placeholders para módulos (Manter e atualizar com novos)
    #    Adapte conforme seus slides
    # ---------------------------------------------------
    # Se um slide contiver o texto "check_tef" e você quiser mantê-lo só se
    # "TEF" estiver em modulos_ativos, defina assim:
    MAPEAMENTO_MODULOS = {
        "slide_sempre": None,
        "check_sistema_kds": "Relatório KDS",
        "check_Hub_de_Delivery": "Hub de Delivery",
        "check_integracao_api": "Integração API",
        "check_integracao_tap": "Integração TAP",
        "check_controle_de_mesas": "Controle de Mesas",
        "check_Delivery": "Delivery",
        "check_producao": "Produção",
        "check_Estoque_em_Grade": "Estoque em Grade",
        "check_Facilita_NFE": "Facilita NFE",
        "check_Importacao_de_xml": "Importação de XML",
        "check_conciliacao_bancaria": "Conciliação Bancária",
        "check_contratos_de_cartoes": "Contratos de cartões e outros",
        "check_ordem_de_servico": "Ordem de Serviço",
        "check_relatorio_dinamico": "Relatório Dinâmico",
        "check_programa_de_fidelidade": "Programa de Fidelidade",
        "check_business_intelligence": "Business Intelligence (BI)",
        "check_smartmenu": "Smart Menu",
        "check_backup_real_time": "Backup Realtime",
        "check_att_tempo_real": "Atualização em Tempo Real",
        "check_promocao": "Promoções",
        "check_marketing": "Marketing",
        "pdv_balcao": "PDV - Frente de Caixa", # Usar nome exato do mandatory list se aplicável
        "qtd_smarttef": "Smart TEF", # Nome da funcionalidade
        "qtd_tef": "TEF", # Nome da funcionalidade
        "qtd_autoatendimento": "Terminais Autoatendimento", # Nome da funcionalidade
        "qtd_cardapio_digital": "Cardápio Digital", # Nome da funcionalidade
        "qtd_app_gestao_cplug": "App Gestão CPlug", # Nome da funcionalidade
        "qtd_delivery_direto_basico": "Delivery Direto Básico", # Nome da funcionalidade
        "check_delivery_direto_vip": "Delivery Direto VIP",
        "check_delivery_direto_profissional": "Delivery Direto Profissional",
        "check_notas_fiscais": { # Verificar se *qualquer* módulo de notas fiscais está ativo
            "30 Notas Fiscais",
            "60 Notas Fiscais",
            "120 Notas Fiscais",
            "250 Notas Fiscais",
            "3000 Notas Fiscais",
            "Notas Fiscais Ilimitadas"
        },
        # Adicionar outros placeholders potenciais dos seus slides e seus nomes de módulo correspondentes
        "check_painel_senha_mobile": "Painel Senha Mobile",
        "check_suporte_estendido": "Suporte Técnico - Estendido",
        "check_painel_senha_tv": "Painel Senha TV", # Nome da funcionalidade
        "check_suporte_chat": "Suporte Técnico - Via chat", # Nome da funcionalidade
        "check_suporte_chamados": "Suporte Técnico - Via chamados", # Nome da funcionalidade
        "check_relatorios": "Relatórios", # Nome da funcionalidade
        "check_relatorio_basico": "Relatório Básico", # Nome da funcionalidade
        "check_vendas_estoque_financeiro": "Vendas - Estoque - Financeiro" # Nome da funcionalidade
    }

    # ---------------------------------------------------
    # 3) Decidir quais slides manter - Lógica atualizada para usar MAPEAMENTO_MODULOS com conjuntos
    # ---------------------------------------------------
    keep_slides = set()

    for i, slide in enumerate(prs.slides):
        texts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                       texts.append(run.text.strip())
        full_txt = " ".join(texts)

        # Flag para saber se manteremos este slide
        slide_ok = False

        # Verificar slides específicos de plano (como Bling)
        if "slide_bling" in full_txt:
            if "Bling" in planos_usados:
                slide_ok = True
        # Adicionar verificações para outros slides específicos de plano se você os tiver
        # elif "slide_plan_pdv" in full_txt:
        #     if "Plano PDV" in planos_usados:
        #         slide_ok = True
        # ... etc ...


        # Se contiver "slide_sempre" => manter sempre
        if "slide_sempre" in full_txt:
            slide_ok = True

        # Agora verificar placeholders do MAPEAMENTO_MODULOS
        for placeholder, mapped_item in MAPEAMENTO_MODULOS.items():
             if placeholder in full_txt:
                 if mapped_item is None: # caso slide_sempre
                     slide_ok = True
                 elif isinstance(mapped_item, str): # Nome de módulo único
                     if mapped_item in modulos_ativos:
                          slide_ok = True
                          #break # Se o primeiro placeholder encontrado for suficiente para manter o slide
                 elif isinstance(mapped_item, set): # Conjunto de módulos mutuamente exclusivos (como Notas)
                      if any(module_name in modulos_ativos for module_name in mapped_item):
                           slide_ok = True
                           #break # Se qualquer módulo do conjunto estiver ativo, manter o slide

        # Se após verificar todas as regras, slide_ok ainda for False, NÃO manter o slide.
        if slide_ok:
            keep_slides.add(i)


    # 4) Remover slides não mantidos - Manter lógica
    for idx in reversed(range(len(prs.slides))):
        if idx not in keep_slides:
            rid = prs.slides._sldIdLst[idx].rId
            prs.part.drop_rel(rid)
            del prs.slides._sldIdLst[idx]

    # 5) Re-mapear índices - Manter lógica
    sorted_kept = sorted(list(keep_slides)) # Converter para lista antes de ordenar
    new_order_map = {}
    # Ajustar a iteração sobre slides após remoção
    for new_idx, slide in enumerate(prs.slides):
        # Encontrar o índice original deste slide
        # Isso requer comparar o slide atual com os slides originais antes da remoção
        # Ou, mais simples, usar o new_idx e sorted_kept para encontrar o old_idx correspondente
        if new_idx < len(sorted_kept): # Verificar limite para evitar IndexError
             old_idx = sorted_kept[new_idx]
             if old_idx in slide_map_aba: # Garantir que o old_idx estava no mapa
                 new_order_map[new_idx] = slide_map_aba[old_idx]
             else:
                 new_order_map[new_idx] = None # Slide genérico ou não mapeado


    # 6) Substituir placeholders (dados globais) - Manter lógica
    fallback_aba = lista_abas[0]
    d_fallback = fallback_aba.gerar_dados_proposta(nome_closer, celular_closer, email_closer)

    for slide in prs.slides:
        substituir_placeholders_no_slide(slide, d_fallback)

    # 7) Salvar pptx final - Manter lógica
    nome_cliente_primeira = d_fallback.get("nome_cliente", "SemNome")
    hoje_str = date.today().strftime("%d-%m-%Y")
    nome_arquivo = f"Material Tecnico ConnectPlug - {nome_cliente_primeira} - {hoje_str}.pptx"

    try:
        prs.save(nome_arquivo)
        showinfo("Sucesso", f"Material Técnico gerado: '{nome_arquivo}'")
        return nome_arquivo
    except Exception as e:
        showerror("Erro", f"Falha ao salvar: {e}")
        return None


# ---------------------------------------------------------
# Google Drive / Auth (Mantido)
# ---------------------------------------------------------
SCOPES = ['https://www.googleapis.com/auth/drive']

def baixar_client_secret_remoto():
    """Baixa o client_secret.json do repositório no GitHub se ainda não estiver salvo localmente."""
    url = "https://github.com/DevRGS/Gerador/raw/refs/heads/main/config/client_secret_788265418970-ur6f189oqvsttseeg6g77fegt0su67dj.apps.googleusercontent.com.json"
    nome_local = "client_secret_temp.json"

    if not os.path.exists(nome_local):
        print("Baixando client_secret do GitHub...")
        try:
            r = requests.get(url)
            r.raise_for_status() # Lança exceção para códigos de status ruins
            with open(nome_local, "w", encoding="utf-8") as f:
                f.write(r.text)
            print(f"'{nome_local}' baixado com sucesso.")
        except requests.exceptions.RequestException as e:
             raise Exception(f"Erro ao baixar o client_secret.json: {e}")
        except IOError as e:
             raise Exception(f"Erro ao salvar o client_secret.json: {e}")

    return nome_local

def get_gdrive_service():
    """Autentica e retorna o serviço do Google Drive com base no client_secret remoto."""
    creds = None
    try:
        CLIENT_SECRET_FILE = baixar_client_secret_remoto()
    except Exception as e:
        showerror("Erro de Configuração", str(e))
        return None

    TOKEN_FILE = 'token.json'

    # Tenta carregar o token salvo anteriormente
    if os.path.exists(TOKEN_FILE):
        try:
            with open(TOKEN_FILE, 'rb') as token:
                creds = pickle.load(token)
        except (EOFError, pickle.UnpicklingError, FileNotFoundError): # Lidar com erros potenciais ao carregar o token
             creds = None


    # Se não houver token ou ele for inválido/expirado, roda o fluxo de autenticação
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"Erro ao renovar token: {e}")
                creds = None # Forçar re-autenticação se a renovação falhar
        if not creds or not creds.valid:
            try:
                flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
                creds = flow.run_local_server(port=0)
            except FileNotFoundError:
                 showerror("Erro de Autenticação", f"Arquivo client_secret não encontrado: {CLIENT_SECRET_FILE}")
                 return None
            except Exception as e:
                 showerror("Erro de Autenticação", f"Erro durante o fluxo de autenticação: {e}")
                 return None


        # Salva o token local para reutilização
        try:
            with open(TOKEN_FILE, 'wb') as token:
                pickle.dump(creds, token)
        except IOError as e:
            print(f"Warning: Could not save token file: {e}")


    # Constrói o serviço Google Drive autenticado
    try:
        service = build('drive', 'v3', credentials=creds)
        return service
    except Exception as e:
        showerror("Erro do Google Drive", f"Falha ao construir o serviço Google Drive: {e}")
        return None


def upload_pptx_and_export_to_pdf(local_pptx_path):
    """
    Faz upload do .pptx convertendo em Google Slides,
    e baixa PDF local trocando .pptx -> .pdf
    """
    if not os.path.exists(local_pptx_path):
        showerror("Erro", f"Arquivo {local_pptx_path} não foi encontrado.")
        return

    service = get_gdrive_service()
    if service is None:
        return # Sair se o serviço não puder ser obtido

    pdf_output_name = local_pptx_path.replace(".pptx", ".pdf")

    uploaded_file_id = None # Manter controle do ID do arquivo enviado para limpeza

    try:
        # 1) Upload (convertendo)
        file_metadata = {
            'name': os.path.basename(local_pptx_path),
            'mimeType': 'application/vnd.google-apps.presentation'
        }
        media = MediaFileUpload(
            local_pptx_path,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            resumable=True
        )
        uploaded_file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()
        uploaded_file_id = uploaded_file.get('id')
        print(f"Arquivo '{local_pptx_path}' enviado como Google Slides. ID: {uploaded_file_id}")

        # 2) Exportar para PDF
        request = service.files().export_media(fileId=uploaded_file_id, mimeType='application/pdf')
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            if status:
                print(f"Progresso PDF: {int(status.progress() * 100)}%")

        with open(pdf_output_name, 'wb') as f:
            f.write(fh.getvalue())

        showinfo("Google Drive", f"PDF gerado localmente: '{pdf_output_name}'")

    except Exception as e:
        showerror("Erro ao gerar PDF", f"Falha ao processar arquivo no Google Drive: {e}")

    finally:
        # Limpar o arquivo Google Slides enviado
        if uploaded_file_id:
            try:
                service.files().delete(fileId=uploaded_file_id).execute()
                print(f"Arquivo temporário {uploaded_file_id} excluído do Google Drive.")
            except Exception as e:
                print(f"Warning: Could not delete temporary Google Drive file {uploaded_file_id}: {e}")


# ---------------------------------------------------------
# MainApp (Mantida)
# ---------------------------------------------------------
class MainApp(ttkb.Window):
    def __init__(self):
        super().__init__(themename="litera")
        self.title("Gerador de Propostas e Materiais ConnectPlug") # Título atualizado
        self.geometry("1200x800")

        self.nome_closer_var = tk.StringVar()
        self.celular_closer_var = tk.StringVar()
        self.email_closer_var = tk.StringVar()

        # Variáveis compartilhadas para TODAS as abas
        self.nome_cliente_var_shared = tk.StringVar(value="Nome do Cliente") # Placeholder padrão
        self.validade_proposta_var_shared = tk.StringVar(value="DD/MM/YYYY") # Placeholder padrão

        carregar_config(self.nome_closer_var, self.celular_closer_var, self.email_closer_var)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # Barra superior
        top_bar = ttkb.Frame(self)
        top_bar.pack(side="top", fill="x", pady=5)

        ttkb.Label(top_bar, text="Vendedor:").pack(side="left", padx=5)
        ttkb.Entry(top_bar, textvariable=self.nome_closer_var, width=15).pack(side="left", padx=5)
        ttkb.Label(top_bar, text="Cel:").pack(side="left", padx=5)
        ttkb.Entry(top_bar, textvariable=self.celular_closer_var, width=15).pack(side="left", padx=5)
        ttkb.Label(top_bar, text="Email:").pack(side="left", padx=5)
        ttkb.Entry(top_bar, textvariable=self.email_closer_var, width=20).pack(side="left", padx=5)

        self.btn_add = ttkb.Button(top_bar, text="+ Nova Aba", command=self.add_aba)
        self.btn_add.pack(side="right", padx=5)

        self.notebook = ttkb.Notebook(self)
        self.notebook.pack(fill="both", expand=True)

        bot_frame = ttkb.Frame(self)
        bot_frame.pack(side="bottom", fill="x", pady=5)

        # Botões unificados (gera .pptx e PDF)
        ttkb.Button(bot_frame, text="Gerar Proposta + PDF", command=self.on_gerar_proposta).pack(side="left", padx=5)
        ttkb.Button(bot_frame, text="Gerar Material + PDF", command=self.on_gerar_mat_tecnico).pack(side="left", padx=5)
        ttkb.Button(bot_frame, text="Gerar TUDO + PDF", command=self.on_gerar_tudo).pack(side="left", padx=5)

        self.abas_criadas = {}
        self.ultimo_indice = 0

        # Cria ao menos 1 aba no começo
        self.add_aba()

        # Baixar arquivos de modelo se necessário
        baixar_arquivo_if_needed(
            "Proposta Comercial ConnectPlug.pptx",
            "https://github.com/DevRGS/Gerador/raw/refs/heads/main/assets/Proposta%20Comercial%20ConnectPlug.pptx"
        )
        baixar_arquivo_if_needed(
            "Material Tecnico ConnectPlug.pptx",
            "https://github.com/DevRGS/Gerador/raw/refs/heads/main/assets/Material%20Tecnico%20ConnectPlug.pptx"
        )


    def on_close(self):
        salvar_config(
            self.nome_closer_var.get(),
            self.celular_closer_var.get(),
            self.email_closer_var.get()
        )
        self.destroy()


    def add_aba(self):
        if len(self.abas_criadas) >= MAX_ABAS:
            showinfo("Limite de Abas", f"O número máximo de abas ({MAX_ABAS}) foi atingido.")
            return
        self.ultimo_indice += 1
        idx = self.ultimo_indice
        frame_aba = PlanoFrame(
            self.notebook,
            idx,
            nome_cliente_var_shared=self.nome_cliente_var_shared,
            validade_proposta_var_shared=self.validade_proposta_var_shared,
            on_close_callback=self.fechar_aba
        )
        self.notebook.add(frame_aba, text=f"Aba {idx}")
        self.abas_criadas[idx] = frame_aba
        self.notebook.select(frame_aba) # Mudar para a nova aba


    def fechar_aba(self, indice):
        if indice in self.abas_criadas:
            frame_aba = self.abas_criadas[indice]
            self.notebook.forget(frame_aba)
            del self.abas_criadas[indice]

            # Se não sobrar nenhuma aba, adicionar uma nova
            if not self.abas_criadas:
                self.add_aba()


    def get_abas_ativas(self):
        # Retorna as abas ordenadas pelo índice
        return [self.abas_criadas[k] for k in sorted(self.abas_criadas.keys())]

    def on_gerar_proposta(self):
        """Gera Proposta (.pptx) e em seguida converte em PDF."""
        abas_ativas = self.get_abas_ativas()
        if not abas_ativas:
            showerror("Erro", "Nenhuma aba criada para gerar Proposta.")
            return
        pptx_file = gerar_proposta(
            abas_ativas,
            self.nome_closer_var.get(),
            self.celular_closer_var.get(),
            self.email_closer_var.get()
        )
        if pptx_file and os.path.exists(pptx_file):
            # Executar upload em um thread separado se demorar, mas por simplicidade manter aqui por enquanto
            upload_pptx_and_export_to_pdf(pptx_file)

    def on_gerar_mat_tecnico(self):
        """Gera Material Técnico (.pptx) e em seguida converte em PDF."""
        abas_ativas = self.get_abas_ativas()
        if not abas_ativas:
            showerror("Erro", "Nenhuma aba criada para gerar Material Técnico.")
            return
        pptx_file = gerar_material(
            abas_ativas,
            self.nome_closer_var.get(),
            self.celular_closer_var.get(),
            self.email_closer_var.get()
        )
        if pptx_file and os.path.exists(pptx_file):
             # Executar upload em um thread separado se demorar
             upload_pptx_and_export_to_pdf(pptx_file)

    def on_gerar_tudo(self):
        """Gera Proposta e Material Técnico, cada um em .pptx, depois converte em PDF."""
        abas_ativas = self.get_abas_ativas()
        if not abas_ativas:
            showerror("Erro", "Nenhuma aba criada para gerar.")
            return

        # 1) Gera Proposta
        pptx_prop = gerar_proposta(
            abas_ativas,
            self.nome_closer_var.get(),
            self.celular_closer_var.get(),
            self.email_closer_var.get()
        )
        if pptx_prop and os.path.exists(pptx_prop):
            upload_pptx_and_export_to_pdf(pptx_prop)

        # 2) Gera Material
        pptx_mat = gerar_material(
            abas_ativas,
            self.nome_closer_var.get(),
            self.celular_closer_var.get(),
            self.email_closer_var.get()
        )
        if pptx_mat and os.path.exists(pptx_mat):
            upload_pptx_and_export_to_pdf(pptx_mat)


def main():
    app = MainApp()
    app.mainloop()

if __name__ == "__main__":
    main()