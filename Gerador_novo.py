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
        except json.JSONDecodeError:
            pass

def salvar_config(nome_closer, celular_closer, email_closer):
    dados = {
        "nome_vendedor": nome_closer,
        "celular_vendedor": celular_closer,
        "email_vendedor": email_closer
    }
    if os.path.exists(CONFIG_FILE):
        os.chmod(CONFIG_FILE, 0o666)
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(dados, f, indent=4, ensure_ascii=False)
    except PermissionError:
        pass

# ---------------------------------------------------------
# Dados de Planos e Tabelas de Preço - ATUALIZADO
# ---------------------------------------------------------
# Lista de planos para a UI
LISTA_PLANOS_UI = ["PDV Básico", "Gestão", "Performance", "Autoatendimento", "Bling", "Em Branco"] # Para os botões/gatilho
LISTA_PLANOS_BLING = ["Bling - Básico", "Bling - Com Estoque em Grade"] # Mantido

# Informações detalhadas dos planos
PLAN_INFO = {
    "PDV Básico": {
        "base_mensal": 99.00,
        "base_anual": 0.0, # Cálculo anual será feito dinamicamente
        "min_pdv": 1,
        "min_users": 2,
        "max_extra_users": 1, # Limite de usuários extras
        "max_extra_pdvs": 0, # Não permite PDV extra explícito na descrição
        "mandatory": [
            "Usuários", # O contador cuida da quantidade
            "30 Notas Fiscais",
            "Suporte Técnico - Via chamados",
            "Relatório Básico",
            "PDV - Frente de Caixa" # O contador cuida da quantidade
        ],
        "allowed_optionals": [
             "Smart Menu", "Terminais Autoatendimento", "Hub de Delivery",
             "Delivery Direto Profissional", "Delivery Direto VIP", "TEF",
             "Importação de XML", "Cardápio digital"
        ]
    },
    "Gestão": {
        "base_mensal": 199.00,
        "base_anual": 0.0, # Cálculo anual será feito dinamicamente
        "min_pdv": 2,
        "min_users": 3,
        "max_extra_users": 2, # Limite de usuários extras
        "max_extra_pdvs": 1,  # Limite de PDVs extras
        "mandatory": [
            "Notas Fiscais Ilimitadas", "Importação de XML",
            "PDV - Frente de Caixa", # Contador cuida da qtd
            "Usuários", # Contador cuida da qtd
            "Painel Senha TV", "Estoque em Grade", "Relatórios",
            "Suporte Técnico - Via chamados", "Suporte Técnico - Via chat",
            "Delivery", "Relatório KDS"
        ],
         "allowed_optionals": [
             "Facilita NFE", "Conciliação Bancária", "Contratos de cartões e outros",
             "Delivery Direto Profissional", "Delivery Direto VIP", "TEF",
             "Integração API", "Business Intelligence (BI)", "Backup Realtime",
             "Cardápio digital", "Smart Menu", "Hub de Delivery", "Ordem de Serviço",
             "App Gestão CPlug", "Painel Senha Mobile", "Controle de Mesas",
             "Produção", "Promoções", "Marketing", "Relatório Dinâmico",
             "Atualização em tempo real", "Smart TEF", # Limitado a 3
             "Terminais Autoatendimento", "Suporte Técnico - Estendido"
         ]
    },
    "Performance": {
        "base_mensal": 499.00,
        "base_anual": 0.0, # Cálculo anual será feito dinamicamente
        "min_pdv": 3,
        "min_users": 5,
        "max_extra_users": 5, # Limite de usuários extras
        "max_extra_pdvs": 2,  # Limite de PDVs extras
        "mandatory": [
            "Produção", "Promoções", "Notas Fiscais Ilimitadas", "Importação de XML",
            "Hub de Delivery", "Ordem de Serviço", "Delivery", "App Gestão CPlug",
            "Relatório KDS", "Painel Senha TV", "Painel Senha Mobile", "Controle de Mesas",
            "Estoque em Grade", "Marketing", "Relatórios", "Relatório Dinâmico",
            "Atualização em tempo real", "Facilita NFE", "Conciliação Bancária",
            "Contratos de cartões e outros", "Suporte Técnico - Via chamados",
            "Suporte Técnico - Via chat", "Suporte Técnico - Estendido",
            "PDV - Frente de Caixa", # Contador cuida da qtd
            "Smart TEF", # Contador cuida da qtd (3 incluídos)
            "Usuários" # Contador cuida da qtd
        ],
        "allowed_optionals": [
            "TEF", "Programa de Fidelidade", "Integração Tap", "Integração API",
            "Business Intelligence (BI)", "Backup Realtime", "Cardápio digital",
            "Smart Menu", "Terminais Autoatendimento", "Delivery Direto Profissional",
            "Delivery Direto VIP"
        ]
    },
    "Autoatendimento": { # Mantido como no original, pode precisar de revisão se a lógica mudou
        "base_mensal": 0.0,
        "base_anual": 419.90, # Valor base anual para 1 terminal
        "min_pdv": 0,
        "min_users": 1,
        "max_extra_users": 998, # Sem limite especificado, manter alto
        "max_extra_pdvs": 99,  # Sem limite especificado, manter alto
        "mandatory": [
            "Contratos de cartões e outros","Estoque em Grade","Notas Fiscais Ilimitadas",
            "Produção","Vendas - Estoque - Financeiro" # Verificar se "Vendas - Estoque - Financeiro" ainda existe/é relevante
        ],
        "allowed_optionals": [] # Definir quais opcionais se aplicam aqui, se houver
    },
    # --- Variações do Bling (Mantido como no original, revisar se necessário) ---
    "Bling - Básico": {
        "base_mensal_original": 369.80,
        "base_anual": 189.90,
        "min_pdv": 1,
        "min_users": 5,
        "max_extra_users": 994, # Sem limite especificado, manter alto
        "max_extra_pdvs": 98,  # Sem limite especificado, manter alto
        "mandatory": [
            "Relatórios",
            "Vendas - Estoque - Financeiro", # Verificar se ainda existe/relevante
            "Notas Fiscais Ilimitadas"
        ],
        "base_mensal": 369.80, # Mantido para cálculo inicial
        "allowed_optionals": [] # Definir quais opcionais se aplicam
    },
    "Bling - Com Estoque em Grade": {
        "base_mensal_original": 399.80,
        "base_anual": 219.90,
        "min_pdv": 1,
        "min_users": 5,
         "max_extra_users": 994, # Sem limite especificado, manter alto
         "max_extra_pdvs": 98,  # Sem limite especificado, manter alto
        "mandatory": [
            "Relatórios",
            "Vendas - Estoque - Financeiro", # Verificar se ainda existe/relevante
            "Notas Fiscais Ilimitadas",
            "Estoque em Grade"
        ],
        "base_mensal": 399.80, # Mantido para cálculo inicial
        "allowed_optionals": [] # Definir quais opcionais se aplicam
    },
    # --- Fim Variações Bling ---
    "Em Branco": {
        "base_mensal": 0.0,
        "base_anual": 0.0,
        "min_pdv": 0,
        "min_users": 0,
        "max_extra_users": 999, # Sem limite
        "max_extra_pdvs": 99,  # Sem limite
        "mandatory": [],
        "allowed_optionals": [] # Definir quais opcionais se aplicam
    }
}

# Módulos que não recebem o desconto padrão de 10% no anual (se aplicável)
# Revisar esta lista com base nos novos preços e regras, se necessário
SEM_DESCONTO = {
    "TEF", # Explicitamente opcional com preço fixo
    #"Autoatendimento", # O plano Autoatendimento tem lógica própria
    "Terminais Autoatendimento", # Opcional com preço fixo
    "Smart TEF", # Opcional/Incluído com preço fixo
    "Delivery Direto Profissional", # Opcional com preço fixo
    "Delivery Direto VIP", # Opcional com preço fixo
    # Adicionar outros conforme necessário (ex: Programa de Fidelidade, Integrações?)
    "Programa de Fidelidade", "Integração Tap", "Integração API",
    "Business Intelligence (BI)", "Backup Realtime",
    # Manter os antigos se ainda forem relevantes?
    "Domínio Próprio", "Gestão de Entregadores", "Robô de WhatsApp + Recuperador de Pedido",
    "Gestão de Redes Sociais", "Combo de Logística", "Painel MultiLojas",
    "Central Telefônica (Base)", "Central Telefônica (Por Loja)"
}

# Dicionário de preços mensais - ATUALIZADO
precos_mensais = {
    # --- Módulos Fixos (Preço zero pois incluídos na base, mas listados para referência) ---
    "Usuários": 0.0, # Preço do extra é tratado separadamente
    "30 Notas Fiscais": 0.0,
    "Suporte Técnico - Via chamados": 0.0,
    "Relatório Básico": 0.0,
    "PDV - Frente de Caixa": 0.0, # Preço do extra é tratado separadamente
    "Notas Fiscais Ilimitadas": 0.0,
    "Importação de XML": 29.00, # Listado como fixo no Gestão/Performance, mas opcional com preço no PDV
    "Painel Senha TV": 0.0,
    "Estoque em Grade": 0.0, # Fixo em Gestão/Performance, era opcional antes
    "Relatórios": 0.0, # Fixo em Gestão/Performance, era mandatório antes
    "Suporte Técnico - Via chat": 0.0,
    "Delivery": 0.0, # Fixo em Gestão/Performance, era opcional antes
    "Relatório KDS": 0.0, # Fixo em Gestão/Performance, era opcional antes
    "Produção": 30.00, # Fixo em Performance, opcional no Gestão
    "Promoções": 24.50, # Fixo em Performance, opcional no Gestão
    "Hub de Delivery": 79.00, # Fixo em Performance, opcional em PDV/Gestão
    "Ordem de Serviço": 20.00, # Fixo em Performance, opcional no Gestão
    "App Gestão CPlug": 20.00, # Fixo em Performance, opcional no Gestão
    "Painel Senha Mobile": 49.00, # Fixo em Performance, opcional no Gestão
    "Controle de Mesas": 49.00, # Fixo em Performance, opcional no Gestão
    "Marketing": 24.50, # Fixo em Performance, opcional no Gestão
    "Relatório Dinâmico": 50.00, # Fixo em Performance, opcional no Gestão
    "Atualização em tempo real": 49.00, # Fixo em Performance, opcional no Gestão
    "Facilita NFE": 99.00, # Fixo em Performance, opcional no Gestão
    "Conciliação Bancária": 50.00, # Fixo em Performance, opcional no Gestão
    "Contratos de cartões e outros": 50.00, # Fixo em Performance, opcional no Gestão
    "Suporte Técnico - Estendido": 99.00, # Fixo em Performance, opcional no Gestão
    "Smart TEF": 49.90, # Fixo em Performance (3 incluídos), opcional no Gestão (limite 3)

    # --- Módulos Opcionais (com preços) ---
    "Smart Menu": 99.90, # Usando 99.90 conforme mais frequente/última menção
    "Terminais Autoatendimento": 199.00,
    "Delivery Direto Profissional": 200.00,
    "Delivery Direto VIP": 300.00,
    "TEF": 99.90, # Usando 99.90
    "Cardápio digital": 99.00,
    "Backup Realtime": 199.90, # Preço do Gestão/Performance
    "Business Intelligence (BI)": 199.00, # Preço do Gestão (Performance é diferente)
    # Preços específicos do Performance (Opcionais)
    "Programa de Fidelidade": 299.90,
    "Integração Tap": 299.00,
    "Integração API": 299.00, # Preço do Gestão/Performance (era 199.90 antes)
    # Ajustar preços de BI e Backup para Performance se forem diferentes lá
    #"Business Intelligence (BI) - Performance": 99.00, # Se necessário criar entradas distintas
    #"Backup Realtime - Performance": 99.90, # Se necessário

    # --- Módulos Antigos (revisar se ainda são usados/precisam de preço) ---
    "60 Notas Fiscais": 40.00, # Não mencionado nos novos planos
    "120 Notas Fiscais": 70.00, # Não mencionado
    "250 Notas Fiscais": 90.00, # Não mencionado
    "3000 Notas Fiscais": 0.0, # Não mencionado (substituído por Ilimitadas ou 30)
    "Vendas - Estoque - Financeiro": 0.0, # Era mandatório, verificar se foi substituído por "Relatórios" ou similar
    "Delivery Direto Básico": 247.00, # Não mencionado
    "Painel de Senha": 49.90, # Substituído por TV/Mobile?
    "Domínio Próprio": 19.90, # Não mencionado
    "Gestão de Entregadores": 19.90, # Não mencionado
    "Robô de WhatsApp + Recuperador de Pedido": 99.90, # Não mencionado
    "Gestão de Redes Sociais": 9.90, # Não mencionado
    "Combo de Logística": 74.90, # Não mencionado
    "Painel MultiLojas": 199.00, # Não mencionado
    "Central Telefônica (Base)": 399.90, # Não mencionado
    "Central Telefônica (Por Loja)": 49.90 # Não mencionado
}

# --- Custos Adicionais (fixos por item extra) ---
PRECO_EXTRA_USUARIO = 19.00
PRECO_EXTRA_PDV_GESTAO_PERFORMANCE = 59.90
# PRECO_EXTRA_PDV_PDV_BASICO = ? # Não definido na descrição, assumir 0 ou não permitir

# ---------------------------------------------------------
# Função utilitária para substituir placeholders no Slide
# ---------------------------------------------------------
def substituir_placeholders_no_slide(slide, dados):
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    txt = run.text
                    for k, v in dados.items():
                        if k in txt:
                            # Garantir que v seja string
                            v_str = str(v) if v is not None else ""
                            txt = txt.replace(k, v_str)
                    run.text = txt


# ---------------------------------------------------------
# Classe PlanoFrame (Aba)
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

        # Variáveis compartilhadas
        self.nome_cliente_var = nome_cliente_var_shared
        self.validade_proposta_var = validade_proposta_var_shared
        self.nome_plano_var = tk.StringVar(value="") # Nome do plano (editável opcionalmente)

        self.current_plan = "PDV Básico" # Padrão inicial
        self.spin_pdv_var = tk.IntVar(value=1)
        self.spin_users_var = tk.IntVar(value=1)
        # Manter spinboxes antigos? Avaliar necessidade com base nos novos opcionais
        self.spin_auto_var = tk.IntVar(value=0) # Renomear para spin_terminais_auto?
        self.spin_cardapio_var = tk.IntVar(value=0) # Cardapio agora é checkbox?
        self.spin_tef_var = tk.IntVar(value=0) # TEF agora é checkbox?
        self.spin_smart_tef_var = tk.IntVar(value=0) # Smart TEF agora é checkbox/fixo?
        self.spin_app_cplug_var = tk.IntVar(value=0) # App Gestão agora é checkbox/fixo?
        # self.spin_delivery_direto_basico_var = tk.IntVar(value=0) # Não existe mais

        # Nota fiscal: Simplificado, pois agora é fixa (30 ou Ilimitada)
        # self.var_notas = tk.StringVar(value="NONE") # Remover ou adaptar se necessário

        # Módulos (checkboxes) - ATUALIZADO com novos módulos
        self.modules = {
            # Fixos (serão desabilitados, mas precisam da variável)
            "Usuários": tk.IntVar(),
            "30 Notas Fiscais": tk.IntVar(),
            "Suporte Técnico - Via chamados": tk.IntVar(),
            "Relatório Básico": tk.IntVar(),
            "PDV - Frente de Caixa": tk.IntVar(),
            "Notas Fiscais Ilimitadas": tk.IntVar(),
            "Importação de XML": tk.IntVar(),
            "Painel Senha TV": tk.IntVar(),
            "Estoque em Grade": tk.IntVar(),
            "Relatórios": tk.IntVar(),
            "Suporte Técnico - Via chat": tk.IntVar(),
            "Delivery": tk.IntVar(),
            "Relatório KDS": tk.IntVar(),
            "Produção": tk.IntVar(),
            "Promoções": tk.IntVar(),
            "Hub de Delivery": tk.IntVar(),
            "Ordem de Serviço": tk.IntVar(),
            "App Gestão CPlug": tk.IntVar(),
            "Painel Senha Mobile": tk.IntVar(),
            "Controle de Mesas": tk.IntVar(),
            "Marketing": tk.IntVar(),
            "Relatório Dinâmico": tk.IntVar(),
            "Atualização em tempo real": tk.IntVar(),
            "Facilita NFE": tk.IntVar(),
            "Conciliação Bancária": tk.IntVar(),
            "Contratos de cartões e outros": tk.IntVar(),
            "Suporte Técnico - Estendido": tk.IntVar(),
            "Smart TEF": tk.IntVar(),

            # Opcionais (reais checkboxes)
            "Smart Menu": tk.IntVar(),
            "Terminais Autoatendimento": tk.IntVar(), # Substitui spin_auto_var?
            "Delivery Direto Profissional": tk.IntVar(),
            "Delivery Direto VIP": tk.IntVar(),
            "TEF": tk.IntVar(), # Substitui spin_tef_var?
            "Cardápio digital": tk.IntVar(), # Substitui spin_cardapio_var?
            "Integração API": tk.IntVar(),
            "Business Intelligence (BI)": tk.IntVar(),
            "Backup Realtime": tk.IntVar(),
            "Programa de Fidelidade": tk.IntVar(),
            "Integração Tap": tk.IntVar(),

             # Antigos (manter se ainda usados em Bling/Autoatendimento ou remover)
            # "Vendas - Estoque - Financeiro": tk.IntVar(),
            # "Atualização em Tempo Real": tk.IntVar(), # Duplicado? Verificar nome exato
            # "Painel de Senha": tk.IntVar(), # Duplicado?
            # ... outros antigos ...
        }
        self.check_buttons = {} # Dicionário para guardar os widgets Checkbutton

        # Overrides de cálculo
        self.user_override_anual_active = tk.BooleanVar(value=False)
        self.user_override_discount_active = tk.BooleanVar(value=False)
        self.valor_anual_editavel = tk.StringVar(value="0.00")
        self.desconto_personalizado = tk.StringVar(value="0")

        # Armazenar valores calculados
        self.computed_mensal = 0.0
        self.computed_anual = 0.0
        self.computed_desconto_percent = 0.0
        self.computed_custo_adicional = 0.0 # Custo de Treinamento/Implantação unificado

        # --- Layout com Scrollbar (sem alterações) ---
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
        # --- Fim Layout Scrollbar ---

        # Construir a UI
        self._montar_layout_esquerda()
        self._montar_layout_direita()
        self.configurar_plano("PDV Básico") # Configura o plano inicial

    def fechar_aba(self):
        if self.on_close_callback:
            self.on_close_callback(self.aba_index)

    def on_bling_selected(self, event=None):
        selected_bling_plan = self.bling_var.get()
        if selected_bling_plan in LISTA_PLANOS_BLING:
            self.configurar_plano(selected_bling_plan)
        self.bling_var.set("Selecionar Bling...")

    def _montar_layout_esquerda(self):
        top_bar = ttkb.Frame(self.frame_left)
        top_bar.pack(fill="x", pady=5)
        ttkb.Label(top_bar, text=f"Aba Plano {self.aba_index}", font="-size 12 -weight bold").pack(side="left")
        btn_close = ttkb.Button(top_bar, text="Fechar Aba", command=self.fechar_aba)
        btn_close.pack(side="right")

        frame_planos = ttkb.Labelframe(self.frame_left, text="Planos")
        frame_planos.pack(fill="x", pady=5)

        self.bling_combobox = None
        for p in LISTA_PLANOS_UI:
            if p == "Bling":
                self.bling_var = tk.StringVar(value="Selecionar Bling...")
                self.bling_combobox = ttk.Combobox(frame_planos, textvariable=self.bling_var,
                                                   values=LISTA_PLANOS_BLING, state="readonly", width=25)
                self.bling_combobox.pack(side="left", padx=5)
                self.bling_combobox.bind("<<ComboboxSelected>>", self.on_bling_selected)
            else:
                ttkb.Button(frame_planos, text=p,
                            command=lambda pl=p: self.configurar_plano(pl)
                           ).pack(side="left", padx=5)

        # Frame Notas Fiscais removido (agora é fixo no plano)

        # Módulos (checkboxes) - Organização Mantida
        frame_mod = ttkb.Labelframe(self.frame_left, text="Módulos Opcionais") # Renomeado
        frame_mod.pack(fill="both", expand=True, pady=5)
        f_mod_cols = ttkb.Frame(frame_mod)
        f_mod_cols.pack(fill="both", expand=True)

        f_mod_left = ttkb.Frame(f_mod_cols)
        f_mod_left.pack(side="left", fill="both", expand=True, padx=5)
        f_mod_right = ttkb.Frame(f_mod_cols)
        f_mod_right.pack(side="left", fill="both", expand=True, padx=5)

        # Filtrar módulos para exibir apenas os que NÃO são sempre mandatórios
        # (Módulos como 'Usuários', 'PDV - Frente de Caixa' são controlados por spinbox)
        # (Módulos fixos específicos do plano serão desabilitados em configurar_plano)
        displayable_mods = sorted([
            m for m in self.modules.keys() if m not in [
                "Usuários", "PDV - Frente de Caixa", "30 Notas Fiscais",
                "Notas Fiscais Ilimitadas", "Suporte Técnico - Via chamados",
                "Relatório Básico", "Relatórios", "Suporte Técnico - Via chat",
                # Adicionar outros que são *sempre* fixos em algum plano e não devem ser checkboxes clicáveis
            ]
        ])

        mid = len(displayable_mods)//2
        left_side = displayable_mods[:mid]
        right_side = displayable_mods[mid:]
        self.check_buttons = {} # Limpa para garantir

        for m in left_side:
            if m in self.modules: # Garante que o módulo existe no dicionário
                 cb = ttk.Checkbutton(f_mod_left, text=m,
                                      variable=self.modules[m],
                                      command=self.atualizar_valores)
                 cb.pack(anchor="w", pady=2)
                 self.check_buttons[m] = cb # Armazena o widget

        for m in right_side:
             if m in self.modules: # Garante que o módulo existe no dicionário
                 cb = ttk.Checkbutton(f_mod_right, text=m,
                                      variable=self.modules[m],
                                      command=self.atualizar_valores)
                 cb.pack(anchor="w", pady=2)
                 self.check_buttons[m] = cb # Armazena o widget

        # Frame Dados Cliente e Plano (Mantido)
        frame_dados = ttkb.Labelframe(self.frame_left, text="Dados do Cliente e Proposta")
        frame_dados.pack(fill="x", pady=5)
        ttkb.Label(frame_dados, text="Nome do Cliente:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttkb.Entry(frame_dados, textvariable=self.nome_cliente_var, width=30).grid(row=0, column=1, padx=5, pady=2)
        ttkb.Label(frame_dados, text="Validade Proposta:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttkb.Entry(frame_dados, textvariable=self.validade_proposta_var, width=15).grid(row=1, column=1, padx=5, pady=2, sticky="w")
        ttkb.Label(frame_dados, text="Nome do Plano (Opcional):").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        ttkb.Entry(frame_dados, textvariable=self.nome_plano_var, width=30).grid(row=2, column=1, padx=5, pady=2)


    def _montar_layout_direita(self):
        frame_inc = ttkb.Labelframe(self.frame_right, text="Quantidades") # Renomeado
        frame_inc.pack(fill="x", pady=5)

        ttkb.Label(frame_inc, text="PDVs - Frente de Caixa").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        # Limites (to=...) serão ajustados dinamicamente em configurar_plano
        self.sp_pdv = ttkb.Spinbox(frame_inc, from_=0, to=99, # 'from' será ajustado
                              textvariable=self.spin_pdv_var, width=5,
                              command=self.atualizar_valores)
        self.sp_pdv.grid(row=0, column=1, padx=5, pady=2)

        ttkb.Label(frame_inc, text="Usuários").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        # Limites (to=...) serão ajustados dinamicamente em configurar_plano
        self.sp_usr = ttkb.Spinbox(frame_inc, from_=0, to=999, # 'from' será ajustado
                              textvariable=self.spin_users_var, width=5,
                              command=self.atualizar_valores)
        self.sp_usr.grid(row=1, column=1, padx=5, pady=2)

        # --- Remover Spinboxes antigos que viraram Checkboxes ---
        # ttkb.Label(frame_inc, text="Autoatendimento").grid(row=2, column=0, sticky="w") ...
        # ttkb.Label(frame_inc, text="Cardápio Digital").grid(row=3, column=0, sticky="w") ...
        # ttkb.Label(frame_inc, text="TEF").grid(row=4, column=0, sticky="w") ...
        # ttkb.Label(frame_inc, text="App Gestão CPlug").grid(row=6, column=0, sticky="w") ...
        # ttkb.Label(frame_inc, text="Delivery Direto Básico").grid(row=7, column=0, sticky="w") ...

        # --- Manter/Adaptar Spinbox de Smart TEF (Limitado no Gestão) ---
        ttkb.Label(frame_inc, text="Smart TEF").grid(row=2, column=0, sticky="w", padx=5, pady=2) # Ajustar row se necessário
        self.sp_smf = ttkb.Spinbox(frame_inc, from_=0, to=99, # 'to' será ajustado para 3 no Gestão
                              textvariable=self.spin_smart_tef_var, width=5,
                              command=self.atualizar_valores)
        self.sp_smf.grid(row=2, column=1, padx=5, pady=2) # Ajustar row

        # Frame Valores Finais (Mantido)
        frame_valores = ttkb.Labelframe(self.frame_right, text="Valores Finais")
        frame_valores.pack(fill="x", pady=5, expand=True)

        self.lbl_plano_mensal = ttkb.Label(frame_valores, text="Plano (Mensal): R$ 0,00", font="-size 12 -weight bold")
        self.lbl_plano_mensal.pack(pady=2)
        self.lbl_plano_anual = ttkb.Label(frame_valores, text="Plano (Anual): R$ 0,00", font="-size 12 -weight bold")
        self.lbl_plano_anual.pack(pady=2)
        self.lbl_treinamento = ttkb.Label(frame_valores, text="Custo Treinamento: R$ 0,00", font="-size 10") # Ajustado
        self.lbl_treinamento.pack(pady=2)
        self.lbl_desconto = ttkb.Label(frame_valores, text="Desconto Anual: 0%", font="-size 10") # Ajustado
        self.lbl_desconto.pack(pady=2)

        # Frame Edição Anual (Mantido)
        frame_edit_anual = ttkb.Labelframe(self.frame_right, text="Anual (editável)")
        frame_edit_anual.pack(pady=5, fill="x")
        e_anual = ttkb.Entry(frame_edit_anual, textvariable=self.valor_anual_editavel, width=10)
        e_anual.pack(side="left", padx=5)
        e_anual.bind("<KeyRelease>", self.on_user_edit_valor_anual)
        b_reset_anual = ttkb.Button(frame_edit_anual, text="Reset", command=self.on_reset_anual, width=5)
        b_reset_anual.pack(side="left", padx=5)

        # Frame Edição Desconto (Mantido)
        frame_edit_desc = ttkb.Labelframe(self.frame_right, text="Desconto (%) (editável)")
        frame_edit_desc.pack(pady=5, fill="x")
        e_desc = ttkb.Entry(frame_edit_desc, textvariable=self.desconto_personalizado, width=10)
        e_desc.pack(side="left", padx=5)
        e_desc.bind("<KeyRelease>", self.on_user_edit_desconto)
        b_reset_desc = ttkb.Button(frame_edit_desc, text="Reset", command=self.on_reset_desconto, width=5)
        b_reset_desc.pack(side="left", padx=5)

    # --- Funções de Edição/Reset (Mantidas) ---
    def on_user_edit_valor_anual(self, *args):
        self.user_override_anual_active.set(True)
        self.user_override_discount_active.set(False)
        self.atualizar_valores()

    def on_reset_anual(self):
        self.user_override_anual_active.set(False)
        self.valor_anual_editavel.set("0.00")
        self.atualizar_valores()

    def on_user_edit_desconto(self, *args):
        self.user_override_discount_active.set(True)
        self.user_override_anual_active.set(False)
        self.atualizar_valores()

    def on_reset_desconto(self):
        self.user_override_discount_active.set(False)
        self.desconto_personalizado.set("0")
        self.atualizar_valores()
    # --- Fim Funções Edição ---

    def configurar_plano(self, plano):
        # Reset Bling Combobox se outro plano for selecionado
        if not plano.startswith("Bling -") and self.bling_combobox:
             self.bling_var.set("Selecionar Bling...")

        if plano not in PLAN_INFO:
             showerror("Erro de Configuração", f"Plano '{plano}' não encontrado na definição.")
             return

        info = PLAN_INFO[plano]
        self.current_plan = plano

        # Configurar Spinboxes (Mínimos e Máximos)
        min_pdv = info.get("min_pdv", 0)
        min_users = info.get("min_users", 0)
        max_pdv = min_pdv + info.get("max_extra_pdvs", 99) # 99 como default alto
        max_users = min_users + info.get("max_extra_users", 999) # 999 como default alto

        self.spin_pdv_var.set(min_pdv)
        self.sp_pdv.config(from_=min_pdv, to=max_pdv)
        self.spin_users_var.set(min_users)
        self.sp_usr.config(from_=min_users, to=max_users)

        # Limite específico para Smart TEF no plano Gestão
        if plano == "Gestão":
             self.sp_smf.config(to=3)
             self.spin_smart_tef_var.set(0) # Resetar valor ao trocar para gestão
        else:
             # Precisa definir um 'to' padrão para outros planos, ou ler do PLAN_INFO se existir
             default_max_smart_tef = 99 # Ou outro valor padrão
             self.sp_smf.config(to=default_max_smart_tef)
             # Resetar se o plano Performance já incluir (precisa checar 'mandatory')
             if "Smart TEF" in info.get("mandatory", []):
                 # O plano performance inclui 3 Smart TEF fixos
                 # A spinbox deveria controlar os *extras* ou o *total*?
                 # Se total, o 'from_' deveria ser 3 e 'to' 3 + extras permitidos (se houver)
                 # Se extras, 'from_' 0 e 'to' extras permitidos.
                 # Por simplicidade, vamos manter a spinbox controlando o total e
                 # ajustar o cálculo de custo em atualizar_valores.
                 # No Performance, inclui 3. Opcionais não listam extras.
                 self.sp_smf.config(from_=3, to=3) # Fixa em 3 no performance
                 self.spin_smart_tef_var.set(3)
             else:
                 # Outros planos (PDV Básico), Smart TEF é opcional normal
                 self.sp_smf.config(from_=0) # Começa do zero
                 self.spin_smart_tef_var.set(0) # Resetar valor

        # Resetar Spinboxes antigos (se ainda existirem no layout)
        self.spin_auto_var.set(0)
        self.spin_cardapio_var.set(0)
        self.spin_tef_var.set(0)
        self.spin_app_cplug_var.set(0)

        # Resetar todos os módulos/checkboxes
        for m, var in self.modules.items():
            var.set(0)
            if m in self.check_buttons:
                self.check_buttons[m].config(state='disabled') # Começa desabilitado

        # Habilitar opcionais permitidos e marcar/desabilitar mandatórios
        allowed = info.get("allowed_optionals", [])
        mandatory = info.get("mandatory", [])

        for m, var in self.modules.items():
            is_mandatory = m in mandatory
            is_allowed_optional = m in allowed

            if is_mandatory:
                var.set(1)
                if m in self.check_buttons:
                    self.check_buttons[m].config(state='disabled')
                # Lógica especial para itens controlados por spinbox que são mandatórios
                # Ex: Usuários, PDV, Smart TEF (no Performance)
                # A quantidade já foi setada nos spinboxes acima.
            elif is_allowed_optional:
                if m in self.check_buttons:
                    self.check_buttons[m].config(state='normal') # Habilita
            else:
                # Não é mandatório nem opcional permitido para este plano
                if m in self.check_buttons:
                    self.check_buttons[m].config(state='disabled') # Mantém desabilitado

        # Resetar overrides
        self.user_override_anual_active.set(False)
        self.user_override_discount_active.set(False)
        self.valor_anual_editavel.set("0.00")
        self.desconto_personalizado.set("0")

        self.atualizar_valores()


    def atualizar_valores(self, *args):
        if not self.current_plan or self.current_plan not in PLAN_INFO:
            return # Segurança

        info = PLAN_INFO[self.current_plan]
        is_bling_plan = self.current_plan.startswith("Bling -")
        is_autoatendimento_plano = self.current_plan == "Autoatendimento"
        is_em_branco_plano = self.current_plan == "Em Branco"
        mandatory = info.get("mandatory", [])
        base_mensal = info.get("base_mensal", 0.0)

        # --- Calcular Custo Total dos Extras ---
        total_extras_cost = 0.0
        total_extras_descontavel = 0.0 # Parte dos extras que recebe desconto anual
        total_extras_nao_descontavel = 0.0 # Parte dos extras sem desconto (TEF, etc.)

        # 1. PDVs Extras
        pdv_atuais = self.spin_pdv_var.get()
        pdv_incluidos = info.get("min_pdv", 0)
        pdv_extras = max(0, pdv_atuais - pdv_incluidos)
        pdv_price = 0.0
        if self.current_plan in ["Gestão", "Performance"]:
            pdv_price = PRECO_EXTRA_PDV_GESTAO_PERFORMANCE
        # Adicionar preço para Bling/Autoatendimento se diferente
        elif is_bling_plan:
            pdv_price = 40.00 # Preço antigo do Bling, verificar se mantém
        # Ignorar PDV extra para PDV Básico (não definido)

        cost_pdv_extra = pdv_extras * pdv_price
        total_extras_cost += cost_pdv_extra
        # Assumir que PDV extra é descontável, a menos que especificado o contrário
        # Adicionar "PDV Extra" a SEM_DESCONTO se necessário
        if "PDV Extra" not in SEM_DESCONTO:
             total_extras_descontavel += cost_pdv_extra
        else:
             total_extras_nao_descontavel += cost_pdv_extra

        # 2. Users Extras
        users_atuais = self.spin_users_var.get()
        users_incluidos = info.get("min_users", 0)
        users_extras = max(0, users_atuais - users_incluidos)
        user_price = PRECO_EXTRA_USUARIO

        cost_users_extra = users_extras * user_price
        total_extras_cost += cost_users_extra
        # Assumir que User extra é descontável
        if "User Extra" not in SEM_DESCONTO:
             total_extras_descontavel += cost_users_extra
        else:
             total_extras_nao_descontavel += cost_users_extra

        # 3. Módulos Extras (Checkboxes selecionados que NÃO são mandatórios)
        for m, var_m in self.modules.items():
            if var_m.get() == 1 and m not in mandatory:
                # Módulos como Smart TEF no Performance são mandatórios E controlados por spinbox
                # O custo deles deve ser calculado pelo spinbox, não aqui.
                if m == "Smart TEF" and self.current_plan == "Performance":
                    continue # Custo tratado abaixo

                price = precos_mensais.get(m, 0.0)
                total_extras_cost += price
                if m not in SEM_DESCONTO:
                    total_extras_descontavel += price
                else:
                    total_extras_nao_descontavel += price

        # 4. Spinboxes Extras (Itens quantificáveis opcionais)
        #    - Smart TEF (Apenas contar extras além do incluído no Performance)
        smart_tef_atuais = self.spin_smart_tef_var.get()
        smart_tef_incluidos = 0
        if self.current_plan == "Performance":
             smart_tef_incluidos = 3 # Performance inclui 3

        smart_tef_extras = max(0, smart_tef_atuais - smart_tef_incluidos)
        if smart_tef_extras > 0:
             price = smart_tef_extras * precos_mensais.get("Smart TEF", 0.0)
             total_extras_cost += price
             if "Smart TEF" not in SEM_DESCONTO:
                 total_extras_descontavel += price
             else:
                 total_extras_nao_descontavel += price

        #    - Outros Spinboxes (se houver algum relevante restante)
        #    Ex: Terminais Autoatendimento (se fosse spinbox e não checkbox)
        #    auto_qty = self.spin_auto_var.get() ... etc.

        # --- Calcular Valor Mensal Potencial (Base + Extras) ---
        valor_mensal_potencial = base_mensal + total_extras_cost

        # Lógica específica para Bling (se mantida)
        if is_bling_plan:
            base_mensal_orig = info.get("base_mensal_original", 0.0) # Para exibição "De..."
            # O valor mensal *efetivo* no Bling não é claro, focar no anual
            valor_mensal_potencial = base_mensal_orig + total_extras_cost # Valor "De" + extras

        # Lógica específica para Autoatendimento (se mantida)
        if is_autoatendimento_plano:
             valor_mensal_potencial = 0.0 # Não tem mensal direto


        # --- Calcular Custo Adicional Unificado (Implementação/Treinamento) ---
        # Regra antiga: Se mensal < 549.90, custo = 549.90 - mensal
        # Manter essa regra? Ou os novos planos têm custo fixo de implantação?
        # Assumindo a regra antiga por enquanto para PDV Básico, Gestão, Performance
        custo_adicional = 0.0
        label_custo = "Treinamento" # Padrão
        if is_bling_plan: label_custo = "Implementação"

        # Aplicar regra apenas se não for Bling, Autoatendimento ou Em Branco
        if not is_bling_plan and not is_autoatendimento_plano and not is_em_branco_plano:
            limite_custo = 549.90 # Limite antigo
            if valor_mensal_potencial > 0 and valor_mensal_potencial < limite_custo:
                custo_adicional = limite_custo - valor_mensal_potencial

        # --- Calcular Valor Anual Efetivo ---
        final_anual = 0.0
        desconto_aplicado_percent = 0.0 # Para exibir

        if is_bling_plan:
            base_anual_rate = info.get("base_anual", 0.0) # Valor MENSAL efetivo no anual
            final_anual = base_anual_rate + total_extras_cost # Bling = Base Fixa + Extras (sem desconto padrão)
            # Calcular desconto implícito para exibição
            if valor_mensal_potencial > 0:
                 desconto_aplicado_percent = ((valor_mensal_potencial - final_anual) / valor_mensal_potencial) * 100
            # Override Bling Anual
            if self.user_override_anual_active.get():
                try: final_anual = float(self.valor_anual_editavel.get())
                except ValueError: pass
                # Recalcular % de desconto se o valor foi editado
                if valor_mensal_potencial > 0:
                     desconto_aplicado_percent = ((valor_mensal_potencial - final_anual) / valor_mensal_potencial) * 100

        elif is_autoatendimento_plano:
             # Lógica específica para Autoatendimento (mantida do original, revisar)
             base_anual_rate = info.get("base_anual", 419.90) # Custo MENSAL para 1 terminal no anual
             auto_module_name = "Terminais Autoatendimento" # Usar o nome do módulo opcional
             auto_qty = 0
             if self.modules[auto_module_name].get() == 1: auto_qty = 1 # Se o checkbox está marcado
             # Adicionar lógica para ler quantidade de um spinbox se Terminais Autoatendimento voltar a ser spinbox
             # auto_qty = self.spin_auto_var.get()

             if auto_qty < 1: auto_qty = 1 # Mínimo 1 se o plano foi escolhido? Ou 0? Assumir 0 se não marcado.

             if auto_qty >= 1:
                 preco_auto_extra = precos_mensais.get(auto_module_name, 199.00) # Usar o preço do módulo opcional
                 # O cálculo original era confuso. Simplificando:
                 # Custo anual = Qtd * Preço Mensal do Módulo * 12 (com algum desconto?)
                 # Ou usar a base_anual do plano + extras?
                 # Assumindo que base_anual (419.90) é o preço MENSAL para 1 terminal no contrato ANUAL
                 # E extras custam o preço MENSAL do módulo (199.00)
                 final_anual = base_anual_rate + max(0, auto_qty - 1) * preco_auto_extra
             else:
                 final_anual = 0.0 # Nenhum terminal selecionado

             # Override Anual
             if self.user_override_anual_active.get():
                try: final_anual = float(self.valor_anual_editavel.get())
                except ValueError: pass

        else: # Planos Padrão (PDV Básico, Gestão, Performance, Em Branco)
            # Calcular parte descontável total (Base Mensal + Extras descontáveis)
            total_descontavel_calc = base_mensal + total_extras_descontavel
            total_nao_descontavel_calc = total_extras_nao_descontavel

            # Calcular Anual com base nos descontos/overrides
            if self.user_override_anual_active.get():
                # Valor anual foi digitado diretamente
                try:
                    final_anual = float(self.valor_anual_editavel.get())
                    # Calcular % de desconto efetivo para exibição
                    valor_mensal_total_sem_desconto = total_descontavel_calc + total_nao_descontavel_calc
                    if valor_mensal_total_sem_desconto > 0:
                        desconto_aplicado_percent = ((valor_mensal_total_sem_desconto - final_anual) / valor_mensal_total_sem_desconto) * 100
                except ValueError:
                    # Fallback se valor digitado for inválido: aplicar 10% padrão
                    desc_padrao = 0.10
                    final_anual = (total_descontavel_calc * (1 - desc_padrao)) + total_nao_descontavel_calc
                    desconto_aplicado_percent = desc_padrao * 100

            elif self.user_override_discount_active.get():
                # Desconto percentual foi digitado
                try:
                    desc_custom = float(self.desconto_personalizado.get())
                except ValueError: desc_custom = 0.0
                desc_dec = desc_custom / 100.0
                final_anual = (total_descontavel_calc * (1 - desc_dec)) + total_nao_descontavel_calc
                desconto_aplicado_percent = desc_custom
            else:
                # Desconto Padrão 10% (se aplicável)
                # A descrição não menciona mais desconto anual padrão. Remover?
                # Por segurança, manter o cálculo de 10% como padrão se nada for editado.
                desc_padrao = 0.10
                final_anual = (total_descontavel_calc * (1 - desc_padrao)) + total_nao_descontavel_calc
                desconto_aplicado_percent = desc_padrao * 100

        # --- Atualizar Campo Editável Anual ---
        # Formatar para evitar problemas com locale (ponto como decimal)
        self.valor_anual_editavel.set(f"{final_anual:.2f}")

        # --- Atualizar Labels da UI ---
        # Label Mensal
        mensal_str = ""
        if is_autoatendimento_plano:
            mensal_str = "Plano (Mensal): Não disponível"
        elif valor_mensal_potencial >= 0: # Exibe para todos os outros
            mensal_pot_str = f"{valor_mensal_potencial:.2f}".replace(".", ",")
            if custo_adicional > 0:
                custo_adic_str = f"{custo_adicional:.2f}".replace(".", ",")
                mensal_str = f"Plano Mensal: R$ {mensal_pot_str} + R$ {custo_adic_str} ({label_custo})"
            else:
                mensal_str = f"Plano Mensal: R$ {mensal_pot_str}"
        self.lbl_plano_mensal.config(text=mensal_str)

        # Label Anual
        anual_str = f"{final_anual:.2f}".replace(".", ",")
        self.lbl_plano_anual.config(text=f"Plano (Anual): R$ {anual_str}")

        # Label Custo Adicional (Treinamento/Implementação)
        custo_adic_str_lbl = f"{custo_adicional:.2f}".replace(".", ",")
        if custo_adicional > 0:
             self.lbl_treinamento.config(text=f"Custo {label_custo}: R$ {custo_adic_str_lbl}")
        else:
             self.lbl_treinamento.config(text="") # Oculta se for zero

        # Label Desconto
        desconto_final_percent = max(0, desconto_aplicado_percent) # Garante não ser negativo
        # Atualizar campo de edição se o desconto foi calculado (não editado)
        if not self.user_override_discount_active.get() and not is_bling_plan and not is_autoatendimento_plano:
             self.desconto_personalizado.set(f"{round(desconto_final_percent)}")

        self.lbl_desconto.config(text=f"Desconto Anual: {round(desconto_final_percent)}%")

        # --- Armazenar Valores Computados Finais ---
        self.computed_mensal = valor_mensal_potencial if not is_autoatendimento_plano else 0.0
        self.computed_anual = final_anual
        self.computed_desconto_percent = round(desconto_final_percent)
        self.computed_custo_adicional = custo_adicional


    def montar_lista_modulos(self):
        """ Cria a string formatada para o slide com os módulos incluídos. """
        linhas = []
        info = PLAN_INFO.get(self.current_plan, {})
        mandatory = info.get("mandatory", [])

        # 1. Itens Quantificáveis (PDV, Usuários, Smart TEF)
        pdv_val = self.spin_pdv_var.get()
        if pdv_val > 0:
            linhas.append(f"{pdv_val}x PDV - Frente de Caixa")

        usr_val = self.spin_users_var.get()
        if usr_val > 0:
            # Adicionar a lógica de usuário extra/limite aqui se necessário mostrar no slide
            # Ex: "3x Usuários (Limite: 5)"
            linhas.append(f"{usr_val}x Usuários")
            # Lógica de usuário cortesia por PDV extra foi removida da descrição
            # min_pdv = info.get("min_pdv", 0)
            # pdv_extras = max(0, pdv_val - min_pdv)
            # if pdv_extras > 0 and self.current_plan in ["PDV Básico", "Gestão", "Performance"]:
            #    linhas.append(f"{pdv_extras}x Usuário Cortesia por PDV Extra")


        smart_tef_val = self.spin_smart_tef_var.get()
        if smart_tef_val > 0:
             # No Performance, 3 são incluídos. Mostrar como fixo ou pela quantidade?
             # Mostrar pela quantidade selecionada na spinbox para consistência.
             linhas.append(f"{smart_tef_val}x Smart TEF")
             if self.current_plan == "Gestão":
                 linhas[-1] += " (Limite: 3)" # Adiciona o limite visualmente

        # 2. Módulos Fixos (Mandatórios que não são contadores)
        for m in mandatory:
            if m not in ["PDV - Frente de Caixa", "Usuários", "Smart TEF"]: # Já tratados acima
                 # Adicionar "1x" para clareza?
                 linhas.append(f"1x {m}")

        # 3. Módulos Opcionais Selecionados (Checkboxes)
        for m, var_m in self.modules.items():
            if var_m.get() == 1 and m not in mandatory:
                # Se for um módulo controlado por spinbox (que virou checkbox?) tratar diferente?
                # Ex: TEF, Terminais Autoatendimento, Cardápio Digital
                if m in ["TEF", "Terminais Autoatendimento", "Cardápio digital", "Programa de Fidelidade", "Integração Tap", "Integração API", "Business Intelligence (BI)", "Backup Realtime", "Smart Menu", "Delivery Direto Profissional", "Delivery Direto VIP"]:
                     linhas.append(f"1x {m}") # Assumindo 1 unidade ao marcar checkbox
                # Ignorar módulos já listados como fixos ou contadores
                elif m not in ["PDV - Frente de Caixa", "Usuários", "Smart TEF"]:
                     linhas.append(f"1x {m}")


        # Remover duplicados (preservando ordem o máximo possível)
        unique_mods = []
        for mod in linhas:
            if mod not in unique_mods:
                unique_mods.append(mod)

        # Formatar para o slide
        montagem = "\n".join(f"•    {m}" for m in unique_mods)
        return montagem

    def gerar_dados_proposta(self, nome_closer, cel_closer, email_closer):
            """ Gera o dicionário de dados para preencher o slide da proposta. """
            nome_plano_selecionado = self.current_plan # Nome base do plano (PDV Básico, Gestão, etc.)
            nome_plano_editado = self.nome_plano_var.get().strip()
            # Usar nome editado se preenchido, senão usar o nome do plano selecionado
            nome_plano_final = nome_plano_editado if nome_plano_editado else nome_plano_selecionado

            valor_anual_efetivo = self.computed_anual
            valor_mensal_potencial = self.computed_mensal
            custo_adicional = self.computed_custo_adicional
            desconto_percent = self.computed_desconto_percent
            is_bling_plan = self.current_plan.startswith("Bling -")
            is_autoatendimento_plano = self.current_plan == "Autoatendimento"

            # --- Formatar String Plano Mensal ---
            plano_mensal_str = "Não Aplicável" # Padrão
            label_custo = "Treinamento" if not is_bling_plan else "Implementação"
            if not is_autoatendimento_plano and valor_mensal_potencial > 0:
                mensal_pot_str = f"{valor_mensal_potencial:.2f}".replace(".", ",")
                if custo_adicional > 0:
                    custo_adic_str = f"{custo_adicional:.2f}".replace(".", ",")
                    plano_mensal_str = f"R$ {mensal_pot_str} + R$ {custo_adic_str} ({label_custo})"
                else:
                    plano_mensal_str = f"R$ {mensal_pot_str}"
            elif is_autoatendimento_plano:
                 plano_mensal_str = "Não Disponível"


            # --- Formatar String Plano Anual ---
            plano_anual_str = f"R$ {valor_anual_efetivo:.2f}".replace(".", ",")

            # --- Definir Suporte ---
            # Regra antiga baseada no valor anual. Manter ou adaptar?
            # Verificar se os novos módulos de Suporte (chat, estendido) mudam isso.
            # Assumindo regra antiga por enquanto.
            tipo_suporte = "Regular"
            horario_suporte = "09:00 às 17:00 de Segunda a Sexta-feira"
            # Checar se módulos de suporte avançado foram incluídos
            suporte_chat = self.modules.get("Suporte Técnico - Via chat", tk.IntVar()).get() == 1
            suporte_estendido = self.modules.get("Suporte Técnico - Estendido", tk.IntVar()).get() == 1

            if suporte_estendido:
                 tipo_suporte = "Estendido"
                 horario_suporte = "09:00 às 22:00 de Segunda a Sexta-feira & Sábado e Domingo das 11:00 às 21:00"
            elif suporte_chat: # Se tem chat mas não estendido, qual horário? Assumir estendido também?
                 tipo_suporte = "Chat Incluso" # Ou manter "Estendido"?
                 horario_suporte = "09:00 às 22:00 de Segunda a Sexta-feira & Sábado e Domingo das 11:00 às 21:00" # Suposição
            # Fallback para a regra de valor se nenhum módulo específico for marcado?
            # elif valor_anual_efetivo >= 269.90:
            #    tipo_suporte = "Estendido"
            #    horario_suporte = "09:00 às 22:00 de Segunda a Sexta-feira & Sábado e Domingo das 11:00 às 21:00"

            # --- Montar Lista de Módulos ---
            montagem = self.montar_lista_modulos()

            # --- Calcular Economia Anual ---
            economia_str = ""
            if not is_autoatendimento_plano and valor_mensal_potencial > 0:
                 # Custo total em 12x mensal = (Mensal Potencial * 12) + Custo Adicional (pago 1x)
                 custo_total_mensalizado = (valor_mensal_potencial * 12) + custo_adicional
                 # Custo total no anual = Valor Anual Efetivo * 12 (ou apenas o valor anual se for taxa única?)
                 # Assumindo que computed_anual é o valor MENSAL no plano anual
                 custo_total_anualizado = valor_anual_efetivo * 12

                 economia_val = custo_total_mensalizado - custo_total_anualizado
                 if economia_val > 0.01: # Adicionar pequena margem para evitar mostrar R$ 0,00
                    econ = f"{economia_val:.2f}".replace(".", ",")
                    economia_str = f"Economia de R$ {econ} ao ano"

            # --- Montar Dicionário Final ---
            dados = {
                "montagem_do_plano": montagem,
                "plano_mensal": plano_mensal_str,
                "plano_anual": plano_anual_str,
                "desconto_total": f"{desconto_percent}%", # Desconto calculado para o anual
                "nome_do_plano": nome_plano_final, # Usa o nome editado ou o do plano
                "tipo_de_suporte": tipo_suporte,
                "horario_de_suporte": horario_suporte,
                "validade_proposta": self.validade_proposta_var.get(),
                "nome_closer": nome_closer,
                "celular_closer": cel_closer,
                "email_closer": email_closer,
                "nome_cliente": self.nome_cliente_var.get(),
                "economia_anual": economia_str
            }
            return dados

# ---------------------------------------------------------
# Funções que geram .pptx (Proposta e Material) - SEM ALTERAÇÕES SIGNIFICATIVAS NA LÓGICA INTERNA
# A função 'gerar_dados_proposta' acima já fornece os dados atualizados.
# A função 'montar_lista_modulos' fornece a lista de módulos atualizada.
# A lógica de mapeamento de slides em 'gerar_material' pode precisar de revisão
# se os placeholders ('check_tef', 'slide_bling', etc.) mudaram nos seus templates PPTX.
# ---------------------------------------------------------

# --- Função gerar_proposta (sem mudanças na lógica interna, usa dados da aba) ---
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

    # Pega dados da primeira aba ativa como fallback e para informações globais
    primeira_aba = lista_abas[0]
    dados_primeira_aba = primeira_aba.gerar_dados_proposta(nome_closer, celular_closer, email_closer)

    # Substitui placeholders em todos os slides usando os dados da primeira aba
    # (A lógica original de mapear slides para abas específicas foi removida por simplicidade,
    #  mas pode ser readicionada se necessário ter propostas com múltiplos planos diferentes)
    for slide in prs.slides:
        substituir_placeholders_no_slide(slide, dados_primeira_aba)

    # Salvar
    nome_cliente_primeira = dados_primeira_aba.get("nome_cliente", "SemNome").replace("/", "-").replace("\\", "-") # Evitar barras no nome
    hoje_str = date.today().strftime("%d-%m-%Y")
    nome_arquivo = f"Proposta ConnectPlug - {nome_cliente_primeira} - {hoje_str}.pptx"

    try:
        prs.save(nome_arquivo)
        showinfo("Sucesso", f"Proposta gerada: {nome_arquivo}")
        return nome_arquivo
    except Exception as e:
        showerror("Erro", f"Falha ao salvar '{nome_arquivo}': {e}")
        return None

# --- Função gerar_material (Requer atenção ao MAPEAMENTO_MODULOS) ---
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

    # 1) Coletar Módulos Ativos e Planos Usados de TODAS as abas
    modulos_ativos_geral = set()
    planos_usados_geral = set()
    dados_primeira_aba = lista_abas[0].gerar_dados_proposta(nome_closer, celular_closer, email_closer) # Para placeholders globais

    for aba in lista_abas:
        planos_usados_geral.add(aba.current_plan)
        info_aba = PLAN_INFO.get(aba.current_plan, {})
        mandatory_aba = info_aba.get("mandatory", [])

        # Adiciona mandatórios da aba
        for mod in mandatory_aba:
             modulos_ativos_geral.add(mod)

        # Adiciona opcionais selecionados (checkboxes)
        for nome_mod, var_mod in aba.modules.items():
            if var_mod.get() == 1 and nome_mod not in mandatory_aba:
                 modulos_ativos_geral.add(nome_mod)

        # Adiciona itens de Spinboxes (se relevantes para mostrar no material)
        if aba.spin_pdv_var.get() > 0: modulos_ativos_geral.add("PDV - Frente de Caixa") # Nome consistente
        if aba.spin_users_var.get() > 0: modulos_ativos_geral.add("Usuários")
        if aba.spin_smart_tef_var.get() > 0: modulos_ativos_geral.add("Smart TEF")
        # Adicionar outros spinboxes se necessário


    # 2) MAPEAMENTO_MODULOS: *** CRÍTICO - AJUSTAR CONFORME SEU PPTX ***
    #    Associa um texto/placeholder no seu slide PPTX com o nome interno do módulo no Python.
    #    Se o módulo estiver ativo (em modulos_ativos_geral), o slide com o placeholder é mantido.
    MAPEAMENTO_MODULOS = {
        # Placeholder no PPTX : Nome do Módulo no Python (chave em precos_mensais/self.modules)
        "slide_sempre": None, # Slides que devem sempre aparecer
        "check_sistema_kds": "Relatório KDS",
        "check_Hub_de_Delivery": "Hub de Delivery",
        "check_integracao_api": "Integração API",
        "check_integracao_tap": "Integração Tap", # Novo? Verificar placeholder
        "check_controle_de_mesas": "Controle de Mesas",
        "check_Delivery": "Delivery",
        "check_producao": "Produção",
        "check_Estoque_em_Grade": "Estoque em Grade",
        "check_Facilita_NFE": "Facilita NFE",
        "check_Importacao_de_xml": "Importação de XML", # Verificar case
        "check_conciliacao_bancaria": "Conciliação Bancária",
        "check_contratos_de_cartoes": "Contratos de cartões e outros",
        "check_ordem_de_servico": "Ordem de Serviço",
        "check_relatorio_dinamico": "Relatório Dinâmico",
        "check_programa_de_fidelidade": "Programa de Fidelidade", # Novo? Verificar placeholder
        "check_business_intelligence": "Business Intelligence (BI)",
        "check_smartmenu": "Smart Menu", # Verificar placeholder
        "check_backup_real_time": "Backup Realtime", # Verificar placeholder
        "check_att_tempo_real": "Atualização em tempo real", # Verificar placeholder
        "check_promocao": "Promoções", # Verificar placeholder
        "check_marketing": "Marketing", # Verificar placeholder
        "placeholder_pdv": "PDV - Frente de Caixa", # Renomeado
        "placeholder_smarttef": "Smart TEF", # Verificar placeholder
        "placeholder_tef": "TEF", # Verificar placeholder
        "placeholder_autoatendimento": "Terminais Autoatendimento", # Renomeado
        "placeholder_cardapio_digital": "Cardápio digital", # Renomeado
        "placeholder_app_gestao_cplug": "App Gestão CPlug", # Verificar placeholder
        "check_delivery_direto_vip": "Delivery Direto VIP", # Verificar placeholder
        "check_delivery_direto_profissional": "Delivery Direto Profissional", # Verificar placeholder
        "placeholder_painel_senha_tv": "Painel Senha TV", # Novo
        "placeholder_painel_senha_mobile": "Painel Senha Mobile", # Novo
        "placeholder_suporte_chat": "Suporte Técnico - Via chat", # Novo
        "placeholder_suporte_estendido": "Suporte Técnico - Estendido", # Novo
        "placeholder_notas_fiscais": { # Placeholder genérico para qualquer tipo de NF ativa
            "Notas Fiscais Ilimitadas",
            "30 Notas Fiscais"
            # Adicionar outros tipos se existirem (60, 120, 250?)
        },
         # Adicionar mapeamentos para TODOS os módulos relevantes que têm slides específicos
         # Se um módulo não tem slide dedicado, não precisa mapear
    }

    # 3) Decidir quais slides manter
    keep_slides = set()
    for i, slide in enumerate(prs.slides):
        slide_mantido = False
        # Procurar por placeholders no texto do slide
        for shape in slide.shapes:
             if slide_mantido: break # Já decidiu manter, vai para o próximo slide
             if shape.has_text_frame:
                 for paragraph in shape.text_frame.paragraphs:
                     if slide_mantido: break
                     for run in paragraph.runs:
                         if slide_mantido: break
                         txt_run = run.text.strip()
                         if not txt_run: continue

                         # Checa se algum placeholder do mapeamento está no texto
                         for placeholder, modulo_mapeado in MAPEAMENTO_MODULOS.items():
                             if placeholder in txt_run:
                                 if modulo_mapeado is None: # "slide_sempre"
                                     slide_mantido = True
                                     break
                                 # Se for um conjunto de módulos (como Notas Fiscais)
                                 elif isinstance(modulo_mapeado, set):
                                     # Manter se QUALQUER um dos módulos do conjunto estiver ativo
                                     if any(m in modulos_ativos_geral for m in modulo_mapeado):
                                         slide_mantido = True
                                         break
                                 # Se for um módulo único
                                 elif modulo_mapeado in modulos_ativos_geral:
                                     slide_mantido = True
                                     break
                         # Checa condições específicas de plano (ex: Bling)
                         if "slide_bling" in txt_run and any(p.startswith("Bling -") for p in planos_usados_geral):
                              slide_mantido = True
                              break


        if slide_mantido:
            keep_slides.add(i)

    # 4) Remove slides não mantidos
    for idx in reversed(range(len(prs.slides))):
        if idx not in keep_slides:
            try:
                rId = prs.slides._sldIdLst[idx].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[idx]
            except Exception as e:
                print(f"Aviso: Não foi possível remover slide {idx}. Erro: {e}")


    # 5) Substituir placeholders globais (usando dados da primeira aba)
    for slide in prs.slides:
        substituir_placeholders_no_slide(slide, dados_primeira_aba)

    # 6) Salvar pptx final
    nome_cliente_primeira = dados_primeira_aba.get("nome_cliente", "SemNome").replace("/", "-").replace("\\", "-")
    hoje_str = date.today().strftime("%d-%m-%Y")
    nome_arquivo = f"Material Tecnico ConnectPlug - {nome_cliente_primeira} - {hoje_str}.pptx"

    try:
        prs.save(nome_arquivo)
        showinfo("Sucesso", f"Material Técnico gerado: {nome_arquivo}")
        return nome_arquivo
    except Exception as e:
        showerror("Erro", f"Falha ao salvar '{nome_arquivo}': {e}")
        return None


# ---------------------------------------------------------
# Google Drive / Auth / Upload (SEM ALTERAÇÕES)
# ---------------------------------------------------------
SCOPES = ['https://www.googleapis.com/auth/drive']

def baixar_client_secret_remoto():
    url = "https://github.com/DevRGS/Gerador/raw/refs/heads/main/config/client_secret_788265418970-ur6f189oqvsttseeg6g77fegt0su67dj.apps.googleusercontent.com.json"
    nome_local = "client_secret_temp.json"
    if not os.path.exists(nome_local):
        print("Baixando client_secret do GitHub...")
        try:
            r = requests.get(url, timeout=10) # Adiciona timeout
            r.raise_for_status() # Levanta erro para status ruim (404, 500, etc)
            with open(nome_local, "w", encoding="utf-8") as f:
                f.write(r.text)
        except requests.exceptions.RequestException as e:
             showerror("Erro de Rede", f"Não foi possível baixar o client_secret.json: {e}")
             raise Exception(f"Erro ao baixar o client_secret.json: {e}") # Para o programa
    return nome_local

def get_gdrive_service():
    creds = None
    token_file = 'token.json'
    client_secret_file = None

    try:
        client_secret_file = baixar_client_secret_remoto()
    except Exception as e:
        showerror("Erro Crítico", f"Falha ao obter client_secret necessário para Google Drive.\n{e}")
        return None # Retorna None se não conseguir baixar

    if os.path.exists(token_file):
        try:
            with open(token_file, 'rb') as token:
                creds = pickle.load(token)
        except (pickle.UnpicklingError, EOFError):
             print("Arquivo token.json corrompido ou vazio. Removendo.")
             os.remove(token_file)
             creds = None


    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"Erro ao renovar token: {e}. Reautenticação necessária.")
                # Tentar remover token antigo para forçar fluxo novo
                if os.path.exists(token_file): os.remove(token_file)
                creds = None # Força o fluxo de autenticação
        # else: # Comentado pois o bloco abaixo cuidará disso
        #     pass

        # Se ainda não tem credenciais (novo ou refresh falhou)
        if not creds:
             try:
                flow = InstalledAppFlow.from_client_secrets_file(client_secret_file, SCOPES)
                # Tenta porta 0 (automática), mas pode falhar em alguns sistemas.
                # Se falhar, pode tentar uma porta fixa (ex: 8080), mas requer que ela esteja livre.
                creds = flow.run_local_server(port=0)
             except Exception as e:
                 showerror("Erro de Autenticação", f"Falha ao autenticar com Google Drive: {e}")
                 return None

        # Salva o token (novo ou renovado)
        try:
            with open(token_file, 'wb') as token:
                pickle.dump(creds, token)
        except Exception as e:
             print(f"Aviso: Não foi possível salvar o token.json: {e}")

    # Constrói o serviço Google Drive autenticado
    try:
        service = build('drive', 'v3', credentials=creds)
        return service
    except Exception as e:
        showerror("Erro Google API", f"Falha ao construir serviço do Google Drive: {e}")
        return None


def upload_pptx_and_export_to_pdf(local_pptx_path):
    if not os.path.exists(local_pptx_path):
        showerror("Erro", f"Arquivo '{local_pptx_path}' não foi encontrado para upload.")
        return

    service = get_gdrive_service()
    if not service:
        showerror("Erro Google Drive", "Não foi possível conectar ao Google Drive para gerar PDF.")
        return

    pdf_output_name = local_pptx_path.replace(".pptx", ".pdf")
    base_name = os.path.basename(local_pptx_path)

    try:
        # 1) Upload (convertendo para Google Slides)
        print(f"Iniciando upload de '{base_name}' para conversão...")
        file_metadata = {
            'name': base_name, # Mantém o nome original no Drive
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
        file_id = uploaded_file.get('id')
        if not file_id:
            raise Exception("Upload para o Google Drive falhou (não retornou ID).")
        print(f"Arquivo '{base_name}' enviado como Google Slides. ID: {file_id}")

        # 2) Exportar para PDF
        print(f"Exportando Google Slides (ID: {file_id}) para PDF...")
        request = service.files().export_media(fileId=file_id, mimeType='application/pdf')
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            if status:
                print(f"Progresso download PDF: {int(status.progress() * 100)}%")

        # 3) Salvar PDF localmente
        with open(pdf_output_name, 'wb') as f:
            f.write(fh.getvalue())
        print(f"PDF gerado localmente: '{pdf_output_name}'.")
        showinfo("Google Drive", f"PDF gerado com sucesso:\n'{pdf_output_name}'.")

        # 4) Opcional: Deletar o arquivo Google Slides temporário do Drive
        try:
            service.files().delete(fileId=file_id).execute()
            print(f"Arquivo Google Slides temporário (ID: {file_id}) deletado do Drive.")
        except Exception as delete_err:
            print(f"Aviso: Falha ao deletar arquivo temporário do Google Drive: {delete_err}")


    except Exception as e:
        showerror("Erro Google Drive", f"Ocorreu um erro durante o upload/conversão para PDF:\n{e}")


# ---------------------------------------------------------
# MainApp (SEM ALTERAÇÕES NA LÓGICA PRINCIPAL)
# Os botões chamam as funções gerar_proposta/gerar_material,
# que agora usam a lógica atualizada dentro da classe PlanoFrame.
# ---------------------------------------------------------
class MainApp(ttkb.Window):
    def __init__(self):
        super().__init__(themename="litera") # Tema padrão ttkbootstrap
        self.title("Gerador de Propostas ConnectPlug v2.0") # Título atualizado
        self.geometry("1200x800")

        # Variáveis do Vendedor
        self.nome_closer_var = tk.StringVar()
        self.celular_closer_var = tk.StringVar()
        self.email_closer_var = tk.StringVar()

        # Variáveis Compartilhadas para Cliente/Validade em TODAS as abas
        self.nome_cliente_var_shared = tk.StringVar(value="") # Começa vazio
        self.validade_proposta_var_shared = tk.StringVar(value=date.today().strftime("%d/%m/%Y")) # Data de hoje como padrão

        carregar_config(self.nome_closer_var, self.celular_closer_var, self.email_closer_var)
        self.protocol("WM_DELETE_WINDOW", self.on_close) # Salva config ao fechar

        # --- Top bar (Vendedor) ---
        top_bar = ttkb.Frame(self)
        top_bar.pack(side="top", fill="x", pady=5, padx=5)

        ttkb.Label(top_bar, text="Vendedor:").pack(side="left", padx=(0, 2))
        ttkb.Entry(top_bar, textvariable=self.nome_closer_var, width=20).pack(side="left", padx=(0, 5))
        ttkb.Label(top_bar, text="Celular:").pack(side="left", padx=(0, 2))
        ttkb.Entry(top_bar, textvariable=self.celular_closer_var, width=15).pack(side="left", padx=(0, 5))
        ttkb.Label(top_bar, text="Email:").pack(side="left", padx=(0, 2))
        ttkb.Entry(top_bar, textvariable=self.email_closer_var, width=25).pack(side="left", padx=(0, 10))

        self.btn_add = ttkb.Button(top_bar, text="+ Nova Aba", command=self.add_aba, bootstyle="success") # Estilo botão
        self.btn_add.pack(side="right", padx=5)

        # --- Notebook (Abas) ---
        self.notebook = ttkb.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=(0, 5))

        # --- Bottom Frame (Botões de Ação) ---
        bot_frame = ttkb.Frame(self)
        bot_frame.pack(side="bottom", fill="x", pady=5, padx=5)

        ttkb.Button(bot_frame, text="Gerar Proposta + PDF", command=self.on_gerar_proposta, bootstyle="primary").pack(side="left", padx=5)
        ttkb.Button(bot_frame, text="Gerar Material + PDF", command=self.on_gerar_mat_tecnico, bootstyle="info").pack(side="left", padx=5)
        ttkb.Button(bot_frame, text="Gerar TUDO + PDF", command=self.on_gerar_tudo, bootstyle="secondary").pack(side="left", padx=5)

        # --- Inicialização das Abas ---
        self.abas_criadas = {} # Dicionário para rastrear abas: {indice: widget_frame}
        self.ultimo_indice = 0
        self.add_aba() # Cria a primeira aba ao iniciar

        # --- Baixar Arquivos Base ---
        # Colocar em try/except para não impedir a abertura da UI se falhar
        try:
            baixar_arquivo_if_needed(
                "Proposta Comercial ConnectPlug.pptx",
                "https://github.com/DevRGS/Gerador/raw/refs/heads/main/assets/Proposta%20Comercial%20ConnectPlug.pptx"
            )
            baixar_arquivo_if_needed(
                "Material Tecnico ConnectPlug.pptx",
                "https://github.com/DevRGS/Gerador/raw/refs/heads/main/assets/Material%20Tecnico%20ConnectPlug.pptx"
            )
        except Exception as e:
             showerror("Erro no Download", f"Não foi possível baixar os arquivos de template:\n{e}")


    def on_close(self):
        """ Chamado ao fechar a janela principal. """
        salvar_config(
            self.nome_closer_var.get(),
            self.celular_closer_var.get(),
            self.email_closer_var.get()
        )
        self.destroy()

    def add_aba(self):
        """ Adiciona uma nova aba (PlanoFrame) ao Notebook. """
        if len(self.abas_criadas) >= MAX_ABAS:
            showinfo("Limite Atingido", f"Máximo de {MAX_ABAS} abas alcançado.")
            return

        self.ultimo_indice += 1
        idx = self.ultimo_indice

        # Cria a nova aba (frame)
        frame_aba = PlanoFrame(
            self.notebook,
            idx,
            nome_cliente_var_shared=self.nome_cliente_var_shared, # Passa a variável compartilhada
            validade_proposta_var_shared=self.validade_proposta_var_shared, # Passa a variável compartilhada
            on_close_callback=self.fechar_aba # Função para ser chamada ao clicar em "Fechar Aba"
        )

        # Adiciona a aba ao Notebook
        self.notebook.add(frame_aba, text=f"Plano {idx}")
        self.abas_criadas[idx] = frame_aba # Guarda a referência
        self.notebook.select(frame_aba) # Seleciona a aba recém-criada

        # Desabilita botão de adicionar se atingiu o limite
        if len(self.abas_criadas) >= MAX_ABAS:
            self.btn_add.config(state="disabled")


    def fechar_aba(self, indice):
        """ Remove uma aba do Notebook. """
        if indice in self.abas_criadas:
            frame_aba = self.abas_criadas[indice]
            try:
                self.notebook.forget(frame_aba) # Remove da interface
                del self.abas_criadas[indice] # Remove do rastreamento
                # Reabilita o botão de adicionar se estava no limite
                if len(self.abas_criadas) < MAX_ABAS:
                     self.btn_add.config(state="normal")
            except tk.TclError:
                # Pode acontecer se a aba já foi destruída por algum motivo
                pass
            # Se não houver mais abas, adicionar uma nova? Ou deixar vazio?
            if not self.abas_criadas:
                self.add_aba() # Adiciona uma nova aba se a última foi fechada


    def get_abas_ativas(self):
        """ Retorna uma lista dos widgets PlanoFrame das abas visíveis/ativas. """
        # Ordenar pelos índices para garantir consistência
        indices_ativos = sorted(self.abas_criadas.keys())
        return [self.abas_criadas[idx] for idx in indices_ativos]

    def on_gerar_proposta(self):
        abas_ativas = self.get_abas_ativas()
        if not abas_ativas:
            showerror("Erro", "Nenhuma aba ativa para gerar Proposta.")
            return

        # Validação básica dos dados do vendedor
        if not self.nome_closer_var.get() or not self.celular_closer_var.get() or not self.email_closer_var.get():
            showerror("Dados Incompletos", "Preencha os dados do Vendedor (Nome, Celular, Email).")
            return
        # Validação básica do nome do cliente
        if not self.nome_cliente_var_shared.get():
             showerror("Dados Incompletos", "Preencha o Nome do Cliente.")
             return

        pptx_file = gerar_proposta(
            abas_ativas, # Passa a lista de abas ativas
            self.nome_closer_var.get(),
            self.celular_closer_var.get(),
            self.email_closer_var.get()
        )
        if pptx_file and os.path.exists(pptx_file):
            upload_pptx_and_export_to_pdf(pptx_file)

    def on_gerar_mat_tecnico(self):
        abas_ativas = self.get_abas_ativas()
        if not abas_ativas:
            showerror("Erro", "Nenhuma aba ativa para gerar Material Técnico.")
            return

        # Validações (opcional, mas recomendado)
        if not self.nome_cliente_var_shared.get():
             showerror("Dados Incompletos", "Preencha o Nome do Cliente.")
             return

        pptx_file = gerar_material(
            abas_ativas, # Passa a lista de abas ativas
            self.nome_closer_var.get(),
            self.celular_closer_var.get(),
            self.email_closer_var.get()
        )
        if pptx_file and os.path.exists(pptx_file):
            upload_pptx_and_export_to_pdf(pptx_file)

    def on_gerar_tudo(self):
        abas_ativas = self.get_abas_ativas()
        if not abas_ativas:
            showerror("Erro", "Nenhuma aba ativa para gerar.")
            return

        # Validações
        if not self.nome_closer_var.get() or not self.celular_closer_var.get() or not self.email_closer_var.get():
            showerror("Dados Incompletos", "Preencha os dados do Vendedor (Nome, Celular, Email).")
            return
        if not self.nome_cliente_var_shared.get():
             showerror("Dados Incompletos", "Preencha o Nome do Cliente.")
             return

        # 1) Gera Proposta
        print("--- Gerando Proposta ---")
        pptx_prop = gerar_proposta(
            abas_ativas,
            self.nome_closer_var.get(),
            self.celular_closer_var.get(),
            self.email_closer_var.get()
        )
        pdf_prop_ok = False
        if pptx_prop and os.path.exists(pptx_prop):
            try:
                upload_pptx_and_export_to_pdf(pptx_prop)
                pdf_prop_ok = True
            except Exception as e:
                 print(f"Erro ao gerar PDF da Proposta: {e}")
                 # Não impede a geração do material

        # 2) Gera Material
        print("--- Gerando Material Técnico ---")
        pptx_mat = gerar_material(
            abas_ativas,
            self.nome_closer_var.get(),
            self.celular_closer_var.get(),
            self.email_closer_var.get()
        )
        pdf_mat_ok = False
        if pptx_mat and os.path.exists(pptx_mat):
            try:
                upload_pptx_and_export_to_pdf(pptx_mat)
                pdf_mat_ok = True
            except Exception as e:
                 print(f"Erro ao gerar PDF do Material: {e}")

        # Mensagem final (opcional)
        if pdf_prop_ok and pdf_mat_ok:
             print("--- Geração Concluída (Proposta e Material) ---")
        else:
             print("--- Geração Concluída com possíveis erros na conversão para PDF ---")


# --- Função Principal ---
def main():
    # ttkb.Style(theme='litera') # Definir tema globalmente (opcional)
    app = MainApp()
    app.mainloop()

if __name__ == "__main__":
    main()