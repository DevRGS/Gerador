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
            pass # Ignora arquivo corrompido

def salvar_config(nome_closer, celular_closer, email_closer):
    dados = {
        "nome_vendedor": nome_closer,
        "celular_vendedor": celular_closer,
        "email_vendedor": email_closer
    }
    # Tenta tornar o arquivo gravável se existir (útil em alguns sistemas)
    if os.path.exists(CONFIG_FILE):
        try:
            os.chmod(CONFIG_FILE, 0o666)
        except OSError:
            pass # Ignora erro se não puder mudar permissão
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(dados, f, indent=4, ensure_ascii=False)
    except PermissionError:
        print(f"Aviso: Sem permissão para salvar {CONFIG_FILE}")
    except Exception as e:
        print(f"Erro ao salvar config: {e}")


# ---------------------------------------------------------
# Dados de Planos e Tabelas de Preço - Base Mensal = Custo Efetivo no Anual
# ---------------------------------------------------------
LISTA_PLANOS_UI = ["PDV Básico", "Gestão", "Performance", "Autoatendimento", "Bling", "Em Branco"]
LISTA_PLANOS_BLING = ["Bling - Básico", "Bling - Com Estoque em Grade"]

PLAN_INFO = {
    "PDV Básico": {
        # base_mensal AGORA significa -> Custo Mensal Efetivo no contrato Anual
        "base_mensal": 99.00,
        "min_pdv": 1, "min_users": 2,
        "max_extra_users": 1, "max_extra_pdvs": 0,
        "mandatory": ["Usuários", "30 Notas Fiscais", "Suporte Técnico - Via chamados", "Relatório Básico", "PDV - Frente de Caixa"],
        "allowed_optionals": ["Smart Menu", "Terminais Autoatendimento", "Hub de Delivery", "Delivery Direto Profissional", "Delivery Direto VIP", "TEF", "Importação de XML", "Cardápio digital"]
    },
    "Gestão": {
        "base_mensal": 199.00,
        "min_pdv": 2, "min_users": 3,
        "max_extra_users": 2, "max_extra_pdvs": 1,
        "mandatory": ["Notas Fiscais Ilimitadas", "Importação de XML", "PDV - Frente de Caixa", "Usuários", "Painel Senha TV", "Estoque em Grade", "Relatórios", "Suporte Técnico - Via chamados", "Suporte Técnico - Via chat", "Delivery", "Relatório KDS"],
        "allowed_optionals": ["Facilita NFE", "Conciliação Bancária", "Contratos de cartões e outros", "Delivery Direto Profissional", "Delivery Direto VIP", "TEF", "Integração API", "Business Intelligence (BI)", "Backup Realtime", "Cardápio digital", "Smart Menu", "Hub de Delivery", "Ordem de Serviço", "App Gestão CPlug", "Painel Senha Mobile", "Controle de Mesas", "Produção", "Promoções", "Marketing", "Relatório Dinâmico", "Atualização em tempo real", "Smart TEF", "Terminais Autoatendimento", "Suporte Técnico - Estendido"]
    },
    "Performance": {
        "base_mensal": 499.00,
        "min_pdv": 3, "min_users": 5,
        "max_extra_users": 5, "max_extra_pdvs": 2,
        "mandatory": ["Produção", "Promoções", "Notas Fiscais Ilimitadas", "Importação de XML", "Hub de Delivery", "Ordem de Serviço", "Delivery", "App Gestão CPlug", "Relatório KDS", "Painel Senha TV", "Painel Senha Mobile", "Controle de Mesas", "Estoque em Grade", "Marketing", "Relatórios", "Relatório Dinâmico", "Atualização em tempo real", "Facilita NFE", "Conciliação Bancária", "Contratos de cartões e outros", "Suporte Técnico - Via chamados", "Suporte Técnico - Via chat", "Suporte Técnico - Estendido", "PDV - Frente de Caixa", "Smart TEF", "Usuários"],
        "allowed_optionals": ["TEF", "Programa de Fidelidade", "Integração Tap", "Integração API", "Business Intelligence (BI)", "Backup Realtime", "Cardápio digital", "Smart Menu", "Terminais Autoatendimento", "Delivery Direto Profissional", "Delivery Direto VIP"]
    },
    "Autoatendimento": {
        "base_mensal": 419.90, # Assumindo que este é o Custo Mensal Efetivo no anual para 1 terminal
        "min_pdv": 0, "min_users": 1,
        "max_extra_users": 998, "max_extra_pdvs": 99,
        "mandatory": ["Contratos de cartões e outros", "Estoque em Grade", "Notas Fiscais Ilimitadas", "Produção"], # Vendas.. removido
        "allowed_optionals": [] # Definir se houver
    },
    "Bling - Básico": {
        "base_mensal": 189.90, # Este JÁ É o valor mensal efetivo no anual
        "min_pdv": 1, "min_users": 5,
        "max_extra_users": 994, "max_extra_pdvs": 98,
        "mandatory": ["Relatórios", "Notas Fiscais Ilimitadas"], # Vendas.. removido
        "allowed_optionals": [], # Definir se houver
        # "base_mensal_original": 369.80, # Pode ser útil para exibir "De R$..."
    },
    "Bling - Com Estoque em Grade": {
        "base_mensal": 219.90, # Este JÁ É o valor mensal efetivo no anual
        "min_pdv": 1, "min_users": 5,
        "max_extra_users": 994, "max_extra_pdvs": 98,
        "mandatory": ["Relatórios", "Notas Fiscais Ilimitadas", "Estoque em Grade"], # Vendas.. removido
        "allowed_optionals": [], # Definir se houver
        # "base_mensal_original": 399.80, # Pode ser útil para exibir "De R$..."
    },
    "Em Branco": {
        "base_mensal": 0.0,
        "min_pdv": 0, "min_users": 0,
        "max_extra_users": 999, "max_extra_pdvs": 99,
        "mandatory": [],
        "allowed_optionals": [] # Definir se houver (provavelmente todos?)
    }
}

# Módulos SEM DESCONTO (se a lógica de desconto manual ainda for relevante)
# Itens cujo preço mensal não deve ser afetado pelo % de desconto editado
SEM_DESCONTO = {
    "TEF", "Terminais Autoatendimento", "Smart TEF",
    "Delivery Direto Profissional", "Delivery Direto VIP",
    "Programa de Fidelidade", "Integração Tap", "Integração API",
    "Business Intelligence (BI)", "Backup Realtime",
    # Manter os antigos se ainda relevantes?
    # "Domínio Próprio", ... etc
}

# Dicionário de preços mensais dos EXTRAS (mantido da versão anterior)
precos_mensais = {
    # Módulos que são *sempre* fixos não precisam de preço aqui
    # Preços para módulos que podem ser OPCIONAIS em alguns planos
    "Importação de XML": 29.00,
    "Produção": 30.00,
    "Promoções": 24.50,
    "Hub de Delivery": 79.00,
    "Ordem de Serviço": 20.00,
    "App Gestão CPlug": 20.00,
    "Painel Senha Mobile": 49.00,
    "Controle de Mesas": 49.00,
    "Marketing": 24.50,
    "Relatório Dinâmico": 50.00,
    "Atualização em tempo real": 49.00,
    "Facilita NFE": 99.00,
    "Conciliação Bancária": 50.00,
    "Contratos de cartões e outros": 50.00,
    "Suporte Técnico - Estendido": 99.00,
    "Smart TEF": 49.90, # Usado para cálculo de extras no Gestão

    # Opcionais puros
    "Smart Menu": 99.90,
    "Terminais Autoatendimento": 199.00,
    "Delivery Direto Profissional": 200.00,
    "Delivery Direto VIP": 300.00,
    "TEF": 99.90,
    "Cardápio digital": 99.00,
    "Backup Realtime": 199.90, # Usar o maior valor, ou ter preços por plano?
    "Business Intelligence (BI)": 199.00, # Usar o maior valor?
    "Programa de Fidelidade": 299.90,
    "Integração Tap": 299.00,
    "Integração API": 299.00,

    # Itens obsoletos (manter comentado ou remover)
    # "60 Notas Fiscais": 40.00, ... etc
}

# Custos Adicionais (fixos por item extra)
PRECO_EXTRA_USUARIO = 19.00
PRECO_EXTRA_PDV_GESTAO_PERFORMANCE = 59.90
# PRECO_EXTRA_PDV_BLING = 40.00 # Se Bling ainda permitir extra PDV

# ---------------------------------------------------------
# Função utilitária para substituir placeholders no Slide
# ---------------------------------------------------------
def substituir_placeholders_no_slide(slide, dados):
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                # Iterar em cópia da lista de runs para poder modificar
                runs_copy = list(paragraph.runs)
                paragraph.clear() # Limpa o parágrafo para reconstruir
                new_run = None
                for run in runs_copy:
                    txt = run.text
                    # Aplica substituições
                    for k, v in dados.items():
                        placeholder = f"{{{k}}}" # Assume placeholders como {chave}
                        if placeholder in txt:
                             # Garante que v seja string
                             v_str = str(v) if v is not None else ""
                             txt = txt.replace(placeholder, v_str)

                    # Adiciona o texto modificado ao parágrafo, preservando a formatação do run original
                    # (Nota: isso pode não preservar formatações complexas que abrangem múltiplos runs)
                    if new_run is None:
                        new_run = paragraph.add_run()
                    new_run.text += txt # Adiciona ao run atual
                    # Copia formatação básica (negrito, itálico, tamanho, etc.)
                    # Isso é simplificado, pode precisar de mais detalhes para fontes, cores, etc.
                    new_run.font.bold = run.font.bold
                    new_run.font.italic = run.font.italic
                    new_run.font.size = run.font.size
                    # Se o placeholder foi o último item no run, começa um novo run para o próximo texto
                    # Isso ajuda a evitar que formatações se misturem indevidamente
                    if placeholder in run.text and run.text.endswith(placeholder):
                        new_run = None



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
        self.nome_plano_var = tk.StringVar(value="")

        self.current_plan = "PDV Básico"
        self.spin_pdv_var = tk.IntVar(value=1)
        self.spin_users_var = tk.IntVar(value=1)
        self.spin_smart_tef_var = tk.IntVar(value=0)
        # Remover outros spinboxes antigos se não forem mais usados

        # Módulos (checkboxes)
        self.modules = {
            # Fixos (necessários para lógica, mas controlados por código)
            "Usuários": tk.IntVar(), "30 Notas Fiscais": tk.IntVar(), "Suporte Técnico - Via chamados": tk.IntVar(),
            "Relatório Básico": tk.IntVar(), "PDV - Frente de Caixa": tk.IntVar(), "Notas Fiscais Ilimitadas": tk.IntVar(),
            "Importação de XML": tk.IntVar(), "Painel Senha TV": tk.IntVar(), "Estoque em Grade": tk.IntVar(),
            "Relatórios": tk.IntVar(), "Suporte Técnico - Via chat": tk.IntVar(), "Delivery": tk.IntVar(),
            "Relatório KDS": tk.IntVar(), "Produção": tk.IntVar(), "Promoções": tk.IntVar(), "Hub de Delivery": tk.IntVar(),
            "Ordem de Serviço": tk.IntVar(), "App Gestão CPlug": tk.IntVar(), "Painel Senha Mobile": tk.IntVar(),
            "Controle de Mesas": tk.IntVar(), "Marketing": tk.IntVar(), "Relatório Dinâmico": tk.IntVar(),
            "Atualização em tempo real": tk.IntVar(), "Facilita NFE": tk.IntVar(), "Conciliação Bancária": tk.IntVar(),
            "Contratos de cartões e outros": tk.IntVar(), "Suporte Técnico - Estendido": tk.IntVar(), "Smart TEF": tk.IntVar(),
            # Opcionais (reais checkboxes)
            "Smart Menu": tk.IntVar(), "Terminais Autoatendimento": tk.IntVar(), "Delivery Direto Profissional": tk.IntVar(),
            "Delivery Direto VIP": tk.IntVar(), "TEF": tk.IntVar(), "Cardápio digital": tk.IntVar(),
            "Integração API": tk.IntVar(), "Business Intelligence (BI)": tk.IntVar(), "Backup Realtime": tk.IntVar(),
            "Programa de Fidelidade": tk.IntVar(), "Integração Tap": tk.IntVar(),
        }
        self.check_buttons = {}

        # Overrides de cálculo
        self.user_override_anual_active = tk.BooleanVar(value=False)
        self.user_override_discount_active = tk.BooleanVar(value=False)
        self.valor_anual_editavel = tk.StringVar(value="0.00") # Representa o TOTAL ANUAL PAGO ADIANTADO
        self.desconto_personalizado = tk.StringVar(value="10") # Começa com 10% padrão

        # Armazenar valores calculados
        self.computed_mensal_sem_fidelidade = 0.0
        self.computed_mensal_efetivo_anual = 0.0
        self.computed_anual_total = 0.0
        self.computed_desconto_percent = 0.0
        self.computed_custo_adicional = 0.0

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

        self._montar_layout_esquerda()
        self._montar_layout_direita()
        self.configurar_plano("PDV Básico")

    def fechar_aba(self):
        if self.on_close_callback:
            self.on_close_callback(self.aba_index)

    def on_bling_selected(self, event=None):
        selected_bling_plan = self.bling_var.get()
        if selected_bling_plan in LISTA_PLANOS_BLING:
            self.configurar_plano(selected_bling_plan)
        self.bling_var.set("Selecionar Bling...")

    def _montar_layout_esquerda(self):
        # Top Bar e Seleção de Planos (sem mudanças significativas)
        top_bar = ttkb.Frame(self.frame_left)
        top_bar.pack(fill="x", pady=5)
        ttkb.Label(top_bar, text=f"Aba Plano {self.aba_index}", font="-size 12 -weight bold").pack(side="left")
        btn_close = ttkb.Button(top_bar, text="X", command=self.fechar_aba, bootstyle="danger-outline", width=3)
        btn_close.pack(side="right")

        frame_planos = ttkb.Labelframe(self.frame_left, text="Selecionar Plano Base")
        frame_planos.pack(fill="x", pady=5)
        self.bling_combobox = None
        plan_buttons_frame = ttkb.Frame(frame_planos) # Frame para botões
        plan_buttons_frame.pack(fill="x")
        for i, p in enumerate(LISTA_PLANOS_UI):
            if p == "Bling":
                self.bling_var = tk.StringVar(value="Selecionar Bling...")
                self.bling_combobox = ttk.Combobox(plan_buttons_frame, textvariable=self.bling_var,
                                                   values=LISTA_PLANOS_BLING, state="readonly", width=25)
                self.bling_combobox.grid(row=0, column=i, padx=5, pady=5, sticky="ew")
                self.bling_combobox.bind("<<ComboboxSelected>>", self.on_bling_selected)
            else:
                btn = ttkb.Button(plan_buttons_frame, text=p, width=15,
                            command=lambda pl=p: self.configurar_plano(pl))
                btn.grid(row=0, column=i, padx=5, pady=5, sticky="ew") # Usa grid para melhor alinhamento
            plan_buttons_frame.grid_columnconfigure(i, weight=1) # Faz colunas expansíveis


        # Módulos Opcionais (Checkboxes)
        frame_mod = ttkb.Labelframe(self.frame_left, text="Módulos Opcionais")
        frame_mod.pack(fill="both", expand=True, pady=5)
        f_mod_cols = ttkb.Frame(frame_mod)
        f_mod_cols.pack(fill="both", expand=True)

        f_mod_left = ttkb.Frame(f_mod_cols)
        f_mod_left.pack(side="left", fill="both", expand=True, padx=5)
        f_mod_right = ttkb.Frame(f_mod_cols)
        f_mod_right.pack(side="left", fill="both", expand=True, padx=5)

        # Filtrar módulos para exibir apenas os que podem ser opcionais
        displayable_mods = sorted([
            m for m, var in self.modules.items() if m not in PLAN_INFO["PDV Básico"]["mandatory"] or
                                                    m not in PLAN_INFO["Gestão"]["mandatory"] or
                                                    m not in PLAN_INFO["Performance"]["mandatory"] or
                                                    m in PLAN_INFO["PDV Básico"]["allowed_optionals"] or # Garante que opcionais apareçam
                                                    m in PLAN_INFO["Gestão"]["allowed_optionals"] or
                                                    m in PLAN_INFO["Performance"]["allowed_optionals"]
        ])
        # Remover itens controlados por spinbox da lista de checkboxes
        spinbox_items = {"PDV - Frente de Caixa", "Usuários", "Smart TEF"}
        displayable_mods = [m for m in displayable_mods if m not in spinbox_items]
        # Remover itens sempre fixos (ex: Relatórios, Notas Fiscais)
        always_fixed = {"Relatórios", "30 Notas Fiscais", "Notas Fiscais Ilimitadas", "Suporte Técnico - Via chamados", "Relatório Básico"}
        displayable_mods = [m for m in displayable_mods if m not in always_fixed]


        mid = len(displayable_mods)//2
        left_side = displayable_mods[:mid]
        right_side = displayable_mods[mid:]
        self.check_buttons = {}

        for m in left_side:
             if m in self.modules:
                 cb = ttk.Checkbutton(f_mod_left, text=m, variable=self.modules[m], command=self.atualizar_valores)
                 cb.pack(anchor="w", pady=2)
                 self.check_buttons[m] = cb

        for m in right_side:
             if m in self.modules:
                 cb = ttk.Checkbutton(f_mod_right, text=m, variable=self.modules[m], command=self.atualizar_valores)
                 cb.pack(anchor="w", pady=2)
                 self.check_buttons[m] = cb

        # Frame Dados Cliente e Plano
        frame_dados = ttkb.Labelframe(self.frame_left, text="Dados do Cliente e Proposta")
        frame_dados.pack(fill="x", pady=5)
        ttkb.Label(frame_dados, text="Nome do Cliente:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttkb.Entry(frame_dados, textvariable=self.nome_cliente_var, width=30).grid(row=0, column=1, padx=5, pady=2)
        ttkb.Label(frame_dados, text="Validade Proposta:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttkb.Entry(frame_dados, textvariable=self.validade_proposta_var, width=15).grid(row=1, column=1, padx=5, pady=2, sticky="w")
        ttkb.Label(frame_dados, text="Nome do Plano (Opcional):").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        ttkb.Entry(frame_dados, textvariable=self.nome_plano_var, width=30).grid(row=2, column=1, padx=5, pady=2)


    def _montar_layout_direita(self):
        # Quantidades (PDV, Usuários, Smart TEF)
        frame_inc = ttkb.Labelframe(self.frame_right, text="Quantidades")
        frame_inc.pack(fill="x", pady=5)

        ttkb.Label(frame_inc, text="PDVs - Frente de Caixa").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.sp_pdv = ttkb.Spinbox(frame_inc, from_=0, to=99, textvariable=self.spin_pdv_var, width=5, command=self.atualizar_valores)
        self.sp_pdv.grid(row=0, column=1, padx=5, pady=2)

        ttkb.Label(frame_inc, text="Usuários").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.sp_usr = ttkb.Spinbox(frame_inc, from_=0, to=999, textvariable=self.spin_users_var, width=5, command=self.atualizar_valores)
        self.sp_usr.grid(row=1, column=1, padx=5, pady=2)

        ttkb.Label(frame_inc, text="Smart TEF").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.sp_smf = ttkb.Spinbox(frame_inc, from_=0, to=99, textvariable=self.spin_smart_tef_var, width=5, command=self.atualizar_valores)
        self.sp_smf.grid(row=2, column=1, padx=5, pady=2)

        # --- Frame Valores Finais (Layout Atualizado) ---
        frame_valores = ttkb.Labelframe(self.frame_right, text="Valores da Proposta")
        frame_valores.pack(fill="both", pady=5, expand=True) # Expandir para preencher espaço

        # 1. Mensal (Sem Fidelidade)
        self.lbl_plano_mensal_sem_fid = ttkb.Label(frame_valores, text="Mensal (Sem Fidelidade): R$ 0,00", font="-size 11")
        self.lbl_plano_mensal_sem_fid.pack(pady=(5, 2), anchor="w", padx=5)

        # 2. Custo Treinamento (Associado ao mensal sem fidelidade)
        self.lbl_treinamento = ttkb.Label(frame_valores, text="+ Custo Treinamento: R$ 0,00", font="-size 9")
        self.lbl_treinamento.pack(pady=(0, 5), anchor="w", padx=15) # Indentado

        # 3. Mensal (No Plano Anual)
        self.lbl_plano_mensal_no_anual = ttkb.Label(frame_valores, text="Mensal (no Plano Anual): R$ 0,00", font="-size 12 -weight bold")
        self.lbl_plano_mensal_no_anual.pack(pady=5, anchor="w", padx=5)

        # 4. Anual (Pagamento Único)
        self.lbl_plano_anual_total = ttkb.Label(frame_valores, text="Anual (Pagamento Único): R$ 0,00", font="-size 12 -weight bold")
        self.lbl_plano_anual_total.pack(pady=5, anchor="w", padx=5)

        # 5. Desconto Aplicado (Informativo)
        self.lbl_desconto = ttkb.Label(frame_valores, text="Desconto Anual Aplicado: 10%", font="-size 9")
        self.lbl_desconto.pack(pady=(5, 10), anchor="w", padx=5)

        # Separador
        ttk.Separator(frame_valores, orient='horizontal').pack(fill='x', pady=5, padx=5)

        # --- Frames de Edição Manual ---
        frame_edicao = ttkb.Frame(frame_valores)
        frame_edicao.pack(fill="x", pady=5)

        # 6. Edição Anual Total
        frame_edit_anual = ttkb.Labelframe(frame_edicao, text="Editar Anual Total (R$)")
        frame_edit_anual.pack(side="left", padx=5, fill="x", expand=True)
        e_anual = ttkb.Entry(frame_edit_anual, textvariable=self.valor_anual_editavel, width=10)
        e_anual.pack(side="left", padx=5, pady=2)
        e_anual.bind("<KeyRelease>", self.on_user_edit_valor_anual)
        b_reset_anual = ttkb.Button(frame_edit_anual, text="Reset", command=self.on_reset_anual, width=5, bootstyle="warning-outline")
        b_reset_anual.pack(side="left", padx=5, pady=2)

        # 7. Edição Desconto %
        frame_edit_desc = ttkb.Labelframe(frame_edicao, text="Editar Desconto (%)")
        frame_edit_desc.pack(side="left", padx=5, fill="x", expand=True)
        e_desc = ttkb.Entry(frame_edit_desc, textvariable=self.desconto_personalizado, width=5)
        e_desc.pack(side="left", padx=5, pady=2)
        e_desc.bind("<KeyRelease>", self.on_user_edit_desconto)
        b_reset_desc = ttkb.Button(frame_edit_desc, text="Reset", command=self.on_reset_desconto, width=5, bootstyle="warning-outline")
        b_reset_desc.pack(side="left", padx=5, pady=2)


    # --- Funções de Edição/Reset (Mantidas, mas disparam recálculo) ---
    def on_user_edit_valor_anual(self, *args):
        # Usuário editou o VALOR TOTAL ANUAL
        self.user_override_anual_active.set(True)
        self.user_override_discount_active.set(False) # Desativa override de desconto
        # Força a atualização do campo de desconto para refletir o valor digitado
        self.atualizar_valores() # Recalcula tudo, inclusive o desconto implícito
        # Tenta ler o desconto calculado e atualizar o campo de entrada de desconto
        try:
            desc_calc = self.computed_desconto_percent
            self.desconto_personalizado.set(str(round(desc_calc)))
        except: pass # Ignora erro se o cálculo falhar


    def on_reset_anual(self):
        self.user_override_anual_active.set(False)
        # Não reseta valor_anual_editavel aqui, deixa atualizar_valores recalcular
        self.atualizar_valores()

    def on_user_edit_desconto(self, *args):
        # Usuário editou o PERCENTUAL DE DESCONTO
        self.user_override_discount_active.set(True)
        self.user_override_anual_active.set(False) # Desativa override de valor anual
        # Força a atualização do campo de valor anual para refletir o desconto digitado
        self.atualizar_valores() # Recalcula tudo, inclusive o valor anual total
        # Tenta ler o valor anual calculado e atualizar o campo de entrada anual
        try:
            anual_calc = self.computed_anual_total
            self.valor_anual_editavel.set(f"{anual_calc:.2f}")
        except: pass


    def on_reset_desconto(self):
        self.user_override_discount_active.set(False)
        # Não reseta desconto_personalizado aqui, deixa atualizar_valores recalcular (para 10%)
        self.atualizar_valores()

    def configurar_plano(self, plano):
        # Reset Bling Combobox
        if not plano.startswith("Bling -") and self.bling_combobox:
             self.bling_var.set("Selecionar Bling...")

        if plano not in PLAN_INFO:
             showerror("Erro de Configuração", f"Plano '{plano}' não encontrado.")
             return

        info = PLAN_INFO[plano]
        self.current_plan = plano

        # Configurar Spinboxes (Mínimos e Máximos)
        min_pdv = info.get("min_pdv", 0); max_pdv = min_pdv + info.get("max_extra_pdvs", 99)
        min_users = info.get("min_users", 0); max_users = min_users + info.get("max_extra_users", 999)

        self.spin_pdv_var.set(min_pdv)
        self.sp_pdv.config(from_=min_pdv, to=max_pdv)
        self.spin_users_var.set(min_users)
        self.sp_usr.config(from_=min_users, to=max_users)

        # Limite Smart TEF no Gestão
        min_smart_tef = 0; max_smart_tef = 99 # Padrão
        if plano == "Gestão":
             max_smart_tef = 3
        elif plano == "Performance":
             min_smart_tef = 3 # Performance inclui 3
             max_smart_tef = 3 # Não permite adicionar extras pelo spinbox

        self.spin_smart_tef_var.set(min_smart_tef)
        self.sp_smf.config(from_=min_smart_tef, to=max_smart_tef)

        # Resetar/Configurar Módulos (Checkboxes)
        for m, var in self.modules.items(): var.set(0) # Limpa todos
        allowed = info.get("allowed_optionals", [])
        mandatory = info.get("mandatory", [])
        for m in mandatory: # Marca os mandatórios
            if m in self.modules: self.modules[m].set(1)

        # Habilita/Desabilita Checkboxes
        for m, cb in self.check_buttons.items():
             is_mandatory = m in mandatory
             is_allowed_optional = m in allowed
             if is_mandatory:
                 cb.config(state='disabled')
             elif is_allowed_optional:
                 cb.config(state='normal')
             else: # Nem mandatório, nem opcional permitido
                 cb.config(state='disabled')

        # Resetar overrides e recalcular
        self.user_override_anual_active.set(False)
        self.user_override_discount_active.set(False)
        # Definir o desconto padrão como 10 ao configurar o plano
        self.desconto_personalizado.set("10")
        self.atualizar_valores() # Dispara o recálculo inicial

    def _calcular_extras(self):
        """Calcula o custo MENSAL total dos extras e separa por descontável/não descontável."""
        total_extras_cost = 0.0
        total_extras_descontavel = 0.0
        total_extras_nao_descontavel = 0.0
        info = PLAN_INFO[self.current_plan]
        mandatory = info.get("mandatory", [])

        # 1. PDVs Extras
        pdv_atuais = self.spin_pdv_var.get(); pdv_incluidos = info.get("min_pdv", 0)
        pdv_extras = max(0, pdv_atuais - pdv_incluidos)
        pdv_price = 0.0
        if self.current_plan in ["Gestão", "Performance"]: pdv_price = PRECO_EXTRA_PDV_GESTAO_PERFORMANCE
        elif self.current_plan.startswith("Bling"): pdv_price = 40.00 # PRECO_EXTRA_PDV_BLING?
        cost_pdv_extra = pdv_extras * pdv_price
        total_extras_cost += cost_pdv_extra
        if "PDV Extra" not in SEM_DESCONTO: total_extras_descontavel += cost_pdv_extra
        else: total_extras_nao_descontavel += cost_pdv_extra

        # 2. Users Extras
        users_atuais = self.spin_users_var.get(); users_incluidos = info.get("min_users", 0)
        users_extras = max(0, users_atuais - users_incluidos)
        cost_users_extra = users_extras * PRECO_EXTRA_USUARIO
        total_extras_cost += cost_users_extra
        if "User Extra" not in SEM_DESCONTO: total_extras_descontavel += cost_users_extra
        else: total_extras_nao_descontavel += cost_users_extra

        # 3. Módulos Extras (Checkboxes Opcionais Selecionados)
        for m, var_m in self.modules.items():
            if m in self.check_buttons and var_m.get() == 1 and m not in mandatory:
                 price = precos_mensais.get(m, 0.0)
                 total_extras_cost += price
                 if m not in SEM_DESCONTO: total_extras_descontavel += price
                 else: total_extras_nao_descontavel += price

        # 4. Spinboxes Extras (Apenas Smart TEF no Gestão)
        if self.current_plan == "Gestão":
            smart_tef_atuais = self.spin_smart_tef_var.get()
            smart_tef_incluidos = 0 # Gestão não inclui nenhum fixo
            smart_tef_extras = max(0, smart_tef_atuais - smart_tef_incluidos)
            if smart_tef_extras > 0:
                 price = smart_tef_extras * precos_mensais.get("Smart TEF", 0.0)
                 total_extras_cost += price
                 if "Smart TEF" not in SEM_DESCONTO: total_extras_descontavel += price
                 else: total_extras_nao_descontavel += price

        return total_extras_cost, total_extras_descontavel, total_extras_nao_descontavel

    def atualizar_valores(self, *args):
        if not self.current_plan or self.current_plan not in PLAN_INFO: return
        info = PLAN_INFO[self.current_plan]
        is_bling_plan = self.current_plan.startswith("Bling -")
        is_autoatendimento_plano = self.current_plan == "Autoatendimento"
        is_em_branco_plano = self.current_plan == "Em Branco"

        # --- Obter Custos Base e Extras ---
        # base_mensal aqui é o CUSTO MENSAL EFETIVO NO PLANO ANUAL
        base_mensal_efetivo_anual = info.get("base_mensal", 0.0)
        total_extras_cost, total_extras_descontavel, total_extras_nao_descontavel = self._calcular_extras()

        # --- Calcular Valores Principais ---
        # 1. Base Mensal Sem Fidelidade (Base + 10%)
        #    Use divisão por 0.9 para ser o inverso exato do desconto de 10%
        base_mensal_sem_fidelidade = (base_mensal_efetivo_anual / 0.90) if base_mensal_efetivo_anual > 0 else 0.0

        # 2. Total Mensal Sem Fidelidade (Base Sem Fid + Extras)
        total_mensal_sem_fidelidade = base_mensal_sem_fidelidade + total_extras_cost

        # 3. Total Mensal Efetivo Anual (Base Efetiva Anual + Extras) - Valor base para cálculo com overrides
        total_mensal_efetivo_anual_base_calc = base_mensal_efetivo_anual + total_extras_cost

        # --- Aplicar Overrides ---
        final_mensal_efetivo_anual = 0.0
        final_anual_total = 0.0
        desconto_aplicado_percent = 10.0 # Padrão implícito

        if self.user_override_anual_active.get():
            # Usuário digitou o TOTAL ANUAL PAGO ADIANTADO
            try:
                edited_total_anual = float(self.valor_anual_editavel.get())
                final_anual_total = edited_total_anual
                final_mensal_efetivo_anual = final_anual_total / 12.0
                # Recalcular desconto implícito para exibição
                if total_mensal_sem_fidelidade > 0:
                    desconto_aplicado_percent = ((total_mensal_sem_fidelidade - final_mensal_efetivo_anual) / total_mensal_sem_fidelidade) * 100
                else:
                    desconto_aplicado_percent = 0.0
                # Atualiza o campo de desconto para refletir
                self.desconto_personalizado.set(str(round(max(0, desconto_aplicado_percent))))

            except ValueError:
                # Se valor inválido, recalcula como padrão
                final_mensal_efetivo_anual = total_mensal_efetivo_anual_base_calc
                final_anual_total = final_mensal_efetivo_anual * 12.0
                if total_mensal_sem_fidelidade > 0:
                    desconto_aplicado_percent = ((total_mensal_sem_fidelidade - final_mensal_efetivo_anual) / total_mensal_sem_fidelidade) * 100
                else:
                    desconto_aplicado_percent = 0.0
                self.desconto_personalizado.set("10") # Volta pro padrão visual
                self.valor_anual_editavel.set(f"{final_anual_total:.2f}") # Corrige valor anual

        elif self.user_override_discount_active.get():
            # Usuário digitou o PERCENTUAL DE DESCONTO
            try:
                desc_custom = float(self.desconto_personalizado.get())
                desc_dec = desc_custom / 100.0
                desconto_aplicado_percent = desc_custom

                # Aplica desconto sobre a Base Sem Fidelidade + Extras Descontáveis
                base_sem_fid_mais_extras_descont = base_mensal_sem_fidelidade + total_extras_descontavel
                # Calcula o valor mensal efetivo com o desconto customizado
                final_mensal_efetivo_anual = (base_sem_fid_mais_extras_descont * (1 - desc_dec)) + total_extras_nao_descontavel
                final_anual_total = final_mensal_efetivo_anual * 12.0
                # Atualiza o campo anual para refletir
                self.valor_anual_editavel.set(f"{final_anual_total:.2f}")

            except ValueError:
                # Se desconto inválido, recalcula como padrão
                final_mensal_efetivo_anual = total_mensal_efetivo_anual_base_calc
                final_anual_total = final_mensal_efetivo_anual * 12.0
                if total_mensal_sem_fidelidade > 0:
                     desconto_aplicado_percent = ((total_mensal_sem_fidelidade - final_mensal_efetivo_anual) / total_mensal_sem_fidelidade) * 100
                else: desconto_aplicado_percent = 0.0
                self.desconto_personalizado.set("10") # Volta pro padrão visual
                self.valor_anual_editavel.set(f"{final_anual_total:.2f}")

        else:
            # Cálculo Padrão (sem overrides ativos)
            final_mensal_efetivo_anual = total_mensal_efetivo_anual_base_calc
            final_anual_total = final_mensal_efetivo_anual * 12.0
            # Calcula o desconto implícito (deve ser ~10% se só a base mudou)
            if total_mensal_sem_fidelidade > 0:
                desconto_aplicado_percent = ((total_mensal_sem_fidelidade - final_mensal_efetivo_anual) / total_mensal_sem_fidelidade) * 100
            else:
                 desconto_aplicado_percent = 0.0
            # Atualiza campos editáveis com valores padrão calculados
            self.valor_anual_editavel.set(f"{final_anual_total:.2f}")
            self.desconto_personalizado.set(str(round(max(0,desconto_aplicado_percent)))) # Garante que o 10% padrão apareça


        # --- Calcular Custo Adicional (Treinamento) ---
        # Baseado no Total Mensal SEM Fidelidade
        custo_adicional = 0.0
        label_custo = "Treinamento"
        if is_bling_plan: label_custo = "Implementação"

        if not is_bling_plan and not is_autoatendimento_plano and not is_em_branco_plano:
            limite_custo = 549.90
            if total_mensal_sem_fidelidade > 0 and total_mensal_sem_fidelidade < limite_custo:
                custo_adicional = limite_custo - total_mensal_sem_fidelidade

        # --- Atualizar Labels da UI ---
        mensal_sem_fid_str = f"{total_mensal_sem_fidelidade:.2f}".replace(".", ",")
        mensal_no_anual_str = f"{final_mensal_efetivo_anual:.2f}".replace(".", ",")
        anual_total_str = f"{final_anual_total:.2f}".replace(".", ",")
        custo_adic_str = f"{custo_adicional:.2f}".replace(".", ",")
        desconto_final_percent = round(max(0, desconto_aplicado_percent)) # Garante >= 0

        self.lbl_plano_mensal_sem_fid.config(text=f"Mensal (Sem Fidelidade): R$ {mensal_sem_fid_str}")
        if custo_adicional > 0.01:
            self.lbl_treinamento.config(text=f"+ Custo {label_custo}: R$ {custo_adic_str}")
            self.lbl_treinamento.pack(pady=(0, 5), anchor="w", padx=15) # Garante que está visível
        else:
            self.lbl_treinamento.pack_forget() # Oculta se for zero

        self.lbl_plano_mensal_no_anual.config(text=f"Mensal (no Plano Anual): R$ {mensal_no_anual_str}")
        self.lbl_plano_anual_total.config(text=f"Anual (Pagamento Único): R$ {anual_total_str}")
        self.lbl_desconto.config(text=f"Desconto Anual Aplicado: {desconto_final_percent}%")

        # --- Armazenar Valores Computados Finais ---
        self.computed_mensal_sem_fidelidade = total_mensal_sem_fidelidade
        self.computed_mensal_efetivo_anual = final_mensal_efetivo_anual
        self.computed_anual_total = final_anual_total
        self.computed_desconto_percent = desconto_final_percent
        self.computed_custo_adicional = custo_adicional

    def montar_lista_modulos(self):
        # (Sem alterações nesta função - ela monta a lista com base nos módulos ativos)
        linhas = []
        info = PLAN_INFO.get(self.current_plan, {})
        mandatory = info.get("mandatory", [])

        pdv_val = self.spin_pdv_var.get()
        if pdv_val > 0: linhas.append(f"{pdv_val}x PDV - Frente de Caixa")
        usr_val = self.spin_users_var.get()
        if usr_val > 0: linhas.append(f"{usr_val}x Usuários")
        smart_tef_val = self.spin_smart_tef_var.get()
        if smart_tef_val > 0:
             linhas.append(f"{smart_tef_val}x Smart TEF")
             if self.current_plan == "Gestão": linhas[-1] += " (Limite: 3)"

        for m in mandatory:
            if m not in ["PDV - Frente de Caixa", "Usuários", "Smart TEF"]:
                 linhas.append(f"1x {m}")

        for m, var_m in self.modules.items():
            if m in self.check_buttons and var_m.get() == 1 and m not in mandatory:
                 linhas.append(f"1x {m}")

        unique_mods = []
        for mod in linhas:
            if mod not in unique_mods: unique_mods.append(mod)

        montagem = "\n".join(f"•    {m}" for m in unique_mods)
        return montagem

    def gerar_dados_proposta(self, nome_closer, cel_closer, email_closer):
            """ Gera o dicionário de dados para preencher o slide da proposta. """
            nome_plano_selecionado = self.current_plan
            nome_plano_editado = self.nome_plano_var.get().strip()
            nome_plano_final = nome_plano_editado if nome_plano_editado else nome_plano_selecionado

            # --- Valores Formatados ---
            # Usar os valores já calculados e armazenados
            mensal_sem_fid_val = self.computed_mensal_sem_fidelidade
            mensal_efetivo_anual_val = self.computed_mensal_efetivo_anual
            anual_total_val = self.computed_anual_total
            custo_adicional_val = self.computed_custo_adicional
            desconto_percent_val = self.computed_desconto_percent

            # Strings formatadas
            mensal_sem_fid_str = f"R$ {mensal_sem_fid_val:.2f}".replace(".", ",")
            mensal_efetivo_anual_str = f"R$ {mensal_efetivo_anual_val:.2f}".replace(".", ",")
            anual_total_str = f"R$ {anual_total_val:.2f}".replace(".", ",")
            custo_adicional_str = f"R$ {custo_adicional_val:.2f}".replace(".", ",")

            # Adicionar custo adicional ao mensal sem fidelidade (se aplicável)
            label_custo = "Treinamento" if not self.current_plan.startswith("Bling") else "Implementação"
            if custo_adicional_val > 0.01:
                mensal_sem_fid_str += f" + {custo_adicional_str} ({label_custo})"

            # --- Definir Suporte --- (Lógica mantida da versão anterior)
            tipo_suporte = "Regular"; horario_suporte = "09:00 às 17:00 de Segunda a Sexta-feira"
            suporte_chat = self.modules.get("Suporte Técnico - Via chat", tk.IntVar()).get() == 1
            suporte_estendido = self.modules.get("Suporte Técnico - Estendido", tk.IntVar()).get() == 1
            if suporte_estendido:
                 tipo_suporte = "Estendido"; horario_suporte = "09:00 às 22:00 Seg-Sex & 11:00 às 21:00 Sab-Dom" # Abrev.
            elif suporte_chat:
                 tipo_suporte = "Chat Incluso"; horario_suporte = "09:00 às 22:00 Seg-Sex & 11:00 às 21:00 Sab-Dom"

            # --- Montar Lista de Módulos ---
            montagem = self.montar_lista_modulos()

            # --- Calcular Economia Anual ---
            economia_str = ""
            custo_total_mensalizado = (mensal_sem_fid_val * 12) + custo_adicional_val
            custo_total_anualizado = anual_total_val # Já é o total pago adiantado

            if custo_total_mensalizado > custo_total_anualizado + 0.01: # Evita R$ 0,00
                 economia_val = custo_total_mensalizado - custo_total_anualizado
                 econ_str = f"{economia_val:.2f}".replace(".", ",")
                 economia_str = f"Economia de R$ {econ_str} no plano anual" # Texto ajustado

            # --- Montar Dicionário Final para Placeholders ---
            # **IMPORTANTE:** As chaves aqui devem corresponder aos placeholders no seu PPTX. Ex: {plano_mensal_sem_fidelidade}
            dados = {
                "montagem_do_plano": montagem,
                "plano_mensal_sem_fidelidade": mensal_sem_fid_str,      # Ex: R$ 110,00 + R$ 439,90 (Treinamento)
                "plano_mensal_no_anual": mensal_efetivo_anual_str,    # Ex: R$ 99,00
                "plano_anual_total": anual_total_str,            # Ex: R$ 1188,00
                "custo_treinamento": custo_adicional_str if custo_adicional_val > 0.01 else "Incluso", # Ex: R$ 439,90 ou Incluso
                "desconto_aplicado": f"{desconto_percent_val}%",     # Ex: 10%
                "nome_do_plano": nome_plano_final,
                "tipo_de_suporte": tipo_suporte,
                "horario_de_suporte": horario_suporte,
                "validade_proposta": self.validade_proposta_var.get(),
                "nome_closer": nome_closer,
                "celular_closer": cel_closer,
                "email_closer": email_closer,
                "nome_cliente": self.nome_cliente_var.get(),
                "economia_anual": economia_str                   # Ex: Economia de R$ 120,00 no plano anual
            }
            return dados


# --- Funções gerar_proposta, gerar_material (sem alterações na lógica interna) ---
# Elas usarão os dados retornados por gerar_dados_proposta, que agora estão corretos.
# Lembre-se de ajustar os placeholders no seu PPTX para corresponder às chaves do dicionário 'dados' acima.
# Exemplo de placeholders no PPTX:
#   - Plano Mensal: {plano_mensal_sem_fidelidade}
#   - Plano Anual: {plano_anual_total} (Pagamento Único)
#   - (equivalente a {plano_mensal_no_anual} / mês)
#   - Módulos: {montagem_do_plano}
#   - Economia: {economia_anual}

# --- Função gerar_proposta (adaptada para novos placeholders se necessário) ---
def gerar_proposta(lista_abas, nome_closer, celular_closer, email_closer):
    ppt_file = "Proposta Comercial ConnectPlug.pptx"
    if not os.path.exists(ppt_file):
        showerror("Erro", f"Arquivo template '{ppt_file}' não encontrado!")
        return None
    try: prs = Presentation(ppt_file)
    except Exception as e: showerror("Erro", f"Falha ao abrir '{ppt_file}': {e}"); return None
    if not lista_abas: showerror("Erro", "Não há abas para gerar Proposta."); return None

    primeira_aba = lista_abas[0]
    dados_proposta = primeira_aba.gerar_dados_proposta(nome_closer, celular_closer, email_closer)

    for slide in prs.slides:
        substituir_placeholders_no_slide(slide, dados_proposta) # Usa a função atualizada

    nome_cliente_safe = dados_proposta.get("nome_cliente", "SemNome").replace("/", "-").replace("\\", "-")
    hoje_str = date.today().strftime("%d-%m-%Y")
    nome_arquivo = f"Proposta ConnectPlug - {nome_cliente_safe} - {hoje_str}.pptx"

    try:
        prs.save(nome_arquivo)
        showinfo("Sucesso", f"Proposta gerada: {nome_arquivo}")
        return nome_arquivo
    except Exception as e:
        showerror("Erro", f"Falha ao salvar '{nome_arquivo}':\n{e}")
        return None

# --- Função gerar_material (MAPEAMENTO_MODULOS continua crucial) ---
def gerar_material(lista_abas, nome_closer, celular_closer, email_closer):
    mat_file = "Material Tecnico ConnectPlug.pptx"
    if not os.path.exists(mat_file): showerror("Erro", f"Arquivo template '{mat_file}' não encontrado!"); return None
    try: prs = Presentation(mat_file)
    except Exception as e: showerror("Erro", f"Falha ao abrir '{mat_file}': {e}"); return None
    if not lista_abas: showerror("Erro", "Não há abas para gerar Material Técnico."); return None

    modulos_ativos_geral = set(); planos_usados_geral = set()
    dados_primeira_aba = lista_abas[0].gerar_dados_proposta(nome_closer, celular_closer, email_closer)

    for aba in lista_abas:
        planos_usados_geral.add(aba.current_plan)
        info_aba = PLAN_INFO.get(aba.current_plan, {}); mandatory_aba = info_aba.get("mandatory", [])
        for mod in mandatory_aba: modulos_ativos_geral.add(mod)
        for nome_mod, var_mod in aba.modules.items():
            if var_mod.get() == 1 and nome_mod not in mandatory_aba: modulos_ativos_geral.add(nome_mod)
        if aba.spin_pdv_var.get() > 0: modulos_ativos_geral.add("PDV - Frente de Caixa")
        if aba.spin_users_var.get() > 0: modulos_ativos_geral.add("Usuários")
        if aba.spin_smart_tef_var.get() > 0: modulos_ativos_geral.add("Smart TEF")

    # *** AJUSTE O MAPEAMENTO CONFORME SEU PPTX ***
    MAPEAMENTO_MODULOS = {
        "slide_sempre": None, "check_sistema_kds": "Relatório KDS", "check_Hub_de_Delivery": "Hub de Delivery",
        "check_integracao_api": "Integração API", "check_integracao_tap": "Integração Tap",
        "check_controle_de_mesas": "Controle de Mesas", "check_Delivery": "Delivery", "check_producao": "Produção",
        "check_Estoque_em_Grade": "Estoque em Grade", "check_Facilita_NFE": "Facilita NFE",
        "check_Importacao_de_xml": "Importação de XML", "check_conciliacao_bancaria": "Conciliação Bancária",
        "check_contratos_de_cartoes": "Contratos de cartões e outros", "check_ordem_de_servico": "Ordem de Serviço",
        "check_relatorio_dinamico": "Relatório Dinâmico", "check_programa_de_fidelidade": "Programa de Fidelidade",
        "check_business_intelligence": "Business Intelligence (BI)", "check_smartmenu": "Smart Menu",
        "check_backup_real_time": "Backup Realtime", "check_att_tempo_real": "Atualização em tempo real",
        "check_promocao": "Promoções", "check_marketing": "Marketing", "placeholder_pdv": "PDV - Frente de Caixa",
        "placeholder_smarttef": "Smart TEF", "placeholder_tef": "TEF",
        "placeholder_autoatendimento": "Terminais Autoatendimento", "placeholder_cardapio_digital": "Cardápio digital",
        "placeholder_app_gestao_cplug": "App Gestão CPlug", "check_delivery_direto_vip": "Delivery Direto VIP",
        "check_delivery_direto_profissional": "Delivery Direto Profissional",
        "placeholder_painel_senha_tv": "Painel Senha TV", "placeholder_painel_senha_mobile": "Painel Senha Mobile",
        "placeholder_suporte_chat": "Suporte Técnico - Via chat", "placeholder_suporte_estendido": "Suporte Técnico - Estendido",
        "placeholder_notas_fiscais": {"Notas Fiscais Ilimitadas", "30 Notas Fiscais"},
    }

    keep_slides = set()
    # (Lógica de manter/remover slides mantida - depende do MAPEAMENTO acima)
    for i, slide in enumerate(prs.slides):
        slide_mantido = False
        for shape in slide.shapes:
             if slide_mantido: break
             if shape.has_text_frame:
                 for paragraph in shape.text_frame.paragraphs:
                     if slide_mantido: break
                     for run in paragraph.runs:
                         if slide_mantido: break
                         txt_run = run.text.strip()
                         if not txt_run: continue
                         for placeholder, modulo_mapeado in MAPEAMENTO_MODULOS.items():
                             if placeholder in txt_run:
                                 if modulo_mapeado is None: slide_mantido = True; break
                                 elif isinstance(modulo_mapeado, set):
                                     if any(m in modulos_ativos_geral for m in modulo_mapeado): slide_mantido = True; break
                                 elif modulo_mapeado in modulos_ativos_geral: slide_mantido = True; break
                         if "slide_bling" in txt_run and any(p.startswith("Bling -") for p in planos_usados_geral): slide_mantido = True; break
        if slide_mantido: keep_slides.add(i)

    for idx in reversed(range(len(prs.slides))):
        if idx not in keep_slides:
            try:
                rId = prs.slides._sldIdLst[idx].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[idx]
            except Exception as e: print(f"Aviso: Não foi possível remover slide {idx}. Erro: {e}")

    for slide in prs.slides:
        substituir_placeholders_no_slide(slide, dados_primeira_aba) # Usa dados globais

    nome_cliente_safe = dados_primeira_aba.get("nome_cliente", "SemNome").replace("/", "-").replace("\\", "-")
    hoje_str = date.today().strftime("%d-%m-%Y")
    nome_arquivo = f"Material Tecnico ConnectPlug - {nome_cliente_safe} - {hoje_str}.pptx"

    try:
        prs.save(nome_arquivo)
        showinfo("Sucesso", f"Material Técnico gerado: {nome_arquivo}")
        return nome_arquivo
    except Exception as e:
        showerror("Erro", f"Falha ao salvar '{nome_arquivo}':\n{e}")
        return None

# --- Google Drive / Auth / Upload (SEM ALTERAÇÕES) ---
# (Código do Google Drive omitido para brevidade - é o mesmo da versão anterior)
SCOPES = ['https://www.googleapis.com/auth/drive']
def baixar_client_secret_remoto():
    # ... (código igual)
    url = "https://github.com/DevRGS/Gerador/raw/refs/heads/main/config/client_secret_788265418970-ur6f189oqvsttseeg6g77fegt0su67dj.apps.googleusercontent.com.json"
    nome_local = "client_secret_temp.json"
    if not os.path.exists(nome_local):
        print("Baixando client_secret do GitHub...")
        try:
            r = requests.get(url, timeout=15)
            r.raise_for_status()
            with open(nome_local, "w", encoding="utf-8") as f: f.write(r.text)
        except requests.exceptions.RequestException as e:
             showerror("Erro de Rede", f"Não foi possível baixar o client_secret.json:\n{e}")
             raise Exception(f"Erro ao baixar o client_secret.json: {e}")
    return nome_local

def get_gdrive_service():
    # ... (código igual)
    creds = None; token_file = 'token.json'; client_secret_file = None
    try: client_secret_file = baixar_client_secret_remoto()
    except Exception as e: showerror("Erro Crítico", f"Falha ao obter client_secret.\n{e}"); return None
    if os.path.exists(token_file):
        try:
            with open(token_file, 'rb') as token: creds = pickle.load(token)
        except (pickle.UnpicklingError, EOFError, FileNotFoundError):
             if os.path.exists(token_file): os.remove(token_file); print("Arquivo token removido.")
             creds = None
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try: creds.refresh(Request())
            except Exception as e:
                print(f"Erro ao renovar token: {e}. Reautenticando...")
                if os.path.exists(token_file): os.remove(token_file)
                creds = None
        if not creds:
             try:
                flow = InstalledAppFlow.from_client_secrets_file(client_secret_file, SCOPES)
                creds = flow.run_local_server(port=0)
             except Exception as e: showerror("Erro de Autenticação", f"Falha ao autenticar: {e}"); return None
        try:
            with open(token_file, 'wb') as token: pickle.dump(creds, token)
        except Exception as e: print(f"Aviso: Não foi possível salvar token.json: {e}")
    try:
        service = build('drive', 'v3', credentials=creds)
        return service
    except Exception as e: showerror("Erro Google API", f"Falha ao construir serviço: {e}"); return None

def upload_pptx_and_export_to_pdf(local_pptx_path):
    # ... (código igual)
    if not os.path.exists(local_pptx_path): showerror("Erro", f"Arquivo '{local_pptx_path}' não encontrado."); return
    service = get_gdrive_service()
    if not service: showerror("Erro Google Drive", "Não foi possível conectar ao Google Drive."); return
    pdf_output_name = local_pptx_path.replace(".pptx", ".pdf")
    base_name = os.path.basename(local_pptx_path)
    file_id = None # Para garantir que file_id exista no bloco finally
    try:
        print(f"Upload de '{base_name}'...")
        file_metadata = {'name': base_name, 'mimeType': 'application/vnd.google-apps.presentation'}
        media = MediaFileUpload(local_pptx_path, mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation', resumable=True)
        uploaded_file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        file_id = uploaded_file.get('id')
        if not file_id: raise Exception("Upload falhou (sem ID).")
        print(f"Upload concluído. ID: {file_id}. Exportando para PDF...")
        request = service.files().export_media(fileId=file_id, mimeType='application/pdf')
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request); done = False
        while not done: status, done = downloader.next_chunk(); # print(f"Download: {int(status.progress() * 100)}%")
        with open(pdf_output_name, 'wb') as f: f.write(fh.getvalue())
        print(f"PDF gerado: '{pdf_output_name}'.")
        showinfo("Google Drive", f"PDF gerado com sucesso:\n'{pdf_output_name}'.")
    except Exception as e:
        showerror("Erro Google Drive", f"Erro durante upload/conversão:\n{e}")
    finally:
        # Tenta deletar o arquivo temporário do Drive
        if file_id:
            try:
                print(f"Deletando arquivo temporário do Drive (ID: {file_id})...")
                service.files().delete(fileId=file_id).execute()
                print("Arquivo temporário deletado.")
            except Exception as delete_err:
                print(f"Aviso: Falha ao deletar arquivo temporário: {delete_err}")


# --- MainApp (SEM ALTERAÇÕES INTERNAS) ---
class MainApp(ttkb.Window):
    def __init__(self):
        super().__init__(themename="litera")
        self.title("Gerador de Propostas ConnectPlug v2.1") # Versão
        self.geometry("1200x800")
        self.nome_closer_var = tk.StringVar()
        self.celular_closer_var = tk.StringVar()
        self.email_closer_var = tk.StringVar()
        self.nome_cliente_var_shared = tk.StringVar(value="")
        self.validade_proposta_var_shared = tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
        carregar_config(self.nome_closer_var, self.celular_closer_var, self.email_closer_var)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        top_bar = ttkb.Frame(self); top_bar.pack(side="top", fill="x", pady=5, padx=5)
        ttkb.Label(top_bar, text="Vendedor:").pack(side="left", padx=(0, 2))
        ttkb.Entry(top_bar, textvariable=self.nome_closer_var, width=20).pack(side="left", padx=(0, 5))
        ttkb.Label(top_bar, text="Celular:").pack(side="left", padx=(0, 2))
        ttkb.Entry(top_bar, textvariable=self.celular_closer_var, width=15).pack(side="left", padx=(0, 5))
        ttkb.Label(top_bar, text="Email:").pack(side="left", padx=(0, 2))
        ttkb.Entry(top_bar, textvariable=self.email_closer_var, width=25).pack(side="left", padx=(0, 10))
        self.btn_add = ttkb.Button(top_bar, text="+ Nova Aba", command=self.add_aba, bootstyle="success")
        self.btn_add.pack(side="right", padx=5)
        self.notebook = ttkb.Notebook(self); self.notebook.pack(fill="both", expand=True, padx=5, pady=(0, 5))
        bot_frame = ttkb.Frame(self); bot_frame.pack(side="bottom", fill="x", pady=5, padx=5)
        ttkb.Button(bot_frame, text="Gerar Proposta + PDF", command=self.on_gerar_proposta, bootstyle="primary").pack(side="left", padx=5)
        ttkb.Button(bot_frame, text="Gerar Material + PDF", command=self.on_gerar_mat_tecnico, bootstyle="info").pack(side="left", padx=5)
        ttkb.Button(bot_frame, text="Gerar TUDO + PDF", command=self.on_gerar_tudo, bootstyle="secondary").pack(side="left", padx=5)
        self.abas_criadas = {}; self.ultimo_indice = 0
        self.add_aba()
        try:
            baixar_arquivo_if_needed("Proposta Comercial ConnectPlug.pptx", "https://github.com/DevRGS/Gerador/raw/refs/heads/main/assets/Proposta%20Comercial%20ConnectPlug.pptx")
            baixar_arquivo_if_needed("Material Tecnico ConnectPlug.pptx", "https://github.com/DevRGS/Gerador/raw/refs/heads/main/assets/Material%20Tecnico%20ConnectPlug.pptx")
        except Exception as e: showerror("Erro Download Templates", f"Não foi possível baixar arquivos base:\n{e}")

    def on_close(self): salvar_config(self.nome_closer_var.get(), self.celular_closer_var.get(), self.email_closer_var.get()); self.destroy()
    def add_aba(self):
        if len(self.abas_criadas) >= MAX_ABAS: showinfo("Limite Atingido", f"Máximo de {MAX_ABAS} abas."); return
        self.ultimo_indice += 1; idx = self.ultimo_indice
        frame_aba = PlanoFrame(self.notebook, idx, self.nome_cliente_var_shared, self.validade_proposta_var_shared, self.fechar_aba)
        self.notebook.add(frame_aba, text=f"Plano {idx}"); self.abas_criadas[idx] = frame_aba; self.notebook.select(frame_aba)
        if len(self.abas_criadas) >= MAX_ABAS: self.btn_add.config(state="disabled")
    def fechar_aba(self, indice):
        if indice in self.abas_criadas:
            frame_aba = self.abas_criadas[indice]
            try: self.notebook.forget(frame_aba); del self.abas_criadas[indice]
            except tk.TclError: pass
            if len(self.abas_criadas) < MAX_ABAS: self.btn_add.config(state="normal")
            if not self.abas_criadas: self.add_aba()
    def get_abas_ativas(self): indices_ativos = sorted(self.abas_criadas.keys()); return [self.abas_criadas[idx] for idx in indices_ativos]
    def _validar_dados_basicos(self):
        if not self.nome_closer_var.get() or not self.celular_closer_var.get() or not self.email_closer_var.get(): showerror("Dados Incompletos", "Preencha os dados do Vendedor."); return False
        if not self.nome_cliente_var_shared.get(): showerror("Dados Incompletos", "Preencha o Nome do Cliente."); return False
        return True
    def on_gerar_proposta(self):
        abas_ativas = self.get_abas_ativas();
        if not abas_ativas: showerror("Erro", "Nenhuma aba ativa."); return
        if not self._validar_dados_basicos(): return
        pptx_file = gerar_proposta(abas_ativas, self.nome_closer_var.get(), self.celular_closer_var.get(), self.email_closer_var.get())
        if pptx_file and os.path.exists(pptx_file): upload_pptx_and_export_to_pdf(pptx_file)
    def on_gerar_mat_tecnico(self):
        abas_ativas = self.get_abas_ativas();
        if not abas_ativas: showerror("Erro", "Nenhuma aba ativa."); return
        if not self._validar_dados_basicos(): return # Valida cliente também
        pptx_file = gerar_material(abas_ativas, self.nome_closer_var.get(), self.celular_closer_var.get(), self.email_closer_var.get())
        if pptx_file and os.path.exists(pptx_file): upload_pptx_and_export_to_pdf(pptx_file)
    def on_gerar_tudo(self):
        abas_ativas = self.get_abas_ativas();
        if not abas_ativas: showerror("Erro", "Nenhuma aba ativa."); return
        if not self._validar_dados_basicos(): return
        print("--- Gerando Proposta ---"); pdf_prop_ok = False
        pptx_prop = gerar_proposta(abas_ativas, self.nome_closer_var.get(), self.celular_closer_var.get(), self.email_closer_var.get())
        if pptx_prop and os.path.exists(pptx_prop):
            try: upload_pptx_and_export_to_pdf(pptx_prop); pdf_prop_ok = True
            except Exception as e: print(f"Erro PDF Proposta: {e}")
        print("--- Gerando Material Técnico ---"); pdf_mat_ok = False
        pptx_mat = gerar_material(abas_ativas, self.nome_closer_var.get(), self.celular_closer_var.get(), self.email_closer_var.get())
        if pptx_mat and os.path.exists(pptx_mat):
            try: upload_pptx_and_export_to_pdf(pptx_mat); pdf_mat_ok = True
            except Exception as e: print(f"Erro PDF Material: {e}")
        if pdf_prop_ok and pdf_mat_ok: print("--- Geração Concluída ---")
        else: print("--- Geração Concluída com possíveis erros no PDF ---")

# --- Função Principal ---
def main(): app = MainApp(); app.mainloop()
if __name__ == "__main__": main()