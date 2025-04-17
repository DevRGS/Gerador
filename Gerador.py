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
# Dados de Planos e Tabelas de Preço
# ---------------------------------------------------------
PLAN_INFO = {
    "Personalizado": {
        "base_mensal": 189.90,
        "base_anual": 170.91,
        "min_pdv": 1,
        "min_users": 2,
        "mandatory": ["Relatórios","Vendas - Estoque - Financeiro"]
    },
    "Ideal": {
        "base_mensal": 359.90,
        "base_anual": 323.91,
        "min_pdv": 1,
        "min_users": 5,
        "mandatory": [
            "3000 Notas Fiscais","Relatórios","Vendas - Estoque - Financeiro",
            "Estoque em Grade","Importação de XML","Produção"
        ]
    },
    "Completo": {
        "base_mensal": 549.90,
        "base_anual": 494.91,
        "min_pdv": 2,
        "min_users": 10,
        "mandatory": [
            "Conciliação Bancária","Contratos de cartões e outros","Controle de Mesas",
            "Delivery","Estoque em Grade","Facilita NFE","Importação de XML",
            "Notas Fiscais Ilimitadas","Ordem de Serviço","Produção","Relatório Dinâmico",
            "Relatórios","Vendas - Estoque - Financeiro"
        ]
    },
    "Autoatendimento": {
        "base_mensal": 0.0,
        "base_anual": 419.90,
        "min_pdv": 0,
        "min_users": 1,
        "mandatory": [
            "Contratos de cartões e outros","Estoque em Grade","Notas Fiscais Ilimitadas",
            "Produção","Vendas - Estoque - Financeiro"
        ]
    },
    "Bling": {
        "base_mensal": 369.80,
        "base_anual": 189.90,
        "min_pdv": 1,
        "min_users": 5,
        "mandatory": [
            "Relatórios",
            "Vendas - Estoque - Financeiro",
            "Notas Fiscais Ilimitadas"
        ]
    }
}

SEM_DESCONTO = {
    "TEF",
    "Autoatendimento",
    "Smart TEF",
    "Domínio Próprio",
    "Gestão de Entregadores",
    "Robô de WhatsApp + Recuperador de Pedido",
    "Gestão de Redes Sociais",
    "Combo de Logística",
    "Painel MultiLojas","Programa de Fidelidade","Integração API", "Integração TAP",
    "Central Telefônica (Base)","Central Telefônica (Por Loja)"
}

precos_mensais = {
    "Conciliação Bancária": 30.00,
    "Contratos de cartões e outros": 49.90,
    "Controle de Mesas": 30.00,
    "Delivery": 30.00,
    "Estoque em Grade": 30.00,
    "Importação de XML": 30.00,
    "Ordem de Serviço": 30.00,
    "Produção": 30.00,
    "Relatório Dinâmico": 59.90,
    "Notas Fiscais Ilimitadas": 119.90,
    "3000 Notas Fiscais": 0.0,

    "60 Notas Fiscais": 40.00,
    "120 Notas Fiscais": 70.00,
    "250 Notas Fiscais": 90.00,

    "TEF": 99.90,
    "Smart TEF": 49.90,
    "Backup Realtime": 99.90,
    "Atualização em Tempo Real": 49.90,
    "Business Intelligence (BI)": 99.90,
    "Hub de Delivery": 99.90,
    "Facilita NFE": 49.90,
    "Smart Menu": 99.90,
    "Cardápio Digital": 29.90,
    "Programa de Fidelidade": 299.90,
    "Autoatendimento": 299.90,
    "Delivery Direto Básico": 247.00,
    "Delivery Direto Profissional": 347.00,
    "Delivery Direto VIP": 497.00,
    "Promoções": 39.90,
    "Marketing": 49.90,
    "Painel de Senha": 49.90,
    "Integração TAP": 249.90,
    "Integração API": 199.90,
    "Relatório KDS": 29.90,
    "App Gestão CPlug": 19.90,
    "Domínio Próprio": 19.90,
    "Gestão de Entregadores": 19.90,
    "Robô de WhatsApp + Recuperador de Pedido": 99.90,
    "Gestão de Redes Sociais": 9.90,
    "Combo de Logística": 74.90,
    "Painel MultiLojas": 199.00,
    "Central Telefônica (Base)": 399.90,
    "Central Telefônica (Por Loja)": 49.90
}


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
                            txt = txt.replace(k, v)
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

        # Variáveis compartilhadas para Nome do Cliente e Validade
        self.nome_cliente_var = nome_cliente_var_shared
        self.validade_proposta_var = validade_proposta_var_shared

        self.current_plan = "Personalizado"
        self.spin_pdv_var = tk.IntVar(value=1)
        self.spin_users_var = tk.IntVar(value=1)
        self.spin_auto_var = tk.IntVar(value=0)
        self.spin_cardapio_var = tk.IntVar(value=0)
        self.spin_tef_var = tk.IntVar(value=0)
        self.spin_smart_tef_var = tk.IntVar(value=0)
        self.spin_app_cplug_var = tk.IntVar(value=0)
        self.spin_delivery_direto_basico_var = tk.IntVar(value=0)
        self.var_notas = tk.StringVar(value="NONE")

        # Módulos extras (checkboxes)
        self.modules = {}
        self.check_buttons = {}

        # Overrides de cálculo
        self.user_override_anual_active = tk.BooleanVar(value=False)
        self.user_override_discount_active = tk.BooleanVar(value=False)
        self.valor_anual_editavel = tk.StringVar(value="0.00")
        self.desconto_personalizado = tk.StringVar(value="0")

        # Armazenar valores
        self.computed_mensal = 0.0
        self.computed_anual = 0.0
        self.computed_desconto_percent = 0.0

        # Layout com scrollbar
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
        self.configurar_plano("Personalizado")

    def fechar_aba(self):
        if self.on_close_callback:
            self.on_close_callback(self.aba_index)

    def _montar_layout_esquerda(self):
        top_bar = ttkb.Frame(self.frame_left)
        top_bar.pack(fill="x", pady=5)
        ttkb.Label(top_bar, text=f"Aba Plano {self.aba_index}", font="-size 12 -weight bold").pack(side="left")
        btn_close = ttkb.Button(top_bar, text="Fechar Aba", command=self.fechar_aba)
        btn_close.pack(side="right")

        frame_planos = ttkb.Labelframe(self.frame_left, text="Planos")
        frame_planos.pack(fill="x", pady=5)
        for p in ["Personalizado","Ideal","Completo","Autoatendimento","Bling"]:
            ttkb.Button(frame_planos, text=p,
                        command=lambda pl=p: self.configurar_plano(pl)
                       ).pack(side="left", padx=5)

        frame_notas = ttkb.Labelframe(self.frame_left, text="Notas Fiscais")
        frame_notas.pack(fill="x", pady=5)
        f_nf = ttkb.Frame(frame_notas)
        f_nf.pack(fill="x", padx=5, pady=5)
        for nfopt in ["60","120","250"]:
            rb = ttk.Radiobutton(f_nf, text=nfopt+" Notas",
                                 variable=self.var_notas, value=nfopt,
                                 command=self.atualizar_valores)
            rb.pack(side="left", padx=5)

        # Módulos (checkboxes)
        self.modules = {
            "Relatórios": tk.IntVar(),
            "Vendas - Estoque - Financeiro": tk.IntVar(),
            "Conciliação Bancária": tk.IntVar(),
            "Contratos de cartões e outros": tk.IntVar(),
            "Controle de Mesas": tk.IntVar(),
            "Delivery": tk.IntVar(),
            "Estoque em Grade": tk.IntVar(),
            "Importação de XML": tk.IntVar(),
            "Ordem de Serviço": tk.IntVar(),
            "Produção": tk.IntVar(),
            "Relatório Dinâmico": tk.IntVar(),
            "Notas Fiscais Ilimitadas": tk.IntVar(),
            "Backup Realtime": tk.IntVar(),
            "Atualização em Tempo Real": tk.IntVar(),
            "Business Intelligence (BI)": tk.IntVar(),
            "Hub de Delivery": tk.IntVar(),
            "Facilita NFE": tk.IntVar(),
            "Smart Menu": tk.IntVar(),
            "Programa de Fidelidade": tk.IntVar(),
            "Delivery Direto Profissional": tk.IntVar(),
            "Delivery Direto VIP": tk.IntVar(),
            "Promoções": tk.IntVar(),
            "Marketing": tk.IntVar(),
            "Painel de Senha": tk.IntVar(),
            "Relatório KDS": tk.IntVar(),
            "Integração TAP": tk.IntVar(),
            "Integração API": tk.IntVar(),
            "Domínio Próprio": tk.IntVar(),
            "Gestão de Entregadores": tk.IntVar(),
            "Robô de WhatsApp + Recuperador de Pedido": tk.IntVar(),
            "Gestão de Redes Sociais": tk.IntVar(),
            "Combo de Logística": tk.IntVar(),
            "Painel MultiLojas": tk.IntVar(),
            "Central Telefônica (Base)": tk.IntVar(),
            "Central Telefônica (Por Loja)": tk.IntVar(),
            "3000 Notas Fiscais": tk.IntVar()
        }

        frame_mod = ttkb.Labelframe(self.frame_left, text="Outros Módulos")
        frame_mod.pack(fill="both", expand=True, pady=5)
        f_mod_cols = ttkb.Frame(frame_mod)
        f_mod_cols.pack(fill="both", expand=True)

        f_mod_left = ttkb.Frame(f_mod_cols)
        f_mod_left.pack(side="left", fill="both", expand=True, padx=5)
        f_mod_right = ttkb.Frame(f_mod_cols)
        f_mod_right.pack(side="left", fill="both", expand=True, padx=5)

        all_mods = sorted(self.modules.keys())
        mid = len(all_mods)//2
        left_side = all_mods[:mid]
        right_side = all_mods[mid:]
        self.check_buttons = {}

        for m in left_side:
            cb = ttk.Checkbutton(f_mod_left, text=m,
                                 variable=self.modules[m],
                                 command=self.atualizar_valores)
            cb.pack(anchor="w", pady=2)
            self.check_buttons[m] = cb

        for m in right_side:
            cb = ttk.Checkbutton(f_mod_right, text=m,
                                 variable=self.modules[m],
                                 command=self.atualizar_valores)
            cb.pack(anchor="w", pady=2)
            self.check_buttons[m] = cb

        frame_dados = ttkb.Labelframe(self.frame_left, text="Dados do Cliente")
        frame_dados.pack(fill="x", pady=5)
        ttkb.Label(frame_dados, text="Nome do Cliente:").grid(row=0, column=0, sticky="w")
        ttkb.Entry(frame_dados, textvariable=self.nome_cliente_var).grid(row=0, column=1, padx=5, pady=2)
        ttkb.Label(frame_dados, text="Validade Proposta:").grid(row=1, column=0, sticky="w")
        ttkb.Entry(frame_dados, textvariable=self.validade_proposta_var).grid(row=1, column=1, padx=5, pady=2)

    def _montar_layout_direita(self):
        frame_inc = ttkb.Labelframe(self.frame_right, text="Incrementos")
        frame_inc.pack(fill="x", pady=5)

        ttkb.Label(frame_inc, text="PDVs").grid(row=0, column=0, sticky="w")
        sp_pdv = ttkb.Spinbox(frame_inc, from_=0, to=99,
                              textvariable=self.spin_pdv_var,
                              command=self.atualizar_valores)
        sp_pdv.grid(row=0, column=1, padx=5, pady=2)

        ttkb.Label(frame_inc, text="Usuários").grid(row=1, column=0, sticky="w")
        sp_usr = ttkb.Spinbox(frame_inc, from_=0, to=999,
                              textvariable=self.spin_users_var,
                              command=self.atualizar_valores)
        sp_usr.grid(row=1, column=1, padx=5, pady=2)

        ttkb.Label(frame_inc, text="Autoatendimento").grid(row=2, column=0, sticky="w")
        sp_at = ttkb.Spinbox(frame_inc, from_=0, to=999,
                             textvariable=self.spin_auto_var,
                             command=self.atualizar_valores)
        sp_at.grid(row=2, column=1, padx=5, pady=2)

        ttkb.Label(frame_inc, text="Cardápio Digital").grid(row=3, column=0, sticky="w")
        sp_cd = ttkb.Spinbox(frame_inc, from_=0, to=999,
                             textvariable=self.spin_cardapio_var,
                             command=self.atualizar_valores)
        sp_cd.grid(row=3, column=1, padx=5, pady=2)

        ttkb.Label(frame_inc, text="TEF").grid(row=4, column=0, sticky="w")
        sp_tef = ttkb.Spinbox(frame_inc, from_=0, to=99,
                              textvariable=self.spin_tef_var,
                              command=self.atualizar_valores)
        sp_tef.grid(row=4, column=1, padx=5, pady=2)

        ttkb.Label(frame_inc, text="Smart TEF").grid(row=5, column=0, sticky="w")
        sp_smf = ttkb.Spinbox(frame_inc, from_=0, to=99,
                              textvariable=self.spin_smart_tef_var,
                              command=self.atualizar_valores)
        sp_smf.grid(row=5, column=1, padx=5, pady=2)

        ttkb.Label(frame_inc, text="App Gestão CPlug").grid(row=6, column=0, sticky="w")
        sp_app = ttkb.Spinbox(frame_inc, from_=0, to=999,
                              textvariable=self.spin_app_cplug_var,
                              command=self.atualizar_valores)
        sp_app.grid(row=6, column=1, padx=5, pady=2)

        ttkb.Label(frame_inc, text="Delivery Direto Básico").grid(row=7, column=0, sticky="w")
        sp_ddb = ttkb.Spinbox(frame_inc, from_=0, to=999,
                              textvariable=self.spin_delivery_direto_basico_var,
                              command=self.atualizar_valores)
        sp_ddb.grid(row=7, column=1, padx=5, pady=2)

        frame_valores = ttkb.Labelframe(self.frame_right, text="Valores Finais")
        frame_valores.pack(fill="x", pady=5)

        self.lbl_plano_mensal = ttkb.Label(frame_valores, text="Plano (Mensal): R$ 0.00", font="-size 12 -weight bold")
        self.lbl_plano_mensal.pack()
        self.lbl_plano_anual = ttkb.Label(frame_valores, text="Plano (Anual): R$ 0.00", font="-size 12 -weight bold")
        self.lbl_plano_anual.pack()
        self.lbl_treinamento = ttkb.Label(frame_valores, text="Custo Treinamento (Mensal): R$ 0.00", font="-size 12 -weight bold")
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

    def configurar_plano(self, plano):
        info = PLAN_INFO[plano]
        self.current_plan = plano
        self.spin_pdv_var.set(info["min_pdv"])
        self.spin_users_var.set(info["min_users"])
        if plano == "Autoatendimento":
            self.spin_auto_var.set(1)
        for m in self.modules:
            self.modules[m].set(0)
            if m in self.check_buttons:
                self.check_buttons[m].config(state='normal')
        for obrig in info["mandatory"]:
            if obrig in self.modules:
                self.modules[obrig].set(1)
                if obrig in self.check_buttons:
                    self.check_buttons[obrig].config(state='disabled')
        if plano=="Ideal":
            self.modules["3000 Notas Fiscais"].set(1)
            if "3000 Notas Fiscais" in self.check_buttons:
                self.check_buttons["3000 Notas Fiscais"].config(state='disabled')
        else:
            if "3000 Notas Fiscais" in self.check_buttons:
                self.modules["3000 Notas Fiscais"].set(0)
                self.check_buttons["3000 Notas Fiscais"].config(state='disabled')
        if plano!="Personalizado":
            self.var_notas.set("NONE")

        self.user_override_anual_active.set(False)
        self.user_override_discount_active.set(False)
        self.valor_anual_editavel.set("0.00")
        self.desconto_personalizado.set("0")
        self.atualizar_valores()

    def atualizar_valores(self, *args):
        try:
            info = PLAN_INFO[self.current_plan]
        except KeyError:
            return
        base_mensal = info["base_mensal"]
        mandatory = info["mandatory"]

        parte_descontavel = base_mensal
        parte_sem_desc = 0.0

        # Módulos com desconto
        for m, var_m in self.modules.items():
            if var_m.get() == 1:
                if (m not in SEM_DESCONTO) and (m not in mandatory):
                    parte_descontavel += precos_mensais.get(m, 0.0)

        # Notas
        if self.modules.get("3000 Notas Fiscais", tk.IntVar()).get() == 1:
            parte_descontavel += precos_mensais.get("3000 Notas Fiscais", 0.0)
        else:
            nf_opt = self.var_notas.get()
            if nf_opt in ["60","120","250"]:
                parte_descontavel += precos_mensais.get(nf_opt+" Notas Fiscais", 0.0)

        # PDVs e Users
        pdv_extras = max(0, self.spin_pdv_var.get() - info["min_pdv"])
        if self.current_plan == "Bling":
            parte_descontavel += pdv_extras * 40.00
        else:
            parte_descontavel += pdv_extras * 59.90

        user_extras = max(0, self.spin_users_var.get() - info["min_users"])
        parte_descontavel += user_extras * 20.00

        # Módulos sem desconto
        for m, var_m in self.modules.items():
            if var_m.get() == 1 and (m in SEM_DESCONTO) and (m not in mandatory):
                parte_sem_desc += precos_mensais.get(m, 0.0)

        # TEF
        parte_sem_desc += self.spin_tef_var.get()*99.90
        parte_sem_desc += self.spin_smart_tef_var.get()*49.90

        # Autoatendimento
        auto_qty = self.spin_auto_var.get()
        if self.current_plan == "Autoatendimento":
            if auto_qty >=1:
                parte_sem_desc += 419.90 + (auto_qty-1)*399.90
        else:
            parte_sem_desc += auto_qty*299.90

        # App Gestão CPlug
        parte_descontavel += self.spin_app_cplug_var.get()*19.90

        # Delivery Direto Básico
        parte_descontavel += self.spin_delivery_direto_basico_var.get()*247.00

        # Cardápio Digital
        card_qt = self.spin_cardapio_var.get()
        if card_qt == 1:
            parte_descontavel += 29.90
        elif card_qt>1:
            parte_descontavel += card_qt*24.90

        valor_mensal_automatico = parte_descontavel + parte_sem_desc

        # Cálculo Anual
        if self.user_override_anual_active.get():
            try:
                final_anual = float(self.valor_anual_editavel.get())
            except ValueError:
                final_anual = valor_mensal_automatico
                self.valor_anual_editavel.set(f"{final_anual:.2f}")
        elif self.user_override_discount_active.get():
            try:
                desc_custom = float(self.desconto_personalizado.get())
            except ValueError:
                desc_custom = 0.0
            desc_dec = desc_custom / 100.0
            final_anual = (parte_descontavel*(1-desc_dec)) + parte_sem_desc
            self.valor_anual_editavel.set(f"{final_anual:.2f}")
        else:
            desc_padrao = 0.10
            final_anual = (parte_descontavel*(1-desc_padrao)) + parte_sem_desc
            self.valor_anual_editavel.set(f"{final_anual:.2f}")

        # Custo treinamento
        if self.current_plan == "Autoatendimento":
            training_cost = 0.0
        else:
            if valor_mensal_automatico < 549.90:
                training_cost = 549.90 - valor_mensal_automatico
            else:
                training_cost = 0.0

        # Atualização das labels
        if self.current_plan == "Autoatendimento":
            self.lbl_plano_mensal.config(text="Plano (Mensal): Não disponível")
        else:
            self.lbl_plano_mensal.config(text=f"Plano (Mensal): R$ {valor_mensal_automatico:.2f}")

        self.lbl_plano_anual.config(text=f"Plano (Anual): R$ {final_anual:.2f}")
        self.lbl_treinamento.config(text=f"Custo Treinamento (Mensal): R$ {training_cost:.2f}")

        if valor_mensal_automatico>0:
            desconto_calc = ((valor_mensal_automatico - final_anual)/valor_mensal_automatico)*100
        else:
            desconto_calc = 0.0
        self.lbl_desconto.config(text=f"Desconto: {round(desconto_calc)}%")

        self.computed_mensal = valor_mensal_automatico
        self.computed_anual = final_anual
        self.computed_desconto_percent = round(desconto_calc)

    def montar_lista_modulos(self):
        linhas = []
        inc = []
        
        # Planos com PDVs e cortesias de usuário
        pdv_val = self.spin_pdv_var.get()
        if pdv_val > 0:
            inc.append(f"{pdv_val} PDVs")
            # Adiciona "Usuário Cortesia" para PDVs extras em planos específicos
            if self.current_plan in ["Personalizado", "Ideal", "Completo"]:
                min_pdv = PLAN_INFO[self.current_plan]["min_pdv"]
                pdv_extras = max(0, pdv_val - min_pdv)
                if pdv_extras > 0:
                    inc.append(f"{pdv_extras} Usuário{'s' if pdv_extras > 1 else ''} Cortesia")

        # Usuários
        usr_val = self.spin_users_var.get()
        if usr_val > 0:
            inc.append(f"{usr_val} Usuários")

        # Autoatendimento com TEF Cortesia
        aut_val = self.spin_auto_var.get()
        if aut_val > 0:
            inc.append(f"{aut_val} Autoatendimento")
            # Adiciona "TEF Cortesia" por terminal de Autoatendimento
            if aut_val >= 1:
                inc.append(f"{aut_val} TEF Cortesia")

        # Outros incrementos
        card_val = self.spin_cardapio_var.get()
        if card_val > 0:
            inc.append(f"{card_val} Cardápio(s) Digital(is)")
        
        qtd_tef = self.spin_tef_var.get()
        if qtd_tef > 0:
            inc.append(f"{qtd_tef} TEF")
        
        smqtd_tef = self.spin_smart_tef_var.get()
        if smqtd_tef > 0:
            inc.append(f"{smqtd_tef} Smart TEF")
        
        app_val = self.spin_app_cplug_var.get()
        if app_val > 0:
            inc.append(f"{app_val} App Gestão CPlug")
        
        ddb_val = self.spin_delivery_direto_basico_var.get()
        if ddb_val > 0:
            inc.append(f"{ddb_val} Delivery Direto Básico")

        # Notas Fiscais (prioriza Ilimitadas sobre 3000)
        if self.modules["Notas Fiscais Ilimitadas"].get() == 1:
            inc.append("Notas Fiscais Ilimitadas")
        else:
            if self.modules["3000 Notas Fiscais"].get() == 1:
                inc.append("3000 Notas Fiscais")
            else:
                opt = self.var_notas.get()
                if opt in ["60", "120", "250"]:
                    inc.append(f"{opt} Notas Fiscais")

        # Adiciona incrementos à lista principal
        if inc:
            linhas.extend(inc)

        # Módulos obrigatórios
        linhas.append("Relatórios")
        linhas.append("Vendas - Estoque - Financeiro")

        # Módulos adicionais (checkboxes)
        cbox = []
        for m, var_m in self.modules.items():
            if var_m.get() == 1:
                # Excluir duplicados e módulos já listados nos incrementos
                if m not in ["Relatórios", "Vendas - Estoque - Financeiro", "3000 Notas Fiscais", "TEF", "Smart TEF", "Autoatendimento", "Cardápio Digital"]:
                    cbox.append(m)
        if cbox:
            linhas.extend(cbox)

        unique_mods = []
        for mod in linhas:
            if mod not in unique_mods:
                unique_mods.append(mod)
        return unique_mods

    def gerar_dados_proposta(self, nome_closer, cel_closer, email_closer):
            valor_mensal = self.computed_mensal
            valor_anual = self.computed_anual
            desconto_percent = self.computed_desconto_percent

            if self.current_plan == "Autoatendimento":
                plano_mensal_str = "Não Disponível"
                training_cost = 0.0
            else:
                training_cost = 0.0
                if valor_mensal < 549.90:
                    training_cost = 549.90 - valor_mensal
                if training_cost > 0:
                    part_mensal = f"{valor_mensal:.2f}".replace(".", ",")
                    part_training = f"{training_cost:.2f}".replace(".", ",")
                    plano_mensal_str = f"R$ {part_mensal} + R$ {part_training}"
                else:
                    plano_mensal_str = f"R$ {valor_mensal:.2f}".replace(".", ",")

            plano_anual_str = f"R$ {valor_anual:.2f}".replace(".", ",")

            if valor_anual >= 269.90:
                tipo_suporte = "Estendido"
                horario_suporte = "09:00 às 22:00 de Segunda a Sexta-feira & Sábado e Domingo das 11:00 às 21:00"
            else:
                tipo_suporte = "Regular"
                horario_suporte = "09:00 às 17:00 de Segunda a Sexta-feira"

            lista_mods = self.montar_lista_modulos()
            montagem = "\n".join(f"•    {m}" for m in lista_mods)

            if self.current_plan == "Autoatendimento":
                economia_str = ""
            else:
                custo_anual_mensalizado = valor_mensal * 12 + training_cost
                custo_anual_plano = valor_anual * 12
                economia_val = custo_anual_mensalizado - custo_anual_plano
                if economia_val > 0:
                    econ = f"{economia_val:.2f}".replace(".", ",")
                    economia_str = f"Economia de R$ {econ} ao ano"
                else:
                    economia_str = ""

            dados = {
                "montagem_do_plano": montagem,
                "plano_mensal": plano_mensal_str,
                "plano_anual": plano_anual_str,
                "desconto_total": f"{desconto_percent}%",
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
# Funções que geram .pptx (Proposta e Material)
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

    # 1) Descobrir quais slides manter (opcional)
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
    
    # 2) Remover slides não mantidos
    for idx in reversed(range(len(prs.slides))):
        if idx not in keep_slides:
            rid = prs.slides._sldIdLst[idx].rId
            prs.part.drop_rel(rid)
            del prs.slides._sldIdLst[idx]
    
    # 3) Re-mapear índices
    sorted_kept = sorted(keep_slides)
    new_order_map = {}
    for new_idx, slide in enumerate(prs.slides):
        old_idx = sorted_kept[new_idx]
        new_order_map[new_idx] = slide_map_aba[old_idx]
    
    # 4) Substituir placeholders
    dados_de_aba = {}
    for aba in lista_abas:
        d = aba.gerar_dados_proposta(nome_closer, celular_closer, email_closer)
        dados_de_aba[aba.aba_index] = d
    
    # Fallback: se não tiver slides específicos, use dados da primeira aba
    fallback_aba = lista_abas[0]
    d_fallback = dados_de_aba[fallback_aba.aba_index]

    for new_idx, slide in enumerate(prs.slides):
        aba_num = new_order_map[new_idx]
        if aba_num is None:
            substituir_placeholders_no_slide(slide, d_fallback)
        else:
            d_aba = dados_de_aba[aba_num]
            substituir_placeholders_no_slide(slide, d_aba)
    
    # 5) Salvar
    nome_cliente_primeira = d_fallback.get("nome_cliente", "SemNome")
    hoje_str = date.today().strftime("%d-%m-%Y")
    nome_arquivo = f"Proposta ConnectPlug - {nome_cliente_primeira} - {hoje_str}.pptx"

    try:
        prs.save(nome_arquivo)
        showinfo("Sucesso", f"Proposta gerada: {nome_arquivo}")
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
    # 1) Descobrir módulos ativos e planos usados
    # ---------------------------------------------------
    modulos_ativos = set()
    planos_usados = set()

    for aba in lista_abas:
        planos_usados.add(aba.current_plan)

        # Módulos de checkboxes
        for nome_mod, var_mod in aba.modules.items():
            if var_mod.get() == 1:
                modulos_ativos.add(nome_mod)

        # Incrementos
        if aba.spin_tef_var.get() > 0:
            modulos_ativos.add("TEF")
        if aba.spin_smart_tef_var.get() > 0:
            modulos_ativos.add("Smart TEF")
        if aba.spin_auto_var.get() > 0:
            modulos_ativos.add("Autoatendimento")
        if aba.spin_cardapio_var.get() > 0:
            modulos_ativos.add("Cardápio Digital")
        if aba.spin_app_cplug_var.get() > 0:
            modulos_ativos.add("App Gestão CPlug")
        if aba.spin_delivery_direto_basico_var.get() > 0:
            modulos_ativos.add("Delivery Direto Básico")
        if aba.spin_pdv_var.get() > 0:
            modulos_ativos.add("PDV")
        # etc., se houver outros increments.

    # ---------------------------------------------------
    # 2) Mapeamento de placeholders para módulos
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
        "pdv_balcao": "PDV",
        "qtd_smarttef": "Smart TEF",
        "qtd_tef": "TEF",
        "qtd_autoatendimento": "Autoatendimento",
        "qtd_cardapio_digital": "Cardápio Digital",
        "qtd_app_gestao_cplug": "App Gestão CPlug",
        "qtd_delivery_direto_basico": "Delivery Direto Básico",
        "check_delivery_direto_vip": "Delivery Direto VIP",
        "check_delivery_direto_profissional": "Delivery Direto Profissional",
        "check_notas_fiscais": {
            "Notas Fiscais Ilimitadas",
            "60 Notas Fiscais",
            "120 Notas Fiscais",
            "250 Notas Fiscais",
            "3000 Notas Fiscais"
        },
    }

    # ---------------------------------------------------
    # 3) Decidir quais slides manter
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

        # Flag para saber se iremos manter este slide
        slide_ok = False

        # Caso tenha "slide_bling" e Bling esteja em use, etc.
        if "slide_bling" in full_txt:
            if "Bling" in planos_usados:
                slide_ok = True

        # Se tiver "slide_sempre" => mantém sempre
        if "slide_sempre" in full_txt:
            slide_ok = True

        # Agora checamos placeholders de módulos do MAPEAMENTO_MODULOS
        for placeholder, nome_modulo in MAPEAMENTO_MODULOS.items():
            if placeholder in full_txt:
                # Se o placeholder está no slide, mantenha só se
                # esse nome_modulo estiver em modulos_ativos
                if nome_modulo in modulos_ativos:
                    slide_ok = True
                    # Não fazemos break aqui, pois podem haver vários placeholders
                    # no mesmo slide. Mas se quiser, pode fazer break se
                    # bastar um para manter.

        # Exemplo: Se não entrou em nenhum if e slide_ok ainda está False,
        # esse slide NÃO será mantido.
        if slide_ok:
            keep_slides.add(i)

    # 4) Remove slides não mantidos
    for idx in reversed(range(len(prs.slides))):
        if idx not in keep_slides:
            rid = prs.slides._sldIdLst[idx].rId
            prs.part.drop_rel(rid)
            del prs.slides._sldIdLst[idx]

    # ---------------------------------------------------
    # 5) Substituir placeholders (dados globais)
    # ---------------------------------------------------
    fallback_aba = lista_abas[0]
    d_fallback = fallback_aba.gerar_dados_proposta(nome_closer, celular_closer, email_closer)

    for slide in prs.slides:
        substituir_placeholders_no_slide(slide, d_fallback)

    # ---------------------------------------------------
    # 6) Salvar pptx final
    # ---------------------------------------------------
    nome_cliente_primeira = d_fallback.get("nome_cliente", "SemNome")
    hoje_str = date.today().strftime("%d-%m-%Y")
    nome_arquivo = f"Material Tecnico ConnectPlug - {nome_cliente_primeira} - {hoje_str}.pptx"

    try:
        prs.save(nome_arquivo)
        showinfo("Sucesso", f"Material Técnico gerado: {nome_arquivo}")
        return nome_arquivo
    except Exception as e:
        showerror("Erro", f"Falha ao salvar: {e}")
        return None



# ---------------------------------------------------------
# Google Drive / Auth
# ---------------------------------------------------------
SCOPES = ['https://www.googleapis.com/auth/drive']

def baixar_client_secret_remoto():
    """Baixa o client_secret.json do repositório no GitHub se ainda não estiver salvo localmente."""
    url = "https://github.com/DevRGS/Gerador/raw/refs/heads/main/config/client_secret_788265418970-ur6f189oqvsttseeg6g77fegt0su67dj.apps.googleusercontent.com.json"
    nome_local = "client_secret_temp.json"

    if not os.path.exists(nome_local):
        print("Baixando client_secret do GitHub...")
        r = requests.get(url)
        if r.status_code == 200:
            with open(nome_local, "w", encoding="utf-8") as f:
                f.write(r.text)
        else:
            raise Exception(f"Erro ao baixar o client_secret.json: {r.status_code}")
    
    return nome_local

def get_gdrive_service():
    """Autentica e retorna o serviço do Google Drive com base no client_secret remoto."""
    creds = None
    CLIENT_SECRET_FILE = baixar_client_secret_remoto()
    TOKEN_FILE = 'token.json'

    # Tenta carregar o token salvo anteriormente
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'rb') as token:
            creds = pickle.load(token)

    # Se não houver token ou ele for inválido/expirado, roda o fluxo de autenticação
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        # Salva o token local para reutilização
        with open(TOKEN_FILE, 'wb') as token:
            pickle.dump(creds, token)

    # Constrói o serviço Google Drive autenticado
    service = build('drive', 'v3', credentials=creds)
    return service

def upload_pptx_and_export_to_pdf(local_pptx_path):
    """
    Faz upload do .pptx convertendo em Google Slides, 
    e baixa PDF local trocando .pptx -> .pdf
    """
    if not os.path.exists(local_pptx_path):
        showerror("Erro", f"Arquivo {local_pptx_path} não foi encontrado.")
        return

    service = get_gdrive_service()
    pdf_output_name = local_pptx_path.replace(".pptx", ".pdf")

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
    file_id = uploaded_file.get('id')
    print(f"Arquivo '{local_pptx_path}' enviado como Google Slides. ID: {file_id}")

    # 2) Exportar para PDF
    request = service.files().export_media(fileId=file_id, mimeType='application/pdf')
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
        if status:
            print(f"Progresso PDF: {int(status.progress() * 100)}%")

    with open(pdf_output_name, 'wb') as f:
        f.write(fh.getvalue())

    showinfo("Google Drive", f"PDF gerado localmente: '{pdf_output_name}'.")


# ---------------------------------------------------------
# MainApp
# ---------------------------------------------------------
class MainApp(ttkb.Window):
    def __init__(self):
        super().__init__(themename="litera")
        self.title("Proposta + Material Técnico + PDF unificado")
        self.geometry("1200x800")

        self.nome_closer_var = tk.StringVar()
        self.celular_closer_var = tk.StringVar()
        self.email_closer_var = tk.StringVar()

        # Variáveis compartilhadas para TODAS as abas
        self.nome_cliente_var_shared = tk.StringVar(value="Cliente Compartilhado")
        self.validade_proposta_var_shared = tk.StringVar(value="DD/MM/YYYY")

        carregar_config(self.nome_closer_var, self.celular_closer_var, self.email_closer_var)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # Top bar
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
        baixar_arquivo_if_needed(
        "Proposta Comercial ConnectPlug.pptx",
        "https://github.com/DevRGS/Gerador/raw/refs/heads/main/assets/Proposta%20Comercial%20Connectplug.pptx")

        baixar_arquivo_if_needed(
        "Material Tecnico ConnectPlug.pptx",
        "https://github.com/DevRGS/Gerador/raw/refs/heads/main/assets/Material%20Tecnico%20ConnectPlug.pptx" )

    def on_close(self):
        salvar_config(
            self.nome_closer_var.get(),
            self.celular_closer_var.get(),
            self.email_closer_var.get()
        )
        self.destroy()

    def add_aba(self):
        if len(self.abas_criadas) >= MAX_ABAS:
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

    def fechar_aba(self, indice):
        if indice in self.abas_criadas:
            frame_aba = self.abas_criadas[indice]
            self.notebook.forget(frame_aba)
            del self.abas_criadas[indice]

    def get_abas_ativas(self):
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
