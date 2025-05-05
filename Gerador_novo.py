# -*- coding: utf-8 -*-
import sys
import subprocess
import importlib
import os
import requests
import io
import pickle
import json
from datetime import date
import traceback # Para debugging

# --- Instalação de Dependências ---
dependencias = {
    "ttkbootstrap": "ttkbootstrap",
    "pptx": "python-pptx",
    "googleapiclient": "google-api-python-client",
    "google.auth.transport": "google-auth-httplib2",
    "google_auth_oauthlib": "google-auth-oauthlib",
    "requests": "requests"
}

print("Verificando dependências...")
for modulo, pacote in dependencias.items():
    try:
        importlib.import_module(modulo)
        # print(f"  [OK] {modulo}") # Descomente para verbosidade
    except ImportError:
        print(f"  Instalando '{pacote}' (módulo '{modulo}' não encontrado)...")
        try:
            # Tenta usar 'pip' diretamente, mais comum
            subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", pacote], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL) # Esconde output
            # subprocess.check_call([sys.executable, "-m", "pip", "install", pacote]) # Mostra output
            importlib.invalidate_caches()
            importlib.import_module(modulo)
            print(f"  '{pacote}' instalado com sucesso.")
        except subprocess.CalledProcessError as e:
             print(f"  ERRO: Falha ao instalar '{pacote}'. Verifique pip/conexão. Erro: {e}")
             showerror("Erro de Dependência", f"Falha ao instalar o pacote '{pacote}'.\nVerifique sua conexão com a internet e a configuração do pip.\nO programa pode não funcionar corretamente.\n\nErro: {e}")
             # Considerar sair: sys.exit(f"Erro Crítico: Não foi possível instalar {pacote}")
        except ImportError:
            print(f"  ERRO: Falha ao importar '{modulo}' após instalação.")
            showerror("Erro de Dependência", f"Falha ao importar o módulo '{modulo}' mesmo após tentar a instalação.\nO programa pode não funcionar corretamente.")
            # Considerar sair: sys.exit(f"Erro Crítico: Não foi possível importar {modulo}")
        except Exception as e:
             print(f"  ERRO inesperado ao instalar/importar {pacote}: {e}")
             showerror("Erro Inesperado", f"Ocorreu um erro inesperado durante a instalação/verificação de dependências:\n{e}")


# --- Imports Principais ---
try:
    import tkinter as tk
    from tkinter import ttk
    import ttkbootstrap as ttkb
    from tkinter.messagebox import showerror, showinfo
    from pptx import Presentation
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
except ImportError as e:
    showerror("Erro Fatal", f"Não foi possível importar um módulo essencial: {e}\nO programa não pode continuar.")
    sys.exit(f"Erro crítico de importação: {e}")


# --- Funções Utilitárias ---
def baixar_arquivo_if_needed(nome_arquivo, url):
    if not os.path.exists(nome_arquivo):
        print(f"Baixando {nome_arquivo}...")
        try:
            # Timeout aumentado e verificação de status
            r = requests.get(url, timeout=30, stream=True)
            r.raise_for_status()
            with open(nome_arquivo, "wb") as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
            print(f"{nome_arquivo} baixado com sucesso.")
        except requests.exceptions.Timeout:
            showerror("Erro de Download", f"Tempo esgotado ao tentar baixar {nome_arquivo}.\nVerifique sua conexão.")
            raise ConnectionError(f"Timeout ao baixar {nome_arquivo}")
        except requests.exceptions.RequestException as e:
            showerror("Erro de Download", f"Não foi possível baixar {nome_arquivo}.\nErro: {e}")
            raise ConnectionError(f"Erro ao baixar {nome_arquivo}: {e}")

# ---------------------------------------------------------
# Ajustes Globais
# ---------------------------------------------------------
try:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    print("Diretório de trabalho:", os.getcwd())
except Exception as e:
    print(f"Erro ao definir diretório de trabalho: {e}")
    # Continuar mesmo assim? Ou sair?

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
            print(f"Aviso: Arquivo de configuração '{CONFIG_FILE}' inválido ou corrompido.")
        except Exception as e:
             print(f"Erro ao carregar configuração de '{CONFIG_FILE}': {e}")

def salvar_config(nome_closer, celular_closer, email_closer):
    dados = {"nome_vendedor": nome_closer, "celular_vendedor": celular_closer, "email_vendedor": email_closer}
    try:
        # Tenta garantir permissão de escrita
        if os.path.exists(CONFIG_FILE):
            try: os.chmod(CONFIG_FILE, 0o666)
            except OSError: pass # Ignora se não puder mudar
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(dados, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Erro ao salvar configuração em '{CONFIG_FILE}': {e}")
        # showerror("Erro ao Salvar", f"Não foi possível salvar a configuração:\n{e}") # Opcional

# ---------------------------------------------------------
# Dados de Planos e Tabelas de Preço (Estrutura Mantida)
# ---------------------------------------------------------
LISTA_PLANOS_UI = ["PDV Básico", "Gestão", "Performance", "Autoatendimento", "Bling", "Em Branco"]
LISTA_PLANOS_BLING = ["Bling - Básico", "Bling - Com Estoque em Grade"] # Revisar se Bling ainda é relevante

PLAN_INFO = {
    "PDV Básico": {
        "base_mensal": 99.00, "min_pdv": 1, "min_users": 2, "max_extra_users": 1, "max_extra_pdvs": 0,
        "mandatory": ["Usuários", "30 Notas Fiscais", "Suporte Técnico - Via chamados", "Relatório Básico", "PDV - Frente de Caixa"],
        "allowed_optionals": ["Smart Menu", "Terminais Autoatendimento", "Hub de Delivery", "Delivery Direto Profissional", "Delivery Direto VIP", "TEF", "Importação de XML", "Cardápio digital"]
    },
    "Gestão": {
        "base_mensal": 199.00, "min_pdv": 2, "min_users": 3, "max_extra_users": 2, "max_extra_pdvs": 1,
        "mandatory": ["Notas Fiscais Ilimitadas", "Importação de XML", "PDV - Frente de Caixa", "Usuários", "Painel Senha TV", "Estoque em Grade", "Relatórios", "Suporte Técnico - Via chamados", "Suporte Técnico - Via chat", "Delivery", "Relatório KDS"],
        "allowed_optionals": ["Facilita NFE", "Conciliação Bancária", "Contratos de cartões e outros", "Delivery Direto Profissional", "Delivery Direto VIP", "TEF", "Integração API", "Business Intelligence (BI)", "Backup Realtime", "Cardápio digital", "Smart Menu", "Hub de Delivery", "Ordem de Serviço", "App Gestão CPlug", "Painel Senha Mobile", "Controle de Mesas", "Produção", "Promoções", "Marketing", "Relatório Dinâmico", "Atualização em tempo real", "Smart TEF", "Terminais Autoatendimento", "Suporte Técnico - Estendido"]
    },
    "Performance": {
        "base_mensal": 499.00, "min_pdv": 3, "min_users": 5, "max_extra_users": 5, "max_extra_pdvs": 2,
        "mandatory": ["Produção", "Promoções", "Notas Fiscais Ilimitadas", "Importação de XML", "Hub de Delivery", "Ordem de Serviço", "Delivery", "App Gestão CPlug", "Relatório KDS", "Painel Senha TV", "Painel Senha Mobile", "Controle de Mesas", "Estoque em Grade", "Marketing", "Relatórios", "Relatório Dinâmico", "Atualização em tempo real", "Facilita NFE", "Conciliação Bancária", "Contratos de cartões e outros", "Suporte Técnico - Via chamados", "Suporte Técnico - Via chat", "Suporte Técnico - Estendido", "PDV - Frente de Caixa", "Smart TEF", "Usuários"],
        "allowed_optionals": ["TEF", "Programa de Fidelidade", "Integração Tap", "Integração API", "Business Intelligence (BI)", "Backup Realtime", "Cardápio digital", "Smart Menu", "Terminais Autoatendimento", "Delivery Direto Profissional", "Delivery Direto VIP"]
    },
    "Autoatendimento": { # Revisar lógica e módulos/preços se necessário
        "base_mensal": 419.90, "min_pdv": 0, "min_users": 1, "max_extra_users": 998, "max_extra_pdvs": 99,
        "mandatory": ["Contratos de cartões e outros", "Estoque em Grade", "Notas Fiscais Ilimitadas", "Produção"],
        "allowed_optionals": []
    },
    "Bling - Básico": { # Revisar lógica e módulos/preços se necessário
        "base_mensal": 189.90, "min_pdv": 1, "min_users": 5, "max_extra_users": 994, "max_extra_pdvs": 98,
        "mandatory": ["Relatórios", "Notas Fiscais Ilimitadas"],
        "allowed_optionals": [],
    },
    "Bling - Com Estoque em Grade": { # Revisar lógica e módulos/preços se necessário
        "base_mensal": 219.90, "min_pdv": 1, "min_users": 5, "max_extra_users": 994, "max_extra_pdvs": 98,
        "mandatory": ["Relatórios", "Notas Fiscais Ilimitadas", "Estoque em Grade"],
        "allowed_optionals": [],
    },
    "Em Branco": {
        "base_mensal": 0.0, "min_pdv": 0, "min_users": 0, "max_extra_users": 999, "max_extra_pdvs": 99, "mandatory": [],
        # Permite todos os módulos com preço como opcionais?
        "allowed_optionals": [m for m in precos_mensais.keys() if precos_mensais[m] > 0] # Exemplo: todos com preço > 0
    }
}

# Módulos SEM DESCONTO (quando desconto manual % é aplicado)
SEM_DESCONTO = {
    "TEF", "Terminais Autoatendimento", "Smart TEF", "Delivery Direto Profissional",
    "Delivery Direto VIP", "Programa de Fidelidade", "Integração Tap", "Integração API",
    "Business Intelligence (BI)", "Backup Realtime", "Hub de Delivery", "Smart Menu",
}

# Preços MENSAIS dos Módulos/Extras
precos_mensais = {
    "Importação de XML": 29.00, "Produção": 30.00, "Promoções": 24.50, "Hub de Delivery": 79.00,
    "Ordem de Serviço": 20.00, "App Gestão CPlug": 20.00, "Painel Senha Mobile": 49.00,
    "Controle de Mesas": 49.00, "Marketing": 24.50, "Relatório Dinâmico": 50.00,
    "Atualização em tempo real": 49.00, "Facilita NFE": 99.00, "Conciliação Bancária": 50.00,
    "Contratos de cartões e outros": 50.00, "Suporte Técnico - Estendido": 99.00, "Smart TEF": 49.90,
    "Smart Menu": 99.90, "Terminais Autoatendimento": 199.00, "Delivery Direto Profissional": 200.00,
    "Delivery Direto VIP": 300.00, "TEF": 99.90, "Cardápio digital": 99.00,
    "Backup Realtime": 199.90, "Business Intelligence (BI)": 199.00,
    "Programa de Fidelidade": 299.90, "Integração Tap": 299.00, "Integração API": 299.00,
    # Fixos/Contadores (preço zero aqui)
    "30 Notas Fiscais": 0.0, "Suporte Técnico - Via chamados": 0.0, "Relatório Básico": 0.0,
    "Notas Fiscais Ilimitadas": 0.0, "Painel Senha TV": 0.0, "Estoque em Grade": 0.0,
    "Relatórios": 0.0, "Suporte Técnico - Via chat": 0.0, "Delivery": 0.0, "Relatório KDS": 0.0,
    "PDV - Frente de Caixa": 0.0, "Usuários": 0.0,
}

# Custos MENSAIS Fixos por Item Extra
PRECO_EXTRA_USUARIO = 19.00
PRECO_EXTRA_PDV_GESTAO_PERFORMANCE = 59.90
PRECO_EXTRA_PDV_BLING = 40.00

# ---------------------------------------------------------
# Função para substituir placeholders - VERSÃO SIMPLES RESTAURADA
# ---------------------------------------------------------
def substituir_placeholders_no_slide(slide, dados):
    """Substitui placeholders (chaves do dicionário 'dados') pelo valor."""
    try:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    # Aplica substituições no texto do run atual
                    run_text_original = run.text
                    run_text_modificado = run_text_original
                    for k, v in dados.items():
                        placeholder = str(k) # Usa a chave como placeholder direto
                        if placeholder in run_text_modificado:
                            v_str = str(v) if v is not None else ""
                            run_text_modificado = run_text_modificado.replace(placeholder, v_str)
                    # Só atualiza o texto do run se ele foi modificado
                    if run_text_modificado != run_text_original:
                        run.text = run_text_modificado
    except Exception as e:
        print(f"ERRO em substituir_placeholders_no_slide: {e}")
        traceback.print_exc()

# ---------------------------------------------------------
# Classe PlanoFrame (Aba)
# ---------------------------------------------------------
class PlanoFrame(ttkb.Frame):
    # __init__ (Mantido como na versão anterior)
    def __init__( self, parent, aba_index, nome_cliente_var_shared, validade_proposta_var_shared, on_close_callback=None):
        super().__init__(parent)
        self.aba_index = aba_index; self.on_close_callback = on_close_callback
        self.nome_cliente_var = nome_cliente_var_shared; self.validade_proposta_var = validade_proposta_var_shared
        self.nome_plano_var = tk.StringVar(value="")
        self.current_plan = "PDV Básico"; self.spin_pdv_var = tk.IntVar(value=1); self.spin_users_var = tk.IntVar(value=1); self.spin_smart_tef_var = tk.IntVar(value=0)
        self.modules = { # Módulos atualizados
            "Usuários": tk.IntVar(), "30 Notas Fiscais": tk.IntVar(), "Suporte Técnico - Via chamados": tk.IntVar(), "Relatório Básico": tk.IntVar(), "PDV - Frente de Caixa": tk.IntVar(),
            "Notas Fiscais Ilimitadas": tk.IntVar(), "Importação de XML": tk.IntVar(), "Painel Senha TV": tk.IntVar(), "Estoque em Grade": tk.IntVar(), "Relatórios": tk.IntVar(),
            "Suporte Técnico - Via chat": tk.IntVar(), "Delivery": tk.IntVar(), "Relatório KDS": tk.IntVar(), "Produção": tk.IntVar(), "Promoções": tk.IntVar(), "Hub de Delivery": tk.IntVar(),
            "Ordem de Serviço": tk.IntVar(), "App Gestão CPlug": tk.IntVar(), "Painel Senha Mobile": tk.IntVar(), "Controle de Mesas": tk.IntVar(), "Marketing": tk.IntVar(),
            "Relatório Dinâmico": tk.IntVar(), "Atualização em tempo real": tk.IntVar(), "Facilita NFE": tk.IntVar(), "Conciliação Bancária": tk.IntVar(), "Contratos de cartões e outros": tk.IntVar(),
            "Suporte Técnico - Estendido": tk.IntVar(), "Smart TEF": tk.IntVar(), "Smart Menu": tk.IntVar(), "Terminais Autoatendimento": tk.IntVar(), "Delivery Direto Profissional": tk.IntVar(),
            "Delivery Direto VIP": tk.IntVar(), "TEF": tk.IntVar(), "Cardápio digital": tk.IntVar(), "Integração API": tk.IntVar(), "Business Intelligence (BI)": tk.IntVar(),
            "Backup Realtime": tk.IntVar(), "Programa de Fidelidade": tk.IntVar(), "Integração Tap": tk.IntVar(),
        }; self.check_buttons = {}
        self.user_override_anual_active = tk.BooleanVar(value=False); self.user_override_discount_active = tk.BooleanVar(value=False)
        self.valor_anual_editavel = tk.StringVar(value="0.00"); self.desconto_personalizado = tk.StringVar(value="10") # Padrão 10%
        self.computed_mensal_sem_fidelidade = 0.0; self.computed_mensal_efetivo_anual = 0.0; self.computed_anual_total = 0.0
        self.computed_desconto_percent = 0.0; self.computed_custo_adicional = 0.0
        # --- Layout ---
        self.canvas = tk.Canvas(self); self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar = ttkb.Scrollbar(self, orient="vertical", command=self.canvas.yview); self.scrollbar.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=self.scrollbar.set); self.container = ttkb.Frame(self.canvas)
        self.canvas.create_window((0,0), window=self.container, anchor="nw"); self.container.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.frame_main = ttkb.Frame(self.container); self.frame_main.pack(fill="both", expand=True)
        self.frame_left = ttkb.Frame(self.frame_main); self.frame_left.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        self.frame_right = ttkb.Frame(self.frame_main); self.frame_right.pack(side="left", fill="y", padx=5, pady=5)
        self._montar_layout_esquerda(); self._montar_layout_direita(); self.configurar_plano("PDV Básico")

    def fechar_aba(self):
        if self.on_close_callback: self.on_close_callback(self.aba_index)

    def on_bling_selected(self, event=None):
        selected_bling_plan = self.bling_var.get();
        if selected_bling_plan in LISTA_PLANOS_BLING: self.configurar_plano(selected_bling_plan)
        self.bling_var.set("Selecionar Bling...")

    def _montar_layout_esquerda(self): # (Layout mantido)
        top_bar = ttkb.Frame(self.frame_left); top_bar.pack(fill="x", pady=5)
        ttkb.Label(top_bar, text=f"Aba Plano {self.aba_index}", font="-size 12 -weight bold").pack(side="left")
        btn_close = ttkb.Button(top_bar, text="X", command=self.fechar_aba, bootstyle="danger-outline", width=3); btn_close.pack(side="right")
        frame_planos = ttkb.Labelframe(self.frame_left, text="Selecionar Plano Base"); frame_planos.pack(fill="x", pady=5)
        self.bling_combobox = None; plan_buttons_frame = ttkb.Frame(frame_planos); plan_buttons_frame.pack(fill="x")
        for i, p in enumerate(LISTA_PLANOS_UI):
            if p == "Bling":
                self.bling_var = tk.StringVar(value="Selecionar Bling...")
                self.bling_combobox = ttk.Combobox(plan_buttons_frame, textvariable=self.bling_var, values=LISTA_PLANOS_BLING, state="readonly", width=25)
                self.bling_combobox.grid(row=0, column=i, padx=5, pady=5, sticky="ew"); self.bling_combobox.bind("<<ComboboxSelected>>", self.on_bling_selected)
            else:
                btn = ttkb.Button(plan_buttons_frame, text=p, width=15, command=lambda pl=p: self.configurar_plano(pl))
                btn.grid(row=0, column=i, padx=5, pady=5, sticky="ew")
            plan_buttons_frame.grid_columnconfigure(i, weight=1)
        frame_mod = ttkb.Labelframe(self.frame_left, text="Módulos Opcionais"); frame_mod.pack(fill="both", expand=True, pady=5)
        f_mod_cols = ttkb.Frame(frame_mod); f_mod_cols.pack(fill="both", expand=True)
        f_mod_left = ttkb.Frame(f_mod_cols); f_mod_left.pack(side="left", fill="both", expand=True, padx=5)
        f_mod_right = ttkb.Frame(f_mod_cols); f_mod_right.pack(side="left", fill="both", expand=True, padx=5)
        spinbox_items = {"PDV - Frente de Caixa", "Usuários", "Smart TEF"}; always_fixed = {"Relatórios", "30 Notas Fiscais", "Notas Fiscais Ilimitadas", "Suporte Técnico - Via chamados", "Relatório Básico"}
        displayable_mods = sorted([m for m in self.modules.keys() if m not in spinbox_items and m not in always_fixed])
        mid = len(displayable_mods)//2; left_side = displayable_mods[:mid]; right_side = displayable_mods[mid:]
        self.check_buttons = {}
        for m in left_side:
             if m in self.modules: cb = ttk.Checkbutton(f_mod_left, text=m, variable=self.modules[m], command=self.atualizar_valores); cb.pack(anchor="w", pady=2); self.check_buttons[m] = cb
        for m in right_side:
             if m in self.modules: cb = ttk.Checkbutton(f_mod_right, text=m, variable=self.modules[m], command=self.atualizar_valores); cb.pack(anchor="w", pady=2); self.check_buttons[m] = cb
        frame_dados = ttkb.Labelframe(self.frame_left, text="Dados do Cliente e Proposta"); frame_dados.pack(fill="x", pady=5)
        ttkb.Label(frame_dados, text="Nome do Cliente:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttkb.Entry(frame_dados, textvariable=self.nome_cliente_var, width=30).grid(row=0, column=1, padx=5, pady=2)
        ttkb.Label(frame_dados, text="Validade Proposta:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttkb.Entry(frame_dados, textvariable=self.validade_proposta_var, width=15).grid(row=1, column=1, padx=5, pady=2, sticky="w")
        ttkb.Label(frame_dados, text="Nome do Plano (Opcional):").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        ttkb.Entry(frame_dados, textvariable=self.nome_plano_var, width=30).grid(row=2, column=1, padx=5, pady=2)

    def _montar_layout_direita(self): # (Layout mantido)
        frame_inc = ttkb.Labelframe(self.frame_right, text="Quantidades"); frame_inc.pack(fill="x", pady=5)
        ttkb.Label(frame_inc, text="PDVs - Frente de Caixa").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.sp_pdv = ttkb.Spinbox(frame_inc, from_=0, to=99, textvariable=self.spin_pdv_var, width=5, command=self.atualizar_valores); self.sp_pdv.grid(row=0, column=1, padx=5, pady=2)
        ttkb.Label(frame_inc, text="Usuários").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.sp_usr = ttkb.Spinbox(frame_inc, from_=0, to=999, textvariable=self.spin_users_var, width=5, command=self.atualizar_valores); self.sp_usr.grid(row=1, column=1, padx=5, pady=2)
        ttkb.Label(frame_inc, text="Smart TEF").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.sp_smf = ttkb.Spinbox(frame_inc, from_=0, to=99, textvariable=self.spin_smart_tef_var, width=5, command=self.atualizar_valores); self.sp_smf.grid(row=2, column=1, padx=5, pady=2)
        frame_valores = ttkb.Labelframe(self.frame_right, text="Valores da Proposta"); frame_valores.pack(fill="both", pady=5, expand=True)
        self.lbl_plano_mensal_sem_fid = ttkb.Label(frame_valores, text="Mensal (Sem Fidelidade): R$ 0,00", font="-size 11"); self.lbl_plano_mensal_sem_fid.pack(pady=(5, 2), anchor="w", padx=5)
        self.lbl_treinamento = ttkb.Label(frame_valores, text="+ Custo Treinamento: R$ 0,00", font="-size 9"); self.lbl_treinamento.pack(pady=(0, 5), anchor="w", padx=15)
        self.lbl_plano_mensal_no_anual = ttkb.Label(frame_valores, text="Mensal (no Plano Anual): R$ 0,00", font="-size 12 -weight bold"); self.lbl_plano_mensal_no_anual.pack(pady=5, anchor="w", padx=5)
        self.lbl_plano_anual_total = ttkb.Label(frame_valores, text="Anual (Pagamento Único): R$ 0,00", font="-size 12 -weight bold"); self.lbl_plano_anual_total.pack(pady=5, anchor="w", padx=5)
        self.lbl_desconto = ttkb.Label(frame_valores, text="Desconto Anual Aplicado: 10%", font="-size 9"); self.lbl_desconto.pack(pady=(5, 10), anchor="w", padx=5)
        ttk.Separator(frame_valores, orient='horizontal').pack(fill='x', pady=5, padx=5)
        frame_edicao = ttkb.Frame(frame_valores); frame_edicao.pack(fill="x", pady=5)
        frame_edit_anual = ttkb.Labelframe(frame_edicao, text="Editar Anual Total (R$)"); frame_edit_anual.pack(side="left", padx=5, fill="x", expand=True)
        e_anual = ttkb.Entry(frame_edit_anual, textvariable=self.valor_anual_editavel, width=10); e_anual.pack(side="left", padx=5, pady=2); e_anual.bind("<KeyRelease>", self.on_user_edit_valor_anual)
        b_reset_anual = ttkb.Button(frame_edit_anual, text="Reset", command=self.on_reset_anual, width=5, bootstyle="warning-outline"); b_reset_anual.pack(side="left", padx=5, pady=2)
        frame_edit_desc = ttkb.Labelframe(frame_edicao, text="Editar Desconto (%)"); frame_edit_desc.pack(side="left", padx=5, fill="x", expand=True)
        e_desc = ttkb.Entry(frame_edit_desc, textvariable=self.desconto_personalizado, width=5); e_desc.pack(side="left", padx=5, pady=2); e_desc.bind("<KeyRelease>", self.on_user_edit_desconto)
        b_reset_desc = ttkb.Button(frame_edit_desc, text="Reset", command=self.on_reset_desconto, width=5, bootstyle="warning-outline"); b_reset_desc.pack(side="left", padx=5, pady=2)

    def on_user_edit_valor_anual(self, *args): self.user_override_anual_active.set(True); self.user_override_discount_active.set(False); self.atualizar_valores()
    def on_reset_anual(self): self.user_override_anual_active.set(False); self.atualizar_valores() # Recalcula com desconto padrão
    def on_user_edit_desconto(self, *args): self.user_override_discount_active.set(True); self.user_override_anual_active.set(False); self.atualizar_valores()
    def on_reset_desconto(self): self.user_override_discount_active.set(False); self.desconto_personalizado.set("10"); self.atualizar_valores() # Volta para 10%

    def configurar_plano(self, plano): # (Lógica mantida)
        if not plano.startswith("Bling -") and self.bling_combobox: self.bling_var.set("Selecionar Bling...")
        if plano not in PLAN_INFO: showerror("Erro", f"Plano '{plano}' não encontrado."); return
        info = PLAN_INFO[plano]; self.current_plan = plano
        min_pdv = info.get("min_pdv", 0); max_pdv = min_pdv + info.get("max_extra_pdvs", 99)
        min_users = info.get("min_users", 0); max_users = min_users + info.get("max_extra_users", 999)
        self.spin_pdv_var.set(min_pdv); self.sp_pdv.config(from_=min_pdv, to=max_pdv)
        self.spin_users_var.set(min_users); self.sp_usr.config(from_=min_users, to=max_users)
        min_smart_tef = 0; max_smart_tef = 99
        if plano == "Gestão": max_smart_tef = 3
        elif plano == "Performance": min_smart_tef = 3; max_smart_tef = 3
        self.spin_smart_tef_var.set(min_smart_tef); self.sp_smf.config(from_=min_smart_tef, to=max_smart_tef)
        for m, var in self.modules.items(): var.set(0)
        allowed = info.get("allowed_optionals", []); mandatory = info.get("mandatory", [])
        for m in mandatory:
            if m in self.modules: self.modules[m].set(1)
        for m, cb in self.check_buttons.items():
            is_mandatory = m in mandatory; is_allowed_optional = m in allowed
            if is_mandatory: cb.config(state='disabled')
            elif is_allowed_optional: cb.config(state='normal')
            else: cb.config(state='disabled')
        self.user_override_anual_active.set(False); self.user_override_discount_active.set(False)
        self.desconto_personalizado.set("10"); self.atualizar_valores()

    def _calcular_extras(self): # (Lógica mantida)
        total_extras_cost = 0.0; total_extras_descontavel = 0.0; total_extras_nao_descontavel = 0.0
        info = PLAN_INFO[self.current_plan]; mandatory = info.get("mandatory", [])
        pdv_atuais = self.spin_pdv_var.get(); pdv_incluidos = info.get("min_pdv", 0)
        pdv_extras = max(0, pdv_atuais - pdv_incluidos); pdv_price = 0.0
        if self.current_plan in ["Gestão", "Performance"]: pdv_price = PRECO_EXTRA_PDV_GESTAO_PERFORMANCE
        elif self.current_plan.startswith("Bling"): pdv_price = PRECO_EXTRA_PDV_BLING
        cost_pdv_extra = pdv_extras * pdv_price; total_extras_cost += cost_pdv_extra
        if "PDV Extra" not in SEM_DESCONTO: total_extras_descontavel += cost_pdv_extra
        else: total_extras_nao_descontavel += cost_pdv_extra
        users_atuais = self.spin_users_var.get(); users_incluidos = info.get("min_users", 0)
        users_extras = max(0, users_atuais - users_incluidos); cost_users_extra = users_extras * PRECO_EXTRA_USUARIO
        total_extras_cost += cost_users_extra
        if "User Extra" not in SEM_DESCONTO: total_extras_descontavel += cost_users_extra
        else: total_extras_nao_descontavel += cost_users_extra
        for m, var_m in self.modules.items():
            if m in self.check_buttons and var_m.get() == 1 and m not in mandatory:
                 price = precos_mensais.get(m, 0.0); total_extras_cost += price
                 if m not in SEM_DESCONTO: total_extras_descontavel += price
                 else: total_extras_nao_descontavel += price
        if self.current_plan == "Gestão":
            smart_tef_atuais = self.spin_smart_tef_var.get(); smart_tef_extras = max(0, smart_tef_atuais - 0)
            if smart_tef_extras > 0:
                 price = smart_tef_extras * precos_mensais.get("Smart TEF", 0.0); total_extras_cost += price
                 if "Smart TEF" not in SEM_DESCONTO: total_extras_descontavel += price
                 else: total_extras_nao_descontavel += price
        return total_extras_cost, total_extras_descontavel, total_extras_nao_descontavel

    def atualizar_valores(self, *args): # (Lógica mantida)
        try:
            if not self.current_plan or self.current_plan not in PLAN_INFO: return
            info = PLAN_INFO[self.current_plan]; is_bling = self.current_plan.startswith("Bling"); is_auto = self.current_plan == "Autoatendimento"; is_branco = self.current_plan == "Em Branco"
            base_mensal_efetivo_anual = info.get("base_mensal", 0.0)
            total_extras_cost, total_extras_descontavel, total_extras_nao_descontavel = self._calcular_extras()
            base_mensal_sem_fidelidade = (base_mensal_efetivo_anual / 0.90) if base_mensal_efetivo_anual > 0 else 0.0
            total_mensal_sem_fidelidade = base_mensal_sem_fidelidade + total_extras_cost
            total_mensal_efetivo_anual_base_calc = base_mensal_efetivo_anual + total_extras_cost
            final_mensal_efetivo_anual = 0.0; final_anual_total = 0.0; desconto_aplicado_percent = 10.0
            if self.user_override_anual_active.get():
                try:
                    edited_total_anual = float(self.valor_anual_editavel.get())
                    final_anual_total = edited_total_anual; final_mensal_efetivo_anual = final_anual_total / 12.0
                    if total_mensal_sem_fidelidade > 0: desconto_aplicado_percent = ((total_mensal_sem_fidelidade - final_mensal_efetivo_anual) / total_mensal_sem_fidelidade) * 100
                    else: desconto_aplicado_percent = 0.0
                    self.desconto_personalizado.set(str(round(max(0, desconto_aplicado_percent))))
                except ValueError: final_mensal_efetivo_anual = total_mensal_efetivo_anual_base_calc; final_anual_total = final_mensal_efetivo_anual * 12.0; desconto_aplicado_percent = 10.0; self.desconto_personalizado.set("10"); self.valor_anual_editavel.set(f"{final_anual_total:.2f}")
            elif self.user_override_discount_active.get():
                try:
                    desc_custom = float(self.desconto_personalizado.get()); desc_dec = desc_custom / 100.0; desconto_aplicado_percent = desc_custom
                    base_sem_fid_mais_extras_descont = base_mensal_sem_fidelidade + total_extras_descontavel
                    final_mensal_efetivo_anual = (base_sem_fid_mais_extras_descont * (1 - desc_dec)) + total_extras_nao_descontavel
                    final_anual_total = final_mensal_efetivo_anual * 12.0
                    self.valor_anual_editavel.set(f"{final_anual_total:.2f}")
                except ValueError: final_mensal_efetivo_anual = total_mensal_efetivo_anual_base_calc; final_anual_total = final_mensal_efetivo_anual * 12.0; desconto_aplicado_percent = 10.0; self.desconto_personalizado.set("10"); self.valor_anual_editavel.set(f"{final_anual_total:.2f}")
            else:
                final_mensal_efetivo_anual = total_mensal_efetivo_anual_base_calc; final_anual_total = final_mensal_efetivo_anual * 12.0
                if total_mensal_sem_fidelidade > 0: desconto_aplicado_percent = ((total_mensal_sem_fidelidade - final_mensal_efetivo_anual) / total_mensal_sem_fidelidade) * 100
                else: desconto_aplicado_percent = 0.0
                self.valor_anual_editavel.set(f"{final_anual_total:.2f}"); self.desconto_personalizado.set(str(round(max(0,desconto_aplicado_percent))))
            custo_adicional = 0.0; label_custo = "Treinamento"
            if is_bling: label_custo = "Implementação"
            if not is_bling and not is_auto and not is_branco:
                limite_custo = 549.90
                if total_mensal_sem_fidelidade > 0 and total_mensal_sem_fidelidade < limite_custo: custo_adicional = limite_custo - total_mensal_sem_fidelidade
            mensal_sem_fid_str = f"{total_mensal_sem_fidelidade:.2f}".replace(".", ","); mensal_no_anual_str = f"{final_mensal_efetivo_anual:.2f}".replace(".", ","); anual_total_str = f"{final_anual_total:.2f}".replace(".", ","); custo_adic_str = f"{custo_adicional:.2f}".replace(".", ",")
            desconto_final_percent = round(max(0, desconto_aplicado_percent))
            self.lbl_plano_mensal_sem_fid.config(text=f"Mensal (Sem Fidelidade): R$ {mensal_sem_fid_str}")
            if custo_adicional > 0.01: self.lbl_treinamento.config(text=f"+ Custo {label_custo}: R$ {custo_adic_str}"); self.lbl_treinamento.pack(pady=(0, 5), anchor="w", padx=15)
            else: self.lbl_treinamento.pack_forget()
            self.lbl_plano_mensal_no_anual.config(text=f"Mensal (no Plano Anual): R$ {mensal_no_anual_str}")
            self.lbl_plano_anual_total.config(text=f"Anual (Pagamento Único): R$ {anual_total_str}")
            self.lbl_desconto.config(text=f"Desconto Anual Aplicado: {desconto_final_percent}%")
            self.computed_mensal_sem_fidelidade = total_mensal_sem_fidelidade; self.computed_mensal_efetivo_anual = final_mensal_efetivo_anual; self.computed_anual_total = final_anual_total
            self.computed_desconto_percent = desconto_final_percent; self.computed_custo_adicional = custo_adicional
        except Exception as e: print(f"ERRO em atualizar_valores: {e}"); traceback.print_exc()

    # --- montar_lista_modulos CORRIGIDO ---
    def montar_lista_modulos(self):
        linhas = []
        info = PLAN_INFO.get(self.current_plan, {})
        mandatory = info.get("mandatory", [])

        # PDVs
        pdv_val = self.spin_pdv_var.get()
        if pdv_val > 0:
            linhas.append(f"{pdv_val}x PDV - Frente de Caixa")

        # Usuários
        usr_val = self.spin_users_var.get()
        if usr_val > 0:
            linhas.append(f"{usr_val}x Usuários")

        # Smart TEF
        smart_tef_val = self.spin_smart_tef_var.get()
        if smart_tef_val > 0:
            linhas.append(f"{smart_tef_val}x Smart TEF")
            # Limite visual removido para simplificar, já é controlado pelo spinbox
            # if self.current_plan == "Gestão":
            #     linhas[-1] += " (Limite: 3)"

        # Módulos Mandatórios (não-spinbox)
        for m in mandatory:
            if m not in ["PDV - Frente de Caixa", "Usuários", "Smart TEF"]: # Já tratados
                linhas.append(f"1x {m}")

        # Módulos Opcionais (Checkboxes selecionados)
        for m, var_m in self.modules.items():
            # Adiciona somente se for um checkbox gerenciado E estiver marcado E não for mandatório
            if m in self.check_buttons and var_m.get() == 1 and m not in mandatory:
                linhas.append(f"1x {m}")

        # Remove duplicados e formata
        unique_mods = []
        [unique_mods.append(mod) for mod in linhas if mod not in unique_mods]
        montagem = "\n".join(f"•    {m}" for m in unique_mods)
        return montagem

    def gerar_dados_proposta(self, nome_closer, cel_closer, email_closer): # (Lógica mantida, chaves ajustadas)
        nome_plano_selecionado = self.current_plan; nome_plano_editado = self.nome_plano_var.get().strip()
        nome_plano_final = nome_plano_editado if nome_plano_editado else nome_plano_selecionado
        mensal_sem_fid_val = self.computed_mensal_sem_fidelidade; mensal_efetivo_anual_val = self.computed_mensal_efetivo_anual
        anual_total_val = self.computed_anual_total; custo_adicional_val = self.computed_custo_adicional; desconto_percent_val = self.computed_desconto_percent
        mensal_sem_fid_str = f"R$ {mensal_sem_fid_val:.2f}".replace(".", ","); mensal_efetivo_anual_str = f"R$ {mensal_efetivo_anual_val:.2f}".replace(".", ",")
        anual_total_str = f"R$ {anual_total_val:.2f}".replace(".", ","); custo_adicional_str = f"R$ {custo_adicional_val:.2f}".replace(".", ",")
        label_custo = "Treinamento" if not self.current_plan.startswith("Bling") else "Implementação"
        plano_mensal_display = mensal_sem_fid_str
        if custo_adicional_val > 0.01: plano_mensal_display += f" + {custo_adicional_str} ({label_custo})"
        tipo_suporte = "Regular"; horario_suporte = "09:00 às 17:00 Seg-Sex"
        suporte_chat = self.modules.get("Suporte Técnico - Via chat", tk.IntVar()).get() == 1
        suporte_estendido = self.modules.get("Suporte Técnico - Estendido", tk.IntVar()).get() == 1
        if suporte_estendido: tipo_suporte = "Estendido"; horario_suporte = "09-22h Seg-Sex & 11-21h Sab-Dom"
        elif suporte_chat: tipo_suporte = "Chat Incluso"; horario_suporte = "09-22h Seg-Sex & 11-21h Sab-Dom"
        montagem = self.montar_lista_modulos()
        economia_str = ""; custo_total_mensalizado = (mensal_sem_fid_val * 12) + custo_adicional_val; custo_total_anualizado = anual_total_val
        if not self.current_plan == "Autoatendimento":
             if custo_total_mensalizado > custo_total_anualizado + 0.01:
                  economia_val = custo_total_mensalizado - custo_total_anualizado; econ_str = f"{economia_val:.2f}".replace(".", ",")
                  economia_str = f"Economia de R$ {econ_str} no plano anual"
        # CHAVES PARA SUBSTITUIÇÃO NO PPTX (Devem coincidir com placeholders)
        dados = {
            "montagem_do_plano": montagem, "plano_mensal": plano_mensal_display, "plano_anual": mensal_efetivo_anual_str,
            "plano_anual_total_pagamento": anual_total_str, "custo_treinamento_valor": custo_adicional_str if custo_adicional_val > 0.01 else "Incluso",
            "desconto_total": f"{desconto_percent_val}%", "nome_do_plano": nome_plano_final, "tipo_de_suporte": tipo_suporte,
            "horario_de_suporte": horario_suporte, "validade_proposta": self.validade_proposta_var.get(), "nome_closer": nome_closer,
            "celular_closer": cel_closer, "email_closer": email_closer, "nome_cliente": self.nome_cliente_var.get(), "economia_anual": economia_str
        }
        return dados


# --- Função gerar_proposta (Usa substituição simples) ---
def gerar_proposta(lista_abas, nome_closer, celular_closer, email_closer):
    ppt_file = "Proposta Comercial ConnectPlug.pptx"
    if not os.path.exists(ppt_file): showerror("Erro", f"Arquivo template '{ppt_file}' não encontrado!"); return None
    try: prs = Presentation(ppt_file)
    except Exception as e: showerror("Erro", f"Falha ao abrir '{ppt_file}': {e}"); return None
    if not lista_abas: showerror("Erro", "Não há abas para gerar Proposta."); return None

    primeira_aba = lista_abas[0]
    dados_proposta = primeira_aba.gerar_dados_proposta(nome_closer, celular_closer, email_closer)

    # Aplica substituição em todos os slides
    for slide in prs.slides:
        substituir_placeholders_no_slide(slide, dados_proposta) # Usa a função simples

    # Salva o arquivo
    nome_cliente_safe = dados_proposta.get("nome_cliente", "SemNome").replace("/", "-").replace("\\", "-")
    hoje_str = date.today().strftime("%d-%m-%Y")
    nome_arquivo = f"Proposta ConnectPlug - {nome_cliente_safe} - {hoje_str}.pptx"
    try:
        prs.save(nome_arquivo); showinfo("Sucesso", f"Proposta gerada: {nome_arquivo}"); return nome_arquivo
    except Exception as e: showerror("Erro", f"Falha ao salvar '{nome_arquivo}':\n{e}"); return None

# --- Função gerar_material (Usa MAPEAMENTO e lógica de exclusão) ---
def gerar_material(lista_abas, nome_closer, celular_closer, email_closer):
    mat_file = "Material Tecnico ConnectPlug.pptx"
    if not os.path.exists(mat_file): showerror("Erro", f"Arquivo '{mat_file}' não encontrado!"); return None
    try: prs = Presentation(mat_file)
    except Exception as e: showerror("Erro", f"Falha ao abrir '{mat_file}': {e}"); return None
    if not lista_abas: showerror("Erro", "Não há abas para gerar Material."); return None

    # 1) Coleta Módulos Ativos e Planos (Nomes Atuais)
    modulos_ativos_geral = set(); planos_usados_geral = set()
    dados_primeira_aba = lista_abas[0].gerar_dados_proposta(nome_closer, celular_closer, email_closer)
    for aba in lista_abas:
        current_plan_name = aba.current_plan; planos_usados_geral.add(current_plan_name)
        info_aba = PLAN_INFO.get(current_plan_name, {}); mandatory_aba = info_aba.get("mandatory", [])
        for mod in mandatory_aba: modulos_ativos_geral.add(mod)
        for nome_mod, var_mod in aba.modules.items():
            if var_mod.get() == 1 and nome_mod not in mandatory_aba: modulos_ativos_geral.add(nome_mod)
        if aba.spin_pdv_var.get() > 0: modulos_ativos_geral.add("PDV - Frente de Caixa")
        if aba.spin_users_var.get() > 0: modulos_ativos_geral.add("Usuários")
        if aba.spin_smart_tef_var.get() > 0: modulos_ativos_geral.add("Smart TEF")
        # Adiciona outros itens quantificáveis ou checkboxes se relevantes para exclusão
        if aba.modules.get("TEF", tk.IntVar()).get() == 1: modulos_ativos_geral.add("TEF")
        if aba.modules.get("Terminais Autoatendimento", tk.IntVar()).get() == 1: modulos_ativos_geral.add("Terminais Autoatendimento")

    # print(f"DEBUG: Módulos Ativos para Material: {modulos_ativos_geral}") # Descomente para depurar

    # 2) MAPEAMENTO_MODULOS: Placeholder no PPTX -> Nome Interno ATUAL no Python
    #    *** REVISE ESTE MAPEAMENTO CUIDADOSAMENTE ***
    MAPEAMENTO_MODULOS = {
        "slide_sempre": None,
        "check_sistema_kds": "Relatório KDS", "check_Hub_de_Delivery": "Hub de Delivery",
        "check_integracao_api": "Integração API", "check_integracao_tap": "Integração Tap",
        "check_controle_de_mesas": "Controle de Mesas", "check_Delivery": "Delivery",
        "check_producao": "Produção", "check_Estoque_em_Grade": "Estoque em Grade",
        "check_Facilita_NFE": "Facilita NFE", "check_Importacao_de_xml": "Importação de XML", # Atenção ao Case
        "check_conciliacao_bancaria": "Conciliação Bancária", "check_contratos_de_cartoes": "Contratos de cartões e outros",
        "check_ordem_de_servico": "Ordem de Serviço", "check_relatorio_dinamico": "Relatório Dinâmico",
        "check_programa_de_fidelidade": "Programa de Fidelidade", "check_business_intelligence": "Business Intelligence (BI)",
        "check_smartmenu": "Smart Menu", "check_backup_real_time": "Backup Realtime",
        "check_att_tempo_real": "Atualização em tempo real", "check_promocao": "Promoções",
        "check_marketing": "Marketing", "pdv_balcao": "PDV - Frente de Caixa", # Ligado ao Spinbox
        "qtd_smarttef": "Smart TEF", # Ligado ao Spinbox
        "qtd_tef": "TEF", # Ligado ao Checkbox TEF
        "qtd_autoatendimento": "Terminais Autoatendimento", # Ligado ao Checkbox Terminais...
        "qtd_cardapio_digital": "Cardápio digital", # Ligado ao Checkbox Cardápio...
        "qtd_app_gestao_cplug": "App Gestão CPlug",
        "check_delivery_direto_vip": "Delivery Direto VIP", "check_delivery_direto_profissional": "Delivery Direto Profissional",
        "placeholder_painel_senha_tv": "Painel Senha TV", # Placeholder Novo?
        "placeholder_painel_senha_mobile": "Painel Senha Mobile", # Placeholder Novo?
        "placeholder_suporte_chat": "Suporte Técnico - Via chat", # Placeholder Novo?
        "placeholder_suporte_estendido": "Suporte Técnico - Estendido", # Placeholder Novo?
        # Placeholder genérico para qualquer NF ativa
        "check_notas_fiscais": {"Notas Fiscais Ilimitadas", "30 Notas Fiscais"},
    }

    # 3) Decide quais slides manter (Lógica Mantida)
    keep_slides = set()
    for i, slide in enumerate(prs.slides):
        slide_mantido = False
        try:
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
                             # Condição Bling (revisar se ainda necessária)
                             if "slide_bling" in txt_run and any(p.startswith("Bling -") for p in planos_usados_geral): slide_mantido = True; break
        except Exception as slide_err: print(f"Aviso: Erro ao processar slide {i}: {slide_err}")
        if slide_mantido: keep_slides.add(i)

    # 4) Remove slides não mantidos
    if not keep_slides: showerror("Erro", "Nenhum slide selecionado para Material. Verifique Mapeamento."); return None
    # print(f"DEBUG: Slides a manter no Material: {sorted(list(keep_slides))}") # Descomente para depurar
    slides_removidos = 0
    for idx in reversed(range(len(prs.slides))):
        if idx not in keep_slides:
            try:
                rId = prs.slides._sldIdLst[idx].rId; prs.part.drop_rel(rId); del prs.slides._sldIdLst[idx]; slides_removidos += 1
            except Exception as e: print(f"Aviso: Falha ao remover slide {idx}. Erro: {e}")
    print(f"Slides removidos do Material: {slides_removidos} (Restantes: {len(prs.slides)})")

    # 5) Substituir placeholders globais
    for slide in prs.slides:
        substituir_placeholders_no_slide(slide, dados_primeira_aba)

    # 6) Salvar
    nome_cliente_safe = dados_primeira_aba.get("nome_cliente", "SemNome").replace("/", "-").replace("\\", "-")
    hoje_str = date.today().strftime("%d-%m-%Y")
    nome_arquivo = f"Material Tecnico ConnectPlug - {nome_cliente_safe} - {hoje_str}.pptx"
    try:
        prs.save(nome_arquivo); showinfo("Sucesso", f"Material Técnico gerado: {nome_arquivo}"); return nome_arquivo
    except Exception as e: showerror("Erro", f"Falha ao salvar '{nome_arquivo}':\n{e}"); return None


# --- Google Drive / Auth / Upload (Código Mantido) ---
SCOPES = ['https://www.googleapis.com/auth/drive']
def baixar_client_secret_remoto():
    url = "https://github.com/DevRGS/Gerador/raw/refs/heads/main/config/client_secret_788265418970-ur6f189oqvsttseeg6g77fegt0su67dj.apps.googleusercontent.com.json"; nome_local = "client_secret_temp.json"
    if not os.path.exists(nome_local):
        print(f"Baixando {nome_local}...")
        try: r = requests.get(url, timeout=15); r.raise_for_status(); open(nome_local, "w", encoding="utf-8").write(r.text)
        except requests.exceptions.RequestException as e: raise ConnectionError(f"Erro ao baixar {nome_local}: {e}") from e
    return nome_local
def get_gdrive_service():
    creds = None; token_file = 'token.json'; client_secret_file = None
    try: client_secret_file = baixar_client_secret_remoto()
    except ConnectionError as e: showerror("Erro Crítico", f"Falha ao obter credenciais do Google Drive:\n{e}"); return None
    except Exception as e: showerror("Erro Crítico", f"Erro inesperado ao obter credenciais:\n{e}"); return None
    if os.path.exists(token_file):
        try: creds = pickle.load(open(token_file, 'rb'))
        except Exception: creds = None; os.remove(token_file) if os.path.exists(token_file) else None
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try: creds.refresh(Request())
            except Exception as e: print(f"Erro refresh token: {e}"); creds = None; os.remove(token_file) if os.path.exists(token_file) else None
        if not creds:
             try: flow = InstalledAppFlow.from_client_secrets_file(client_secret_file, SCOPES); creds = flow.run_local_server(port=0)
             except Exception as e: showerror("Erro de Autenticação", f"Falha ao autenticar com Google:\n{e}"); return None
        try: pickle.dump(creds, open(token_file, 'wb'))
        except Exception as e: print(f"Aviso: Falha ao salvar token: {e}")
    try: return build('drive', 'v3', credentials=creds)
    except Exception as e: showerror("Erro Google API", f"Falha build serviço Google: {e}"); return None
def upload_pptx_and_export_to_pdf(local_pptx_path):
    if not os.path.exists(local_pptx_path): showerror("Erro", f"Arquivo '{local_pptx_path}' não encontrado."); return
    service = get_gdrive_service();
    if not service: showerror("Erro Google Drive", "Não foi possível conectar."); return
    pdf_output_name = local_pptx_path.replace(".pptx", ".pdf"); base_name = os.path.basename(local_pptx_path); file_id = None
    try:
        print(f"Upload de '{base_name}'..."); file_metadata = {'name': base_name, 'mimeType': 'application/vnd.google-apps.presentation'}; media = MediaFileUpload(local_pptx_path, mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation', resumable=True)
        uploaded_file = service.files().create(body=file_metadata, media_body=media, fields='id').execute(); file_id = uploaded_file.get('id')
        if not file_id: raise Exception("Upload falhou (sem ID)."); print(f"Upload OK. ID: {file_id}. Exportando PDF...")
        request = service.files().export_media(fileId=file_id, mimeType='application/pdf'); fh = io.BytesIO(); downloader = MediaIoBaseDownload(fh, request); done = False
        while not done: status, done = downloader.next_chunk()
        open(pdf_output_name, 'wb').write(fh.getvalue()); print(f"PDF gerado: '{pdf_output_name}'."); showinfo("Google Drive", f"PDF gerado:\n'{pdf_output_name}'.")
    except Exception as e: showerror("Erro Google Drive", f"Erro upload/conversão:\n{e}")
    finally:
        if file_id:
            try: print(f"Deletando temp (ID: {file_id})..."); service.files().delete(fileId=file_id).execute(); print("Temp deletado.")
            except Exception as del_err: print(f"Aviso: Falha ao deletar temp: {del_err}")

# --- MainApp (Código Mantido) ---
class MainApp(ttkb.Window):
    def __init__(self):
        super().__init__(themename="litera"); self.title("Gerador de Propostas ConnectPlug v2.3"); self.geometry("1200x800")
        self.nome_closer_var = tk.StringVar(); self.celular_closer_var = tk.StringVar(); self.email_closer_var = tk.StringVar()
        self.nome_cliente_var_shared = tk.StringVar(value=""); self.validade_proposta_var_shared = tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
        carregar_config(self.nome_closer_var, self.celular_closer_var, self.email_closer_var); self.protocol("WM_DELETE_WINDOW", self.on_close)
        top_bar = ttkb.Frame(self); top_bar.pack(side="top", fill="x", pady=5, padx=5)
        ttkb.Label(top_bar, text="Vendedor:").pack(side="left", padx=(0, 2)); ttkb.Entry(top_bar, textvariable=self.nome_closer_var, width=20).pack(side="left", padx=(0, 5))
        ttkb.Label(top_bar, text="Celular:").pack(side="left", padx=(0, 2)); ttkb.Entry(top_bar, textvariable=self.celular_closer_var, width=15).pack(side="left", padx=(0, 5))
        ttkb.Label(top_bar, text="Email:").pack(side="left", padx=(0, 2)); ttkb.Entry(top_bar, textvariable=self.email_closer_var, width=25).pack(side="left", padx=(0, 10))
        self.btn_add = ttkb.Button(top_bar, text="+ Nova Aba", command=self.add_aba, bootstyle="success"); self.btn_add.pack(side="right", padx=5)
        self.notebook = ttkb.Notebook(self); self.notebook.pack(fill="both", expand=True, padx=5, pady=(0, 5))
        bot_frame = ttkb.Frame(self); bot_frame.pack(side="bottom", fill="x", pady=5, padx=5)
        ttkb.Button(bot_frame, text="Gerar Proposta + PDF", command=self.on_gerar_proposta, bootstyle="primary").pack(side="left", padx=5)
        ttkb.Button(bot_frame, text="Gerar Material + PDF", command=self.on_gerar_mat_tecnico, bootstyle="info").pack(side="left", padx=5)
        ttkb.Button(bot_frame, text="Gerar TUDO + PDF", command=self.on_gerar_tudo, bootstyle="secondary").pack(side="left", padx=5)
        self.abas_criadas = {}; self.ultimo_indice = 0; self.add_aba()
        try:
            baixar_arquivo_if_needed("Proposta Comercial ConnectPlug.pptx", "https://github.com/DevRGS/Gerador/raw/refs/heads/main/assets/Proposta%20Comercial%20ConnectPlug.pptx")
            baixar_arquivo_if_needed("Material Tecnico ConnectPlug.pptx", "https://github.com/DevRGS/Gerador/raw/refs/heads/main/assets/Material%20Tecnico%20ConnectPlug.pptx")
        except ConnectionError: pass # Erro já foi mostrado ao usuário
        except Exception as e: showerror("Erro Templates", f"Erro inesperado ao baixar arquivos base:\n{e}")
    def on_close(self): salvar_config(self.nome_closer_var.get(), self.celular_closer_var.get(), self.email_closer_var.get()); self.destroy()
    def add_aba(self):
        if len(self.abas_criadas) >= MAX_ABAS: showinfo("Limite Atingido", f"Máximo de {MAX_ABAS} abas."); return; self.ultimo_indice += 1; idx = self.ultimo_indice
        frame_aba = PlanoFrame(self.notebook, idx, self.nome_cliente_var_shared, self.validade_proposta_var_shared, self.fechar_aba)
        self.notebook.add(frame_aba, text=f"Plano {idx}"); self.abas_criadas[idx] = frame_aba; self.notebook.select(frame_aba)
        if len(self.abas_criadas) >= MAX_ABAS: self.btn_add.config(state="disabled")
    def fechar_aba(self, indice):
        if indice in self.abas_criadas: frame_aba = self.abas_criadas[indice]; self.notebook.forget(frame_aba); del self.abas_criadas[indice]
        if len(self.abas_criadas) < MAX_ABAS: self.btn_add.config(state="normal")
        if not self.abas_criadas: self.add_aba() # Adiciona uma nova se fechar a última
    def get_abas_ativas(self): return [self.abas_criadas[idx] for idx in sorted(self.abas_criadas.keys())]
    def _validar_dados_basicos(self):
        if not self.nome_closer_var.get() or not self.celular_closer_var.get() or not self.email_closer_var.get(): showerror("Dados Incompletos", "Preencha os dados do Vendedor."); return False
        if not self.nome_cliente_var_shared.get(): showerror("Dados Incompletos", "Preencha o Nome do Cliente."); return False
        return True
    def on_gerar_proposta(self):
        abas = self.get_abas_ativas();
        if not abas: showerror("Erro", "Nenhuma aba ativa."); return
        if not self._validar_dados_basicos(): return
        pptx = gerar_proposta(abas, self.nome_closer_var.get(), self.celular_closer_var.get(), self.email_closer_var.get())
        if pptx and os.path.exists(pptx): upload_pptx_and_export_to_pdf(pptx)
    def on_gerar_mat_tecnico(self):
        abas = self.get_abas_ativas();
        if not abas: showerror("Erro", "Nenhuma aba ativa."); return
        if not self._validar_dados_basicos(): return
        pptx = gerar_material(abas, self.nome_closer_var.get(), self.celular_closer_var.get(), self.email_closer_var.get())
        if pptx and os.path.exists(pptx): upload_pptx_and_export_to_pdf(pptx)
    def on_gerar_tudo(self):
        abas = self.get_abas_ativas();
        if not abas: showerror("Erro", "Nenhuma aba ativa."); return
        if not self._validar_dados_basicos(): return
        print("--- Iniciando Geração Completa ---"); pdf_prop_ok = False; pdf_mat_ok = False
        try:
            print("1. Gerando Proposta..."); pptx_prop = gerar_proposta(abas, self.nome_closer_var.get(), self.celular_closer_var.get(), self.email_closer_var.get())
            if pptx_prop and os.path.exists(pptx_prop): upload_pptx_and_export_to_pdf(pptx_prop); pdf_prop_ok = True
            else: print("   Falha ao gerar PPTX da proposta.")
        except Exception as e_prop: print(f"   ERRO na Proposta/PDF: {e_prop}")
        try:
            print("2. Gerando Material Técnico..."); pptx_mat = gerar_material(abas, self.nome_closer_var.get(), self.celular_closer_var.get(), self.email_closer_var.get())
            if pptx_mat and os.path.exists(pptx_mat): upload_pptx_and_export_to_pdf(pptx_mat); pdf_mat_ok = True
            else: print("   Falha ao gerar PPTX do material.")
        except Exception as e_mat: print(f"   ERRO no Material/PDF: {e_mat}")
        print(f"--- Geração Concluída (Proposta PDF: {'OK' if pdf_prop_ok else 'Falha'}, Material PDF: {'OK' if pdf_mat_ok else 'Falha'}) ---")

# --- Função Principal ---
def main():
    # Configura tema ttkbootstrap (opcional, pode ser feito na classe MainApp também)
    # style = ttkb.Style(theme='litera')
    app = MainApp()
    app.mainloop()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        # Captura erros gerais que podem ocorrer antes da UI iniciar
        print(f"ERRO FATAL: {e}")
        traceback.print_exc()
        # Tenta mostrar um erro final se o tkinter ainda puder ser usado
        try:
            root = tk.Tk()
            root.withdraw() # Esconde janela principal do Tkinter
            showerror("Erro Fatal", f"Ocorreu um erro crítico:\n{e}\n\nVerifique o console para detalhes.")
            root.destroy()
        except:
            pass # Ignora se nem o Tkinter puder ser usado