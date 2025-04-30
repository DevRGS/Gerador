Okay, vamos detalhar a documentação deste script Python.

-----

## Documentação Detalhada do Script Gerador de Propostas e Material Técnico

**Índice:**

1.  [Visão Geral](https://www.google.com/search?q=%231-vis%C3%A3o-geral)
2.  [Dependências e Instalação Automática](https://www.google.com/search?q=%232-depend%C3%AAncias-e-instala%C3%A7%C3%A3o-autom%C3%A1tica)
3.  [Estrutura do Código](https://www.google.com/search?q=%233-estrutura-do-c%C3%B3digo)
      * [Importações Iniciais e Verificação de Dependências](https://www.google.com/search?q=%23importa%C3%A7%C3%B5es-iniciais-e-verifica%C3%A7%C3%A3o-de-depend%C3%AAncias)
      * [Importações Principais](https://www.google.com/search?q=%23importa%C3%A7%C3%B5es-principais)
      * [Função Utilitária: `baixar_arquivo_if_needed`](https://www.google.com/search?q=%23fun%C3%A7%C3%A3o-utilit%C3%A1ria-baixar_arquivo_if_needed)
      * [Ajustes Globais](https://www.google.com/search?q=%23ajustes-globais)
      * [Configurações do Vendedor](https://www.google.com/search?q=%23configura%C3%A7%C3%B5es-do-vendedor)
      * [Dados de Planos e Preços](https://www.google.com/search?q=%23dados-de-planos-e-pre%C3%A7os)
      * [Função Utilitária PPTX: `substituir_placeholders_no_slide`](https://www.google.com/search?q=%23fun%C3%A7%C3%A3o-utilit%C3%A1ria-pptx-substituir_placeholders_no_slide)
      * [Classe `PlanoFrame` (Aba de Cálculo)](https://www.google.com/search?q=%23classe-planoframe-aba-de-c%C3%A1lculo)
      * [Funções de Geração de PPTX (`gerar_proposta`, `gerar_material`)](https://www.google.com/search?q=%23fun%C3%A7%C3%B5es-de-gera%C3%A7%C3%A3o-de-pptx-gerar_proposta-gerar_material)
      * [Integração com Google Drive (Autenticação e Conversão PDF)](https://www.google.com/search?q=%23integra%C3%A7%C3%A3o-com-google-drive-autentica%C3%A7%C3%A3o-e-convers%C3%A3o-pdf)
      * [Classe Principal da Aplicação (`MainApp`)](https://www.google.com/search?q=%23classe-principal-da-aplica%C3%A7%C3%A3o-mainapp)
      * [Ponto de Entrada (`main`, `if __name__ == "__main__":`)](https://www.google.com/search?q=%23ponto-de-entrada-main-if-__name__--__main__)
4.  [Fluxo de Execução](https://www.google.com/search?q=%234-fluxo-de-execu%C3%A7%C3%A3o)
5.  [Arquivos Utilizados e Gerados](https://www.google.com/search?q=%235-arquivos-utilizados-e-gerados)

-----

### 1\. Visão Geral

Este script Python implementa uma aplicação de interface gráfica (GUI) utilizando `Tkinter` e o tema `ttkbootstrap`. O objetivo principal é auxiliar vendedores a montar propostas comerciais e materiais técnicos personalizados para clientes da "ConnectPlug" (inferido pelos nomes dos arquivos de template).

A aplicação permite:

  * Configurar e salvar dados do vendedor (nome, celular, email).
  * Criar múltiplas "abas", cada uma representando um cenário ou plano diferente para um mesmo cliente.
  * Selecionar um plano base (Personalizado, Ideal, Completo, etc.).
  * Adicionar módulos e funcionalidades extras.
  * Definir quantidades de Pontos de Venda (PDVs), usuários, terminais de autoatendimento, etc.
  * Calcular automaticamente os valores mensais e anuais do plano, aplicando descontos padrão ou permitindo a edição manual do valor anual ou do percentual de desconto.
  * Calcular o custo de treinamento (se aplicável).
  * Gerar arquivos `.pptx` (PowerPoint) para a Proposta Comercial e o Material Técnico, preenchendo automaticamente placeholders com os dados configurados nas abas e informações do cliente/vendedor. Os slides são filtrados dinamicamente com base nos planos e módulos selecionados.
  * Utilizar a API do Google Drive para fazer upload dos arquivos `.pptx` gerados, convertê-los para o formato Google Slides e, em seguida, exportá-los como arquivos `.pdf`, que são baixados localmente.

### 2\. Dependências e Instalação Automática

O script foi projetado para ser o mais autônomo possível em relação às suas dependências externas. No início da execução, ele verifica se as bibliotecas necessárias estão instaladas no ambiente Python.

  * **Bibliotecas Verificadas:**

      * `ttkbootstrap`: Para a interface gráfica moderna. (Pacote PyPI: `ttkbootstrap`)
      * `python-pptx`: Para manipulação de arquivos PowerPoint `.pptx`. (Pacote PyPI: `python-pptx`)
      * `google-api-python-client`: Biblioteca cliente oficial do Google para interagir com APIs como o Google Drive. (Pacote PyPI: `google-api-python-client`)
      * `google-auth-httplib2`: Biblioteca de transporte HTTP para autenticação Google. (Pacote PyPI: `google-auth-httplib2`)
      * `google-auth-oauthlib`: Para simplificar o fluxo de autenticação OAuth 2.0 do Google. (Pacote PyPI: `google-auth-oauthlib`)
      * `requests`: Para fazer requisições HTTP (usado para baixar templates e `client_secret`). (Pacote PyPI: `requests`)

  * **Mecanismo de Instalação:**

    1.  Um dicionário `dependencias` mapeia o nome do módulo (como é usado no `import`) para o nome do pacote no PyPI.
    2.  O script itera sobre este dicionário.
    3.  Para cada módulo, ele tenta importá-lo usando `importlib.import_module()`.
    4.  Se a importação falhar (gerando um `ImportError`), significa que a biblioteca não está instalada.
    5.  Nesse caso, o script imprime uma mensagem e utiliza o módulo `subprocess` para chamar o `pip` (o gerenciador de pacotes do Python) e instalar o pacote correspondente (`sys.executable` garante que o `pip` do ambiente correto seja usado).
    6.  Após uma tentativa de instalação, `importlib.invalidate_caches()` é chamado para garantir que o sistema de importação reconheça a nova biblioteca.
    7.  Uma segunda tentativa de importação é feita. Se falhar novamente, uma mensagem de erro é exibida.

Este processo garante que o usuário não precise instalar manualmente cada dependência antes de rodar o script pela primeira vez.

### 3\. Estrutura do Código

#### Importações Iniciais e Verificação de Dependências

  * `import sys`, `import subprocess`, `import importlib`: Módulos padrão do Python usados para interagir com o sistema, executar comandos externos (como `pip`) e gerenciar importações dinamicamente.
  * Bloco de verificação/instalação: Conforme descrito na seção [Dependências](https://www.google.com/search?q=%232-depend%C3%AAncias-e-instala%C3%A7%C3%A3o-autom%C3%A1tica).

#### Importações Principais

Após garantir que as dependências estão presentes, o script importa os módulos necessários para o restante da sua funcionalidade:

  * `os`: Interação com o sistema operacional (manipulação de caminhos, diretórios).
  * `requests`: Requisições HTTP.
  * `io`: Manipulação de streams de dados em memória (usado para baixar o PDF).
  * `pickle`: Serialização/desserialização de objetos Python (usado para salvar/carregar o token de autenticação do Google).
  * `json`: Leitura e escrita de dados no formato JSON (usado para o arquivo de configuração do vendedor).
  * `datetime.date`: Para obter a data atual e usá-la nos nomes dos arquivos gerados.
  * `tkinter` (como `tk`), `tkinter.ttk`, `tkinter.messagebox`: Componentes base da GUI e caixas de diálogo padrão.
  * `ttkbootstrap` (como `ttkb`): Widgets com temas modernos e classes base (`Window`, `Frame`, `Notebook`, etc.).
  * `pptx.Presentation`: Classe principal da biblioteca `python-pptx` para criar ou abrir arquivos `.pptx`.
  * Módulos do Google: Classes e funções específicas para autenticação (`InstalledAppFlow`, `Request`), construção do serviço da API (`build`) e manipulação de uploads/downloads (`MediaFileUpload`, `MediaIoBaseDownload`).

#### Função Utilitária: `baixar_arquivo_if_needed`

  * **`baixar_arquivo_if_needed(nome_arquivo, url)`**:
      * Verifica se um arquivo com o `nome_arquivo` especificado já existe no diretório atual.
      * Se **não** existir, imprime uma mensagem e usa a biblioteca `requests` para baixar o conteúdo da `url` fornecida.
      * Salva o conteúdo baixado no `nome_arquivo` local em modo binário (`"wb"`).
      * É usada para obter os templates `.pptx` (`Proposta Comercial ConnectPlug.pptx`, `Material Tecnico ConnectPlug.pptx`) e o arquivo `client_secret.json` do Google a partir de URLs do GitHub, caso ainda não estejam presentes localmente.

#### Ajustes Globais

  * `script_dir = os.path.dirname(os.path.abspath(__file__))`: Obtém o diretório onde o script está localizado.
  * `os.chdir(script_dir)`: Muda o diretório de trabalho atual para o diretório do script. Isso garante que arquivos relativos (como templates e config) sejam encontrados corretamente, independentemente de onde o script foi chamado.
  * `CONFIG_FILE = "config_vendedor.json"`: Nome do arquivo para salvar/carregar as configurações do vendedor.
  * `MAX_ABAS = 10`: Limite máximo de abas que podem ser criadas na interface.

#### Configurações do Vendedor

  * **`carregar_config(nome_closer_var, celular_closer_var, email_closer_var)`**:
      * Verifica se o arquivo `CONFIG_FILE` existe.
      * Se existir, tenta abri-lo e carregar os dados JSON.
      * Atualiza as variáveis `tk.StringVar` fornecidas com os valores lidos do JSON (ou com strings vazias se a chave não existir no arquivo).
      * Ignora erros de decodificação JSON (arquivo malformado) ou outros erros de leitura.
  * **`salvar_config(nome_closer, celular_closer, email_closer)`**:
      * Cria um dicionário com os dados atuais do vendedor.
      * Tenta garantir permissão de escrita no `CONFIG_FILE` (útil em alguns sistemas).
      * Abre o `CONFIG_FILE` em modo de escrita (`"w"`) com codificação UTF-8.
      * Salva o dicionário no arquivo JSON com formatação indentada (`indent=4`) e garantindo que caracteres não-ASCII sejam preservados (`ensure_ascii=False`).
      * Ignora erros de permissão ao salvar.

#### Dados de Planos e Preços

Estas são as estruturas de dados que definem a lógica de negócios da precificação:

  * **`PLAN_INFO` (dicionário)**:
      * Chaves: Nomes dos planos base ("Personalizado", "Ideal", "Completo", "Autoatendimento", "Bling", "Em Branco").
      * Valores: Dicionários contendo informações para cada plano:
          * `base_mensal`: Custo base mensal do plano.
          * `base_anual`: Custo base anual (geralmente com desconto, mas aqui parece ser apenas informativo ou um valor de referência antigo, pois o cálculo anual é feito dinamicamente). **Nota:** O cálculo real do anual usa um desconto sobre o mensal.
          * `min_pdv`: Quantidade mínima de PDVs incluída no plano base.
          * `min_users`: Quantidade mínima de usuários incluída no plano base.
          * `mandatory`: Lista de nomes de módulos que são obrigatórios e incluídos neste plano (não podem ser desmarcados).
  * **`SEM_DESCONTO` (conjunto)**:
      * Contém os nomes dos módulos ou itens que **não** recebem o desconto padrão quando o cliente opta pelo pagamento anual. Seus custos são somados integralmente ao valor anual.
  * **`precos_mensais` (dicionário)**:
      * Mapeia o nome de cada módulo adicional ou opção de incremento (como faixas de notas fiscais, TEF, BI, etc.) ao seu respectivo custo mensal. Módulos obrigatórios listados aqui podem ter custo 0 (como "3000 Notas Fiscais" no plano Ideal) ou um custo real que é considerado no cálculo.

#### Função Utilitária PPTX: `substituir_placeholders_no_slide`

  * **`substituir_placeholders_no_slide(slide, dados)`**:
      * Recebe um objeto `slide` do `python-pptx` e um dicionário `dados` (onde as chaves são os placeholders, e os valores são os textos substitutos).
      * Itera por todas as `shapes` (formas) dentro do `slide`.
      * Se a `shape` contiver texto (`has_text_frame`), itera por seus `paragraphs` e `runs` (segmentos de texto com a mesma formatação).
      * Para cada `run`, obtém o texto (`txt`).
      * Itera pelo dicionário `dados`. Se uma chave (placeholder) for encontrada dentro do `txt`, ela é substituída pelo valor correspondente.
      * Atualiza o texto do `run` com o texto modificado.
      * Esta função é crucial para personalizar os templates `.pptx`.

#### Classe `PlanoFrame` (Aba de Cálculo)

Esta classe representa a interface e a lógica de uma única aba no `Notebook` da aplicação principal. Herda de `ttkb.Frame`.

  * **`__init__(self, parent, aba_index, nome_cliente_var_shared, validade_proposta_var_shared, on_close_callback=None)`**:
      * Inicializa o frame.
      * Armazena o `aba_index` (número identificador da aba).
      * Recebe e armazena as variáveis `tk.StringVar` compartilhadas (`nome_cliente_var`, `validade_proposta_var`) da `MainApp`.
      * Cria uma `tk.StringVar` própria (`nome_plano_var`) para o nome do plano desta aba específica.
      * Inicializa variáveis de estado (`current_plan`, `spin_..._var` para quantidades, `var_notas` para seleção de NFs, `modules` para checkboxes, `check_buttons`, overrides de cálculo, valores calculados).
      * Configura o layout principal com `Canvas` e `Scrollbar` para permitir rolagem se o conteúdo exceder a altura da janela.
      * Chama `_montar_layout_esquerda()` e `_montar_layout_direita()` para criar os widgets.
      * Chama `configurar_plano("Personalizado")` para inicializar a aba com o plano padrão.
  * **`fechar_aba(self)`**: Chama o `on_close_callback` fornecido pela `MainApp` para solicitar o fechamento desta aba.
  * **`_montar_layout_esquerda(self)`**: Cria os widgets da parte esquerda da aba:
      * Barra superior com título da aba e botão "Fechar Aba".
      * `Labelframe` "Planos" com botões para selecionar o plano base.
      * `Labelframe` "Notas Fiscais" com `Radiobuttons` para selecionar a faixa de NFs (quando aplicável).
      * `Labelframe` "Outros Módulos" com `Checkbuttons` para todos os módulos opcionais, organizados em duas colunas. Armazena referências aos `Checkbutton` em `self.check_buttons` para poder habilitá-los/desabilitá-los.
      * `Labelframe` "Dados do Cliente" com `Entry` para Nome do Cliente, Validade da Proposta (usando as variáveis compartilhadas) e Nome do Plano (usando a variável própria da aba).
  * **`_montar_layout_direita(self)`**: Cria os widgets da parte direita:
      * `Labelframe` "Incrementos" com `Spinbox` para definir quantidades de PDVs, Usuários, Autoatendimento, Cardápio Digital, TEF, Smart TEF, App Gestão, Delivery Direto Básico.
      * `Labelframe` "Valores Finais" com `Labels` para exibir os resultados dos cálculos (Plano Mensal, Plano Anual, Custo Treinamento, Desconto %).
      * `Labelframe` "Plano (Anual) (editável)" com `Entry` para o usuário digitar um valor anual customizado e um botão "Reset Anual".
      * `Labelframe` "Desconto (%) (editável)" com `Entry` para o usuário digitar um percentual de desconto customizado e um botão "Reset Desconto".
  * **`on_user_edit_valor_anual(self, *args)` / `on_reset_anual(self)`**: Funções chamadas quando o campo de valor anual é editado ou resetado. Ativam/desativam o override do valor anual e disparam `atualizar_valores`.
  * **`on_user_edit_desconto(self, *args)` / `on_reset_desconto(self)`**: Funções chamadas quando o campo de desconto é editado ou resetado. Ativam/desativam o override do desconto e disparam `atualizar_valores`.
  * **`configurar_plano(self, plano)`**:
      * Chamada ao clicar em um botão de plano.
      * Define `self.current_plan`.
      * Define os valores iniciais dos `Spinbox` (PDVs, Usuários) com base nos mínimos do `PLAN_INFO`.
      * Desmarca todos os módulos opcionais e reabilita seus `Checkbuttons`.
      * Marca e desabilita os `Checkbuttons` dos módulos obrigatórios (`mandatory`) do plano selecionado.
      * Trata casos especiais (ex: marca "3000 Notas Fiscais" e desabilita para o plano Ideal; reseta seleção de NFs para outros planos).
      * Reseta os overrides de valor anual e desconto.
      * Chama `atualizar_valores()` para recalcular tudo.
  * **`atualizar_valores(self, *args)`**:
      * **Coração da lógica de cálculo.** Chamada sempre que qualquer controle que afeta o preço é modificado.
      * Obtém as informações do plano atual (`PLAN_INFO`).
      * Calcula o valor mensal:
          * Começa com a `base_mensal`.
          * Separa os custos em `parte_descontavel` e `parte_sem_desc`.
          * Adiciona custos de módulos marcados (verificando se pertencem a `SEM_DESCONTO` e se não são `mandatory`).
          * Adiciona custo da faixa de NF selecionada (se aplicável).
          * Adiciona custo de PDVs e Usuários extras (com lógica diferente para o plano "Bling").
          * Adiciona custos de itens `SEM_DESCONTO` (como TEF, Autoatendimento - com lógica específica).
          * Adiciona custos de incrementos dos Spinboxes (App Gestão, Delivery Direto, Cardápio Digital - com lógica de preço por quantidade).
          * Soma `parte_descontavel` e `parte_sem_desc` para obter `valor_mensal_automatico`.
      * Calcula o valor anual (`final_anual`):
          * Se `user_override_anual_active`, usa o valor digitado no `Entry` (`self.valor_anual_editavel`).
          * Se `user_override_discount_active`, aplica o desconto percentual digitado (`self.desconto_personalizado`) apenas à `parte_descontavel` e soma a `parte_sem_desc`.
          * Caso contrário (nenhum override), aplica um desconto padrão (10% = `0.10`) à `parte_descontavel` e soma a `parte_sem_desc`.
          * Atualiza o `Entry` `valor_anual_editavel` com o resultado formatado.
      * Calcula o `training_cost`:
          * Zero para "Autoatendimento" e "Em Branco".
          * Para outros planos, se `valor_mensal_automatico` for menor que 549.90, o custo é a diferença; senão, é zero.
      * Atualiza as `Labels` na interface com os valores calculados e formatados (`lbl_plano_mensal`, `lbl_plano_anual`, `lbl_treinamento`, `lbl_desconto`). Trata o caso do Autoatendimento não ter plano mensal.
      * Armazena os resultados finais em `self.computed_mensal`, `self.computed_anual`, `self.computed_desconto_percent`.
  * **`montar_lista_modulos(self)`**:
      * Cria uma lista de strings representando todos os itens incluídos no plano configurado nesta aba, para ser usada no placeholder `montagem_do_plano`.
      * Adiciona itens dos `Spinbox` (PDVs, Usuários, etc.), incluindo lógicas especiais (Usuário Cortesia por PDV extra, TEF Cortesia por Autoatendimento).
      * Adiciona a seleção de Notas Fiscais (priorizando Ilimitadas \> 3000 \> Faixa selecionada).
      * Adiciona módulos base ("Relatórios", "Vendas - Estoque - Financeiro").
      * Adiciona módulos marcados nos `Checkbuttons` (excluindo os que já foram tratados ou são base).
      * Remove duplicatas e retorna a lista final de strings.
  * **`gerar_dados_proposta(self, nome_closer, cel_closer, email_closer)`**:
      * Coleta todos os dados necessários para preencher os placeholders do template da proposta.
      * Obtém nome do plano, valores calculados (`computed_mensal`, `computed_anual`, `computed_desconto_percent`).
      * Formata os valores monetários como strings ("R$ XX,XX").
      * Calcula o `plano_mensal_str`, incluindo o custo de treinamento se houver (`R$ VALOR + R$ TREINO`).
      * Determina o `tipo_de_suporte` e `horario_de_suporte` com base no valor anual final (maior ou igual a R$ 269.90 = Estendido).
      * Chama `montar_lista_modulos()` para obter a lista formatada.
      * Calcula a `economia_anual` comparando o custo de 12x (mensal + treino) com 12x (anual).
      * Obtém dados do cliente e validade das variáveis compartilhadas.
      * Obtém dados do vendedor dos parâmetros.
      * Retorna um dicionário com todos esses dados, prontos para serem usados por `substituir_placeholders_no_slide`.

#### Funções de Geração de PPTX (`gerar_proposta`, `gerar_material`)

Estas funções orquestram a criação dos arquivos `.pptx` finais.

  * **`gerar_proposta(lista_abas, nome_closer, celular_closer, email_closer)`**:
    1.  Verifica se o template `Proposta Comercial ConnectPlug.pptx` existe.
    2.  Abre o template usando `Presentation()`.
    3.  Verifica se `lista_abas` não está vazia.
    4.  **Filtragem de Slides:**
          * Determina os índices das abas ativas (`abas_indices`).
          * Determina o conjunto de nomes de planos usados em todas as abas (`used_plans`).
          * Itera pelos slides do template. Para cada slide, verifica o texto em busca de:
              * Placeholders específicos de plano (ex: `"slide_bling"`): Mantém o slide apenas se o plano correspondente ("Bling") estiver em `used_plans`.
              * Placeholders específicos de aba (ex: `"aba_plano_1"`): Mantém o slide apenas se a aba com o índice correspondente (1) estiver ativa (`1 in abas_indices`). Mapeia o índice do slide original para o índice da aba (`slide_map_aba`).
              * Slides genéricos (sem placeholders de aba): Mantém sempre e mapeia para `None` em `slide_map_aba`.
    5.  Remove os slides que não foram marcados para manter (`keep_slides`), iterando de trás para frente.
    6.  **Substituição de Placeholders:**
          * Gera os dados para cada aba ativa usando `aba.gerar_dados_proposta()` e armazena em `dados_de_aba`.
          * Usa os dados da *primeira* aba como fallback (`d_fallback`).
          * Itera pelos slides restantes (já filtrados).
          * Se o slide foi mapeado para uma aba específica (`aba_num` não é `None`), usa os dados daquela aba (`dados_de_aba[aba_num]`) para chamar `substituir_placeholders_no_slide`.
          * Se o slide for genérico (`aba_num` é `None`), usa os dados de fallback (`d_fallback`).
    7.  **Salvar:**
          * Cria um nome de arquivo dinâmico: `Proposta ConnectPlug - [NomeCliente] - [Data].pptx`.
          * Salva a apresentação modificada com `prs.save()`.
          * Exibe mensagem de sucesso ou erro.
          * Retorna o nome do arquivo gerado ou `None`.
  * **`gerar_material(lista_abas, nome_closer, celular_closer, email_closer)`**:
    1.  Verifica se o template `Material Tecnico ConnectPlug.pptx` existe.
    2.  Abre o template.
    3.  Verifica se `lista_abas` não está vazia.
    4.  **Coleta de Módulos Ativos:**
          * Itera por todas as `lista_abas`.
          * Cria um conjunto (`modulos_ativos`) contendo o nome de **todos** os módulos e incrementos que estão ativos em **qualquer** uma das abas (união de todas as funcionalidades). Inclui itens dos checkboxes e dos spinboxes (TEF, Autoatendimento, etc.).
          * Cria um conjunto `planos_usados` com os nomes dos planos base utilizados.
    5.  **Filtragem de Slides:**
          * Usa um dicionário `MAPEAMENTO_MODULOS` que define qual placeholder de texto em um slide corresponde a qual módulo funcional (ou conjunto de módulos, como para Notas Fiscais). Placeholders como `"slide_sempre"` ou `"slide_bling"` também são considerados.
          * Itera pelos slides do template.
          * Para cada slide, verifica se seu texto contém algum placeholder do `MAPEAMENTO_MODULOS`.
          * Um slide é mantido (`slide_ok = True`) se:
              * Contém um placeholder de módulo (ex: `"check_tef"`) e o módulo correspondente ("TEF") está em `modulos_ativos`.
              * Contém um placeholder de plano (ex: `"slide_bling"`) e o plano correspondente ("Bling") está em `planos_usados`.
              * Contém um placeholder genérico como `"slide_sempre"`.
          * Armazena os índices dos slides a manter em `keep_slides`.
    6.  Remove os slides não marcados para manter.
    7.  **Substituição de Placeholders:**
          * Gera os dados da *primeira* aba ativa (`d_fallback`). O material técnico geralmente usa informações mais genéricas, então os dados da primeira aba são suficientes para placeholders como nome do cliente, vendedor, validade.
          * Itera pelos slides restantes e usa `substituir_placeholders_no_slide` com `d_fallback` em todos eles.
    8.  **Salvar:**
          * Cria um nome de arquivo dinâmico: `Material Tecnico ConnectPlug - [NomeCliente] - [Data].pptx`.
          * Salva a apresentação.
          * Exibe mensagem de sucesso ou erro.
          * Retorna o nome do arquivo gerado ou `None`.

#### Integração com Google Drive (Autenticação e Conversão PDF)

  * **`SCOPES = ['https://www.googleapis.com/auth/drive']`**: Define o escopo de permissão necessário: acesso total ao Google Drive do usuário (necessário para upload e talvez exclusão futura, embora a exclusão não esteja implementada).
  * **`baixar_client_secret_remoto()`**:
      * Define a URL do arquivo `client_secret.json` no GitHub.
      * Define um nome local temporário (`client_secret_temp.json`).
      * Chama `baixar_arquivo_if_needed` para garantir que o arquivo exista localmente.
      * Retorna o nome local do arquivo.
  * **`get_gdrive_service()`**:
      * Chama `baixar_client_secret_remoto()` para obter o caminho do arquivo de segredo do cliente.
      * Define o nome do arquivo de token (`token.json`).
      * Tenta carregar credenciais existentes do `token.json` usando `pickle`.
      * Se não houver credenciais ou se estiverem inválidas/expiradas:
          * Se expiradas, tenta atualizá-las usando o `refresh_token`.
          * Se não for possível atualizar ou não existirem, inicia o fluxo de autenticação OAuth 2.0:
              * Usa `InstalledAppFlow.from_client_secrets_file` para carregar as informações do cliente.
              * `flow.run_local_server(port=0)`: Abre o navegador do usuário para autorização, inicia um servidor local temporário para receber o código de autorização e obtém as credenciais.
              * Salva as novas credenciais (incluindo o `refresh_token`) no `token.json` usando `pickle` para uso futuro.
      * Usa as credenciais válidas para construir e retornar um objeto de serviço da API do Google Drive (`build('drive', 'v3', credentials=creds)`).
  * **`upload_pptx_and_export_to_pdf(local_pptx_path)`**:
    1.  Verifica se o arquivo `.pptx` local existe.
    2.  Chama `get_gdrive_service()` para obter o serviço autenticado.
    3.  Define o nome do arquivo PDF de saída trocando a extensão `.pptx` por `.pdf`.
    4.  **Upload e Conversão para Google Slides:**
          * Define metadados do arquivo, especificando o `mimeType` como `application/vnd.google-apps.presentation` para que o Drive o converta automaticamente.
          * Cria um objeto `MediaFileUpload` com o caminho do `.pptx` local e o `mimeType` correto do PowerPoint.
          * Chama `service.files().create()` para fazer o upload, passando os metadados e a mídia. `fields='id'` pede para retornar apenas o ID do arquivo criado no Drive.
          * Armazena o `file_id` retornado.
    5.  **Exportação para PDF:**
          * Chama `service.files().export_media()`, passando o `file_id` do Google Slides e o `mimeType` desejado (`application/pdf`).
          * Cria um `io.BytesIO()` para receber os dados do PDF em memória.
          * Cria um `MediaIoBaseDownload` para gerenciar o download em partes (chunks).
          * Executa o download em um loop `while not done`, imprimindo o progresso.
    6.  **Salvar PDF Localmente:**
          * Abre o arquivo `pdf_output_name` em modo de escrita binária (`"wb"`).
          * Escreve o conteúdo do `io.BytesIO` (que contém o PDF baixado) no arquivo local.
    7.  Exibe uma mensagem de sucesso informando o nome do arquivo PDF gerado.
    <!-- end list -->
      * **Importante:** Este processo deixa o arquivo convertido (Google Slides) no Google Drive do usuário. O script não o exclui.

#### Classe Principal da Aplicação (`MainApp`)

Herda de `ttkb.Window` e representa a janela principal da aplicação.

  * **`__init__(self)`**:
      * Inicializa a janela com o tema "litera".
      * Define título e tamanho inicial.
      * Cria `tk.StringVar` para os dados do vendedor (`nome_closer_var`, etc.).
      * Cria `tk.StringVar` **compartilhadas** (`nome_cliente_var_shared`, `validade_proposta_var_shared`) que serão passadas para todas as instâncias de `PlanoFrame`. Isso garante que a edição em uma aba reflita em todas.
      * Chama `carregar_config()` para preencher os campos do vendedor.
      * Define `self.protocol("WM_DELETE_WINDOW", self.on_close)` para chamar `salvar_config` ao fechar a janela.
      * Cria a barra superior (`top_bar`) com `Labels` e `Entry` para os dados do vendedor e o botão "+ Nova Aba".
      * Cria o `ttkb.Notebook` que conterá as abas `PlanoFrame`.
      * Cria a barra inferior (`bot_frame`) com os botões de ação: "Gerar Proposta + PDF", "Gerar Material + PDF", "Gerar TUDO + PDF".
      * Inicializa `self.abas_criadas` (dicionário para rastrear as instâncias de `PlanoFrame` por índice) e `self.ultimo_indice`.
      * Chama `self.add_aba()` para criar a primeira aba ao iniciar.
      * Chama `baixar_arquivo_if_needed` para garantir que os templates `.pptx` estejam disponíveis.
  * **`on_close(self)`**: Chamada ao fechar a janela. Salva as configurações do vendedor e destrói a janela.
  * **`add_aba(self)`**:
      * Verifica se o limite `MAX_ABAS` foi atingido.
      * Incrementa `self.ultimo_indice`.
      * Cria uma nova instância de `PlanoFrame`, passando o `notebook` como pai, o novo índice, as variáveis compartilhadas e a referência à função `self.fechar_aba` como callback.
      * Adiciona o novo `frame_aba` ao `notebook`.
      * Armazena a instância no dicionário `self.abas_criadas`.
  * **`fechar_aba(self, indice)`**:
      * Callback chamado pelo botão "Fechar Aba" dentro de `PlanoFrame`.
      * Remove a aba correspondente do `notebook` usando `self.notebook.forget()`.
      * Remove a referência da aba do dicionário `self.abas_criadas`.
  * **`get_abas_ativas(self)`**: Retorna uma lista das instâncias `PlanoFrame` atualmente abertas, ordenadas pelo índice.
  * **`on_gerar_proposta(self)`**:
      * Chamada pelo botão "Gerar Proposta + PDF".
      * Obtém a lista de abas ativas.
      * Chama `gerar_proposta()` passando as abas e os dados do vendedor.
      * Se a proposta `.pptx` foi gerada com sucesso, chama `upload_pptx_and_export_to_pdf()` para criar o PDF.
  * **`on_gerar_mat_tecnico(self)`**:
      * Chamada pelo botão "Gerar Material + PDF".
      * Similar a `on_gerar_proposta`, mas chama `gerar_material()`.
  * **`on_gerar_tudo(self)`**:
      * Chamada pelo botão "Gerar TUDO + PDF".
      * Obtém as abas ativas.
      * Chama `gerar_proposta()` e, se bem-sucedido, `upload_pptx_and_export_to_pdf()`.
      * Chama `gerar_material()` e, se bem-sucedido, `upload_pptx_and_export_to_pdf()`. Gera ambos os documentos e seus PDFs sequencialmente.

#### Ponto de Entrada (`main`, `if __name__ == "__main__":`)

  * **`main()`**: Função simples que cria a instância da `MainApp` e inicia o loop de eventos principal da interface gráfica (`app.mainloop()`).
  * **`if __name__ == "__main__":`**: Construção padrão em Python. Garante que a função `main()` só seja chamada quando o script é executado diretamente (não quando é importado como um módulo).

### 4\. Fluxo de Execução

1.  O script é iniciado.
2.  As dependências são verificadas e instaladas automaticamente via `pip` se necessário.
3.  As bibliotecas principais são importadas.
4.  O diretório de trabalho é ajustado para o local do script.
5.  A classe `MainApp` é instanciada.
6.  Dentro do `__init__` da `MainApp`:
      * Configurações do vendedor são carregadas de `config_vendedor.json`.
      * Templates `.pptx` (`Proposta...` e `Material...`) são baixados do GitHub se não existirem localmente.
      * A interface gráfica principal é montada.
      * A primeira aba (`PlanoFrame`) é criada e adicionada ao `Notebook`.
7.  O loop de eventos Tkinter (`app.mainloop()`) é iniciado, exibindo a janela e aguardando a interação do usuário.
8.  **Interação do Usuário:**
      * Pode editar os dados do vendedor (serão salvos ao fechar).
      * Pode adicionar mais abas (até `MAX_ABAS`).
      * Em cada aba, pode selecionar um plano base, marcar/desmarcar módulos, ajustar quantidades (PDVs, usuários, etc.). A cada mudança, `atualizar_valores` recalcula e exibe os preços e descontos.
      * Pode editar o Nome do Cliente e Validade da Proposta (compartilhado entre abas).
      * Pode definir um nome específico para o plano da aba.
      * Pode sobrescrever o valor anual ou o desconto percentual.
      * Pode fechar abas individuais.
9.  **Geração de Documentos:**
      * O usuário clica em um dos botões "Gerar...".
      * A função correspondente (`on_gerar_proposta`, `on_gerar_mat_tecnico`, `on_gerar_tudo`) é chamada.
      * A função `get_abas_ativas` coleta as instâncias `PlanoFrame` abertas.
      * A função `gerar_proposta` ou `gerar_material` é chamada:
          * Carrega o template `.pptx`.
          * Filtra os slides com base nas seleções das abas.
          * Chama `gerar_dados_proposta` (em cada aba relevante ou na primeira) para obter os dados.
          * Chama `substituir_placeholders_no_slide` para preencher o template.
          * Salva o novo arquivo `.pptx` com nome dinâmico.
      * Se o `.pptx` foi salvo, a função `upload_pptx_and_export_to_pdf` é chamada:
          * `get_gdrive_service` é chamado (pode disparar o fluxo de autenticação OAuth na primeira vez ou se o token expirou).
          * O `.pptx` é enviado para o Google Drive e convertido para Google Slides.
          * O Google Slides é exportado como PDF.
          * O PDF é baixado para o mesmo diretório local.
      * Mensagens de sucesso ou erro são exibidas.
10. **Fechamento:**
      * O usuário fecha a janela principal.
      * A função `on_close` é chamada.
      * Os dados atuais do vendedor são salvos em `config_vendedor.json`.
      * A aplicação termina.

### 5\. Arquivos Utilizados e Gerados

  * **Arquivos de Entrada/Configuração:**
      * `config_vendedor.json`: Armazena nome, celular e email do vendedor (criado/lido localmente).
      * `client_secret_temp.json`: Credenciais da aplicação Google Cloud (baixado do GitHub se não existir).
      * `token.json`: Token de autenticação OAuth 2.0 do usuário (criado localmente após a primeira autorização).
      * `Proposta Comercial ConnectPlug.pptx`: Template base para a proposta (baixado do GitHub se não existir).
      * `Material Tecnico ConnectPlug.pptx`: Template base para o material técnico (baixado do GitHub se não existir).
  * **Arquivos de Saída (Gerados Localmente):**
      * `Proposta ConnectPlug - [NomeCliente] - [DD-MM-YYYY].pptx`
      * `Proposta ConnectPlug - [NomeCliente] - [DD-MM-YYYY].pdf`
      * `Material Tecnico ConnectPlug - [NomeCliente] - [DD-MM-YYYY].pptx`
      * `Material Tecnico ConnectPlug - [NomeCliente] - [DD-MM-YYYY].pdf`
      * *(Onde `[NomeCliente]` é o nome preenchido na interface e `[DD-MM-YYYY]` é a data atual)*
  * **Arquivos no Google Drive (Gerados):**
      * Versões Google Slides dos arquivos `.pptx` enviados. Estes **não** são excluídos pelo script após a conversão para PDF.

-----

Esta documentação abrange os principais aspectos do script, desde a instalação de dependências até a lógica de negócios, interface gráfica e integração com serviços externos.