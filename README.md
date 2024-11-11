# Sistema Automático de Processamento de Notas Fiscais

## Sumário
- [Sistema Automático de Processamento de Notas Fiscais](#sistema-automático-de-processamento-de-notas-fiscais)
  - [Sumário](#sumário)
  - [Descrição Geral](#descrição-geral)
  - [Estrutura do Projeto](#estrutura-do-projeto)
    - [Módulos Principais](#módulos-principais)
      - [1. main.py](#1-mainpy)
      - [2. infoBravos.py](#2-infobravospy)
      - [3. interface.py](#3-interfacepy)
      - [4. tratamentoErros.py](#4-tratamentoerrospy)
  - [Requisitos do Sistema](#requisitos-do-sistema)
    - [Requisitos de Software](#requisitos-de-software)
  - [Configuração e Instalação](#configuração-e-instalação)
  - [Fluxo de Funcionamento](#fluxo-de-funcionamento)
  - [Tratamento de Erros](#tratamento-de-erros)
  - [Manutenção e Suporte](#manutenção-e-suporte)
    - [Logs do Sistema](#logs-do-sistema)
    - [Backup](#backup)
    - [Atualização](#atualização)
  - [Limitações Conhecidas](#limitações-conhecidas)
  - [Considerações de Segurança](#considerações-de-segurança)
  - [Futuras Melhorias](#futuras-melhorias)
  - [Suporte e Contato](#suporte-e-contato)

## Descrição Geral

O Sistema Automático de Processamento de Notas Fiscais é uma solução desenvolvida para automatizar o processo de lançamento de notas fiscais no sistema Bravos. Este sistema realiza a leitura de arquivos XML de notas fiscais, extrai as informações relevantes e as insere automaticamente no sistema Bravos através de automação de interface.

Atualização: O sistema agora inclui uma funcionalidade de tratamento de erros mais robusta, utilizando exceções personalizadas para diferentes tipos de erros que podem ocorrer durante o processamento de notas fiscais e a interação com o sistema Bravos.


## Estrutura do Projeto

### Módulos Principais

#### 1. main.py
- **Função**: Módulo principal que coordena todas as operações do sistema.
- **Classe Principal**: `SistemaNF`
- **Funções Principais**:
  - `parse_nota_fiscal(xml_file_path)`: Realiza o parsing dos arquivos XML.
  - `processar_notas_fiscais(xml_folder)`: Coordena o processamento em lote das notas.
  - `inserir_dados_no_bravos(dados_nf)`: Gerencia a inserção dos dados no sistema Bravos.

#### 2. infoBravos.py
- **Função**: Responsável pela interação com o sistema Bravos.
- **Classe Principal**: `bravos`
- **Métodos Principais**:
  - `acquire_bravos(exec)`: Inicia o sistema Bravos e realiza o login.
  - `upload(dealers)`: Realiza o upload das informações para o Bravos.
  - `kill_warning(timeout_lenght)`: Trata janelas de aviso do Bravos.

#### 3. interface.py
- **Função**: Implementa a interface gráfica do sistema.
- **Classe Principal**: `interfaceMonitoramentoNF`
- **Métodos Principais**:
  - `configurar_interface()`: Configura os elementos da interface gráfica.
  - `atualizar_interface()`: Atualiza a interface com novos eventos.
  - `pausar_processamento()`: Pausa o processamento de notas.
  - `Retomar_processamento()`: Retoma o processamento de notas.
  - `gerar_relatorio()`: Gera um relatório do processamento.

#### 4. tratamentoErros.py
- **Função**: Sistema de tratamento de exceções personalizado.
- **Classes Principais**:
  - `ExcecaoNF`: Classe base para exceções customizadas.
  - `ErrosBravosConexao`: Trata erros de conexão com o Bravos.
  - `ErroParseXML`: Trata erros na leitura de arquivos XML.
- **Classe Auxiliar**: `tratadorErros`
  - Configura o sistema de logging.
  - Processa e registra os erros ocorridos.

## Requisitos do Sistema

### Requisitos de Software
- Python 3.7 ou superior
- Sistema Bravos Client instalado
- Bibliotecas Python:
  - pyautogui
  - pygetwindow
  - tkinter
  - xml.etree.ElementTree
  - yaml
  - TKinterModernThemes

## Configuração e Instalação

1. Clone o repositório do projeto:
'''git clone [URL_DO_REPOSITORIO]'''

2. Navegue até o diretório do projeto:
'''cd [NOME_DO_DIRETORIO]'''

3. Instale as dependências necessárias:
'''pip install -r requirements.txt'''

4. Execute o sistema:
'''python main.py'''

## Fluxo de Funcionamento

1. **Inicialização**
- O sistema carrega as configurações do arquivo `config.yaml`.
- A interface gráfica é inicializada.
- O sistema tenta estabelecer conexão com o Bravos.

2. **Processamento de Notas**
- O sistema busca arquivos XML no diretório especificado.
- Cada arquivo XML é lido e suas informações são extraídas.
- Os dados extraídos são validados.
- O sistema navega pelo Bravos e insere os dados automaticamente.

3. **Monitoramento**
- A interface gráfica exibe o progresso em tempo real.
- Logs são registrados para cada operação.
- Erros são tratados e registrados.
- Relatórios podem ser gerados ao final do processamento.

## Tratamento de Erros

O sistema utiliza um mecanismo robusto de tratamento de erros:

- **Logs Detalhados**: Todos os erros são registrados com timestamp, tipo, mensagem e traceback.
- **Recuperação Automática**: O sistema tenta recuperar-se de erros não críticos.
- **Notificações**: Erros críticos são exibidos na interface e registrados nos logs.
- **Exceções Personalizadas**: Facilitam a identificação e tratamento de erros específicos.

## Manutenção e Suporte

### Logs do Sistema
- **Localização**: `/logs/registro de erros.txt`
- **Formato**: `[Data/Hora] - [Tipo] - [Mensagem] - [Detalhes]`

### Backup
- Recomenda-se realizar backup diário dos arquivos processados.
- Manter cópias dos logs por pelo menos 30 dias.

### Atualização
- Verifique regularmente por atualizações no repositório do projeto.
- Teste as atualizações em um ambiente de desenvolvimento antes de aplicar em produção.

## Limitações Conhecidas

- O sistema requer uma interface gráfica do Windows para funcionar.
- Há dependência direta do sistema Bravos e sua interface.
- O processamento das notas é realizado de forma sequencial.

## Considerações de Segurança

- As credenciais do Bravos devem ser armazenadas de forma segura.
- Implemente validação rigorosa dos arquivos XML antes do processamento.
- Restrinja o acesso ao sistema apenas a usuários autorizados.
- Mantenha o sistema operacional e todas as dependências atualizadas.

## Futuras Melhorias

## Suporte e Contato

Para suporte técnico ou dúvidas sobre o sistema, entre em contato com a equipe de desenvolvimento:

- Email: caetano.apollo@carburgo.com.br
