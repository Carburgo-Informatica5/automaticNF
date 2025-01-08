# Sistema Automático de Processamento de Notas Fiscais

## Sumário
- [Sistema Automático de Processamento de Notas Fiscais](#sistema-automático-de-processamento-de-notas-fiscais)
  - [Sumário](#sumário)
  - [Descrição Geral](#descrição-geral)
  - [Estrutura do Projeto](#estrutura-do-projeto)
    - [Módulos Principais](#módulos-principais)
      - [1. `automaticNF/programa/main.py`](#1-automaticnfprogramamainpy)
      - [2. `automaticNF/programa/DANImail.py`](#2-automaticnfprogramadanimailpy)
      - [3. `automaticNF/programa/db_connection.py`](#3-automaticnfprogramadb_connectionpy)
      - [4. `automaticNF/programa/processar_xml.py`](#4-automaticnfprogramaprocessar_xmlpy)
  - [Requisitos do Sistema](#requisitos-do-sistema)
    - [Requisitos de Software](#requisitos-de-software)
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

#### 1. [`automaticNF/programa/main.py`](automaticNF/programa/main.py)
- **Função**: Módulo principal que coordena todas as operações do sistema.
- **Classe Principal**: `SistemaNF`
- **Funções Principais**:
  - `verificar_emails()`: Verifica e processa emails com notas fiscais anexadas.
  - `extract_values(text)`: Extrai valores do corpo do email.
  - `save_attachment(part, directory)`: Salva anexos de emails.
  - `processar_centros_de_custo(cc_texto, valor_total)`: Processa centros de custo.
  - `enviar_email_erro(dani, destinatario, erro)`: Envia email de erro.
  - `enviar_mensagem_sucesso(dani, destinatario, numero_nota)`: Envia email de sucesso.

#### 2. [`automaticNF/programa/DANImail.py`](automaticNF/programa/DANImail.py)
- **Função**: Gerencia a criação e envio de emails.
- **Classes Principais**:
  - `Message`: Cria mensagens de email.
  - `Queue`: Gerencia a fila de emails a serem enviados.
- **Métodos Principais**:
  - `Message.set_color(color)`: Define a cor da mensagem.
  - `Message.add_text(content, tag, tag_end)`: Adiciona texto à mensagem.
  - `Message.add_img(img)`: Adiciona imagem à mensagem.
  - `Queue.sendto_queue(message_bundle, screenshot, is_info)`: Envia mensagem para a fila.
  - `Queue.make_message(message)`: Cria uma nova mensagem.
  - `Queue.push(message, at)`: Adiciona mensagem à fila.
  - `Queue.flush()`: Envia todas as mensagens na fila.

#### 3. [`automaticNF/programa/db_connection.py`](automaticNF/programa/db_connection.py)
- **Função**: Gerencia a conexão com o banco de dados Oracle.
- **Funções Principais**:
  - `load_db_config()`: Carrega a configuração do banco de dados.
  - `connect_to_db(db_config)`: Conecta ao banco de dados.
  - `revenda(cnpj)`: Consulta informações de revenda no banco de dados.

#### 4. [`automaticNF/programa/processar_xml.py`](automaticNF/programa/processar_xml.py)
- **Função**: Processa arquivos XML de notas fiscais.
- **Funções Principais**:
  - `salvar_dados_em_arquivo(dados_nf, nome_arquivo, pasta_destino)`: Salva dados da nota fiscal em um arquivo JSON.
  - `parse_nota_fiscal(xml_file_path)`: Lê e extrai informações de um arquivo XML de nota fiscal.
  - `testar_processamento_local()`: Testa o processamento de arquivos XML localmente.

## Requisitos do Sistema

### Requisitos de Software
- Python 3.7 ou superior
- Sistema Bravos Client instalado
- Bibliotecas Python:
  - pyautogui
  - pygetwindow
  - unidecode
  - pyyaml
  - pyperclip
  - tkinter
  - xml
  - queue
  - threading
  - TKinterModernThemes

## Fluxo de Funcionamento

1. **Inicialização**
   - O sistema carrega as configurações do arquivo `config.yaml`.
   - O sistema tenta estabelecer conexão com o Bravos.

2. **Processamento de Notas**
   - O sistema busca arquivos XML no diretório especificado.
   - Cada arquivo XML é lido e suas informações são extraídas.
   - Os dados extraídos são validados.
   - O sistema navega pelo Bravos e insere os dados automaticamente.

3. **Monitoramento**
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
- Adicionar funcionalidades de relatórios.

## Suporte e Contato

Para suporte técnico ou dúvidas sobre o sistema, entre em contato com a equipe de desenvolvimento:

- Email: caetano.apollo@carburgo.com.br