from typing import Optional
import logging
import traceback
from datetime import datetime
import os

from typing import Optional
from datetime import datetime

class ExcecaoNF(Exception):
    """Classe base para exceções customizadas do sistema"""
    def __init__(self, mensagem: str, detalhes: Optional[dict] = None):
        self.mensagem = mensagem
        self.detalhes = detalhes or {}
        self.timestamp = datetime.now()
        super().__init__(self.mensagem)
        
    def to_dict(self) -> dict:
        """Converte a exceção para um dicionário"""
        return {
            'tipo': self.__class__.__name__,
            'mensagem': self.mensagem,
            'detalhes': self.detalhes,
            'timestamp': self.timestamp.isoformat()
        }

class ErrosBravosConexao(ExcecaoNF):
    """Erro de conexão do sistema Bravos"""
    def __init__(self, mensagem: str, tentativas: int = 0, ultimo_erro: str = None):
        detalhes = {
            'tentativas_conexao': tentativas,
            'ultimo_erro': ultimo_erro
        }
        super().__init__(mensagem, detalhes)
        
    @classmethod
    def timeout_conexao(cls, tempo_espera: int):
        """Cria uma exceção especifíca para o timeout de conexão"""
        return cls(
            f"Timeout na conexão com o Bravos após {tempo_espera} segundos",
            detalhes={'tempo_espera': tempo_espera}
        )
    
    @classmethod
    def falha_login(cls, usuario: str, tentativas: int):
        """Cria exceção especifíca para a falha de login"""
        return cls(
            f"Falha no login do usuário após {tentativas} tentativas",
            tentativas=tentativas,
            ultimo_erro="Credencias Inválidas"
        )

class ErroParseXML(ExcecaoNF):
    """Erro de processamento do XML do sistema Bravos"""
    def __init__(self, mensagem: str, arquivos: str = None, linha: int = None, coluna: int = None):
        detalhes = {
            'arquivos': arquivos,
            'linha': linha,
            'coluna': coluna
        }
        super().__init__(mensagem, detalhes)
        
    @classmethod
    def arquivo_invalido(cls, caminho_arquivo: str):
        """Cria uma exceção para o arquivo XML inválido"""
        return cls(
            f"Arquivo XML inválido: {caminho_arquivo}",
            arquivos=caminho_arquivo
        )

    @classmethod
    def tag_nao_encontrada(cls, tag: str, arquivo: str):
        """Cria uma exceção para a tag não encontrada no arquivo XML"""
        return cls(
            f"Tag '{tag}' não encontrada no XML"
        )

    @classmethod
    def valor_invalido(cls, campo: str, valor: str, arquivo: str):
        """Cria uma exceção para o valor inválido do campo no arquivo XML"""
        return cls(
            f"Valor Inválido '{valor}' para o campo '{campo}'",
            arquivo=arquivo
        )

class SimuladorBravos:
    def __init__(self, modo_teste=False):
        self.modo_teste = modo_teste
        
    def conectar(self):
        if self.modo_teste:
            print("Conexão simulada com sucesso")
            return True
        else:
            # Colocar a lógica real de conexão
            pass
    
    def parse_xml(self, arquivo):
        if self.modo_teste:
            return{
                'numero_nf': '12345',
                'data_emissao': '2023-10-10',
                'numero_nf': 1000.00
            }
        else:
            # Colocar a lógica real de parse do XML
            pass

# Exemplo de uso
def exemplo_uso():
    try:
        # Simulando a connexão com o Bravos
        raise ErrosBravosConexao.timeout_conexao(30)
    except ErrosBravosConexao as e:
        print(f"Erro: {e.mensagem}")
        print(f"Detalhes: {e.detalhes}")

    try:
        # Simulando um erro de parsing XML
        raise ErroParseXML.tag_nao_encontrada("NFe", "arquivo.xml")
    except ErroParseXML as e:
        print(f"Erro: {e.mensagem}")
        print(f"Detalhes: {e.detalhes}")

class tratadorErros:
    def __init__(self, arquivo_log: str = "automaticNF/logs/registro_de_erros.txt"):
        log_dir = os.path.dirname(arquivo_log)
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        self.logger = self._configurar_logger(arquivo_log)
        self.logger = self._configurar_logger(arquivo_log)
        
    def _configurar_logger(self, arquivo_log: str) -> logging.Logger:
        # Configura o sistema de logging
        logger = logging.getLogger('SistemaNF')
        logger.setLevel(logging.DEBUG)
        
        # Manipulador de arquivo
        manipulador_arquivo = logging.FileHandler(arquivo_log)
        manipulador_arquivo.setLevel(logging.DEBUG)
        
        # Manipulador de Console
        manipulador_console = logging.StreamHandler()
        manipulador_console.setLevel(logging.DEBUG)
        
        # Formato do Log
        formatador = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        manipulador_arquivo.setFormatter(formatador)
        manipulador_console.setFormatter(formatador)
        
        logger.addHandler(manipulador_arquivo)
        logger.addHandler(manipulador_console)
        
        return logger
    
    def tratarErros(self, erro: Exception, contexto: str) -> Optional[dict]:
        """
        Processar e registar os erros 
        Retorna: dict com as informações do erro ou None 
        """
        info_erro = {
            'timestamp': datetime.now().isoformat(),
            'tipo': type(erro).__name__,
            'mensagem': str(erro),
            'contexto': contexto,
            'traceback': traceback.format_exc()
        }
        
        self.logger.error(
            "Erro em %s:\nTipo: %s\nMensagem: %s\nTraceback: %s",
            contexto,
            info_erro['tipo'],
            info_erro['mensagem'],
            info_erro['traceback']
        )
        
        return info_erro