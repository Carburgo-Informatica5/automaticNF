from datetime import datetime
import json
from pathlib import Path
from typing import Any, Dict, List

class RegistradorSistema:
    def __init__(self, diretorio_log: str = "logs"):
        self.diretorio_log = Path(diretorio_log)
        self.diretorio_log.mkdir(exist_ok=True)
        self.sessao_atual = datetime.now().strftime("%Y%m%d_%H%M")
        self.registro_sessao: List[Dict[str, Any]] = []

    def registrar_evento(self,
                         tipo_evento: str,
                         descricao: str,
                         dados: Dict[str, Any] = None,
                         status: str = "info") -> None:
        """Registra um novo evento no sistema"""
        evento = {
            "timestamp": datetime.now().isoformat,
            "tipo": tipo_evento,
            "descricao":  descricao,
            "status": status,
            "dados": dados or {}
        }
        self.registro_sessao.append(evento)
        self._escrever_arquivo(evento)

    def escrever_arquivo(self, evento: Dict[str, Any]) -> None:
        """Escrever Evento"""
        arquivo_log = self.diretorio_log / f"sessao_{self.sessao_atual}.log"
        with open(arquivo_log, "a", encoding="utf-8") as f:
            json.dump(evento, f, ensure_ascii=False)
            f.write("\n")

    def obter_resumo_sessao(self) -> Dict[str, Any]:
        """Gera um resumo da sessão atual"""
        return{
            "id_sessao": self.sessao_atual,
            "hora_inicio": self.registro_sessao[0]["timestamp"] if self.registro_sessao else None,
            "hora_fim": self.registro_sessao[-1]["timestamp"] if self.registro_sessao else None,
            "total_eventos": len(self.registro_sessao),
            "evento_por_status": self._contar_eventos_por_status()
        }

    def _contar_eventos_por_status(self) -> Dict[str, int]:
        """Conta eventos por status"""
        contagem_status = {}
        for evento in self.registro_sessao:
            status = evento["status"]
            contagem_status[status] =  contagem_status.get(status, 0) + 1
        return contagem_status

