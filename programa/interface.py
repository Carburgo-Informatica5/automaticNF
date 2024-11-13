import tkinter as tk
from tkinter import ttk
import threading
from typing import Dict, List
import queue
import TKinterModernThemes as TKMT
import tkinter.font as tkfont


class interfaceMonitoramentoNF(TKMT.ThemedTKinterFrame):
    def __init__(self):
        # Initialize with title, theme, and mode
        super().__init__("Monitoramento NF", "sun-valley", "dark")

        self.fila_eventos = queue.Queue()
        self.configurar_interface()
        self.inicia_thread_atualizacao()

    def run(self):
        self.master.mainloop()

    def update(self):
        self.master.update()

    def configurar_interface(self):
        """Configura os Elementos da interface"""

        botao_fonte = tkfont.Font(size=12, weight="bold")

        estilo_botao = {
            "font": botao_fonte,
            "borderwidth": 0,
            "relief": tk.FLAT,
            "padx": 10,
            "pady": 5,
            "bg": "#00BFFF",
            "fg": "black",
            "activebackground": "#0080FF",
            "activeforeground": "white",
            "cursor": "hand2",
        }

        # Frame Principal
        frame_principal = ttk.Frame(self.master, padding="10")
        frame_principal.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E))

        # Status Geral
        self.frame_status = ttk.LabelFrame(
            frame_principal, text="Status do Sistema", padding="5"
        )
        self.frame_status.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E))

        self.rotulo_status = ttk.Label(self.frame_status, text="Sistema Iniciado")
        self.rotulo_status.grid(row=0, column=0)

        # Área Logs
        frame_log = ttk.LabelFrame(frame_principal, text="Registros", padding="5")
        frame_log.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.texto_log = tk.Text(frame_log, height=20, width=80)
        self.texto_log.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Barra de rolagem para Logs
        barra_rolagem = ttk.Scrollbar(
            frame_log, orient=tk.VERTICAL, command=self.texto_log.yview
        )
        barra_rolagem.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.texto_log["yscrollcommand"] = barra_rolagem.set

        # Estatísticas
        self.frame_estatisticas = ttk.LabelFrame(
            frame_principal, text="Estatísticas", padding="5"
        )
        self.frame_estatisticas.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E))

        # Botões de Controle
        frame_controle = ttk.Frame(frame_principal, padding="5")
        frame_controle.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E))

        tk.Button(
            frame_controle,
            text="Pausar",
            command=self.pausar_processamento,
            **estilo_botao,
        ).grid(row=0, column=0, padx=5)
        tk.Button(
            frame_controle,
            text="Continuar",
            command=self.Retomar_processamento,
            **estilo_botao,
        ).grid(row=0, column=1, padx=5)
        tk.Button(
            frame_controle,
            text="Gerar Relatório",
            command=self.gerar_relatorio,
            **estilo_botao,
        ).grid(row=0, column=2, padx=5)

    def inicia_thread_atualizacao(self):
        """Inicia a thread para atualizar a interface"""
        threading.Thread(target=self.atualizar_interface, daemon=True).start()

    def atualizar_interface(self):
        """Atualiza a interface com os eventos do sistema"""
        while True:
            try:
                evento = self.fila_eventos.get(timeout=1)
                self.texto_log.insert(
                    tk.END, f"{evento['timestamp']} - {evento['descricao']}\n"
                )
                self.texto_log.see(tk.END)
            except queue.Empty:
                continue

    def pausar_processamento(self):
        self.rotulo_status.config(text="Processamento Pausado")

    def Retomar_processamento(self):
        self.rotulo_status.config(text="Processamento Retomado")

    def gerar_relatorio(self):
        self.rotulo_status.config(text="Relatório Gerado")


if __name__ == "__main__":
    app = interfaceMonitoramentoNF()
    app.run()
