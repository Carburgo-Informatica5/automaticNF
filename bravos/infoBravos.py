import time
import yaml
import pyperclip
import subprocess
import pyautogui as gui
import pygetwindow as window
from unidecode import unidecode
import DANImail


class bravos:
    WAITING_PERIOD_LENGTH: int = 120
    WAITING_PERIOD_LENGTH_POPUP: int = WAITING_PERIOD_LENGTH / 3

    v_message_bundle: list[str] = list()

    def v_error_spotter(self, string):
        for stub in self.listof_stubs:
            if stub in unidecode(string).lower():
                self.v_message_bundle.append(string)

    def __init__(self, config: dict, m_queue: DANImail.Queue) -> None:
        # Quando verdadeira, está propriedade levanta uma exceção se o mouse estiver em algum canto da tela.
        gui.FAILSAFE = False
        self.m_queue = m_queue
        self.config = config
        try:
            self.v_config = yaml.safe_load(
                open(config["dani_vc"], "rb").read().decode()
            )
            self.v_queue = DANImail.Queue(self.v_config)
            self.v_queue.push(
                self.m_queue.make_message()
                .add_text("Teste do coletor de erros.", tag="h1")
                .add_text(
                    "Este não é garantido funcional, erros podem estar omitidos do log."
                ),
                DANImail.WriteTo.header,
            )
            self.listof_stubs: list[str] = list()
            for stub in self.v_config["stubs"]:
                self.listof_stubs.append(unidecode(str(stub)).lower())
            print(self.listof_stubs)
        except:
            self.m_queue.push(
                m_queue.make_message()
                .add_text("Falha em criar fila de email para vendas.", tag="h3")
                .add_text(["Desabilitando função.", self.m_queue.print_trace()])
            )
            self.v_error_spotter = lambda _, __: ()

    # Função para atualizar a janela do programa
    def busy_getWindowWithTitle(
        self, title: str, timeout_lenght: int
    ) -> window.Win32Window:
        start_time = time.time()
        while True:
            if time.time() - start_time > timeout_lenght:
                raise ValueError("Couldn't find window")
            try:
                return window.getWindowsWithTitle(title)[0]
            except:
                ()

    # Caso erro na janela do programa ele executa o ALT+F4 (fecha o programa)
    def kill_warning(self, timeout_lenght: int):
        listof_tittles = ["warning", "confirm", "error", "failure", "information"]
        start_time = time.time()
        while time.time() - start_time < timeout_lenght:
            listof_windows: list[window.Win32Window] = list()
            for title in listof_tittles:
                listof_windows += window.getWindowsWithTitle(title)
            for warn in listof_windows:
                gui.sleep(1)
                warn.activate()
                warn.restore()
                if "error" in warn.title.lower():
                    gui.sleep(1)
                    gui.press("enter")
                    if len(window.getWindowsWithTitle("error")) > 0:
                        gui.press("enter")
                else:
                    gui.hotkey("alt", "f4")
                start_time = time.time()
            # if len(listof_windows) > 0: break

    # Fecha o Popup do programa
    def close_document_popup(self, timeout_lenght: int):
        last_document = time.time()
        while True:
            # O espaço em braco é necessario.
            print(
                f"\r > Segundos desde a ultima janela: {int(time.time() - last_document)}       ",
                end="",
            )
            try:
                # Caso o bravos demore para carregar o documento.
                window.getWindowsWithTitle("relatório")[0]
                gui.sleep(0.5)
                continue
            except:
                ()

            try:
                info: window.Win32Window = window.getWindowsWithTitle("Information")[0]
                info.restore()
                gui.press("enter")
                continue
            except:
                ()

            # As vezes o DANI pode encontrar um IO error.
            listof_windows = list()
            try:
                listof_tittles = ["warning", "confirm", "error", "failure"]
                for title in listof_tittles:
                    listof_windows += window.getWindowsWithTitle(title)
            except:
                ()

            if len(listof_windows) > 0:
                warn: window.Win32Window
                for warn in listof_windows:
                    warn.activate()
                    warn.restore()
                    if "error" in warn.title.lower():
                        gui.press("enter")
                        gui.sleep(1)
                        if len(window.getWindowsWithTitle("error")) > 0:
                            gui.press("enter")
                    else:
                        gui.hotkey("alt", "f4")
                raise ValueError("Encontrou erro inexperado. Por favor refazer loja.")

            try:
                # Este timeout é extremamente longo pois a vezes o bravos fica de um a três segundos
                doc_window = window.getWindowsWithTitle("Pré-visualização")[0]
                last_document = time.time()
            except:
                print(self.m_queue.print_trace())
                if time.time() - last_document > timeout_lenght:
                    break
                continue
            doc_window: window.Win32Window
            doc_window.activate()
            doc_window.maximize()
            gui.hotkey("alt", "F")
        print("")

    def __find_login(self, timeout_lenght) -> window.Win32Window:
        TIMEOUT_LENGHT_UPDATE = 250
        start_time = time.time()
        while True:
            if time.time() - start_time > timeout_lenght:
                raise ValueError("Couldn't find window")

            try:
                window.getWindowsWithTitle("Atualização do")[0]
                self.m_queue.push(
                    self.m_queue.make_message()
                    .add_img(gui.screenshot())
                    .add_text(
                        f"Encontrado Janela de update, dormindo por {timeout_lenght} segundos.",
                        tag='h1 style="color: white;"',
                        tag_end="h1",
                    )
                )
                timeout_lenght = TIMEOUT_LENGHT_UPDATE
                gui.sleep(TIMEOUT_LENGHT_UPDATE / 16)
            except:
                ()

            try:
                return window.getWindowsWithTitle("Identificação")[0]
            except:
                ()
    
    def wait_for_login_window(self, timeout_lenght: int) -> window.Win32Window:
        """
        Aguarda a janela de 'Identificação' aparecer.
        """
        start_time = time.time()
        while True:
            if time.time() - start_time > timeout_lenght:
                raise ValueError("Não foi possível encontrar a janela de 'Identificação'")
            try:
                return window.getWindowsWithTitle("Identificação")[0]
            except IndexError:
                time.sleep(1) # Espera 1 segundo antes de tentar novamente


    # Pega as informações do user e faz o login dele
    def acquire_bravos(self, exec: str = "C:\\BravosClient\\FabricaVWClient.exe"):
        subprocess.run(
            [
                "powershell.exe",
                "-Command",
                'Get-Process | Where-Object {$_.Path -like "*bravos*"} | Stop-Process',
            ],
            shell=True,
        )
        subprocess.Popen(exec, start_new_session=True)
        self.kill_warning(5)

        login_window = self.wait_for_login_window(60) # Aumentei o timeout para 60 segundos
        login_window.activate()
        login_window.restore()

        gui.typewrite(self.config["bravos_usr"])
        gui.press("tab")
        gui.typewrite(self.config["bravos_pswd"])
        gui.press("enter")
        self.kill_warning(5)

        bravos: window.Win32Window
        bravos = window.getWindowsWithTitle("BRAVOS")[0]
        bravos.activate()
        bravos.maximize()
        self.kill_warning(3)
        return bravos

    # Reinicia o Bravos caso algum erro aconteça
    def upload(self, dealers: list[dict]) -> bool:
        bravos = self.acquire_bravos()
        for dealer in dealers:

            try:
                # As vezes o bravos pede para fazer a inclusão de um item no estoque que para por completo o processo de importação.
                # Nestes casos, só fechamos o bravos, abrimos ele de volta e pulamos para a proxima loja.
                window.getWindowsWithTitle("Inclusão")[0]
                dealer_msg.set_color("red")
                dealer_msg.add_text("Erro no bravos!", tag="h2")
                dealer_msg.add_text(f"Pulando {dealer['unit']}")
                dealer_msg.add_img(gui.screenshot())
                dealer_msg.add_text("Reinciando Bravos...")
                bravos = self.acquire_bravos()
                self.m_queue.push(dealer_msg)
                continue
            except:
                ()

            dealer_msg = (
                self.m_queue.make_message()
                .set_color("green")
                .add_text(
                    f"{dealer['unit']}:", tag='h1 style="color: white;"', tag_end="h1"
                )
                .insert_raw("<br>")
                .add_text("Informações sobre o cliente:")
            )

            if "bravos" not in window.getActiveWindowTitle().lower():
                try:
                    bravos.activate()
                    bravos.restore()
                except:
                    self.m_queue.push(
                        dealer_msg.set_color("red")
                        .add_text("Erro no bravos!", tag="h2")
                        .add_text(self.m_queue.print_trace(), tag="pre")
                        .add_img(gui.screenshot())
                        .add_text("Reinciando Bravos...")
                    )
                    bravos = self.acquire_bravos()

            try:
                gui.press("alt")
                gui.typewrite("CA")
                gui.press("tab")
                gui.press("tab")
                gui.typewrite(dealer["er"])
                gui.press("enter")
                self.kill_warning(6)

                gui.press("alt")
                gui.typewrite("frfm")
                gui.press("tab", presses=3)
                gui.typewrite(dealer["file_name"])
                gui.hotkey("shift", "tab")
                gui.press("enter")

                gui.sleep(25)
                gui.hotkey("alt", "I")
                gui.press("tab", presses=5)
                gui.press("space")
                gui.press("tab")
                gui.press("space")
                gui.hotkey("alt", "T")
                gui.hotkey("alt", "V")

                self.close_document_popup(self.WAITING_PERIOD_LENGTH_POPUP)
                gui.sleep(2)
                bravos.activate()
                gui.hotkey("alt", "N")
                gui.press("tab", presses=3)
                gui.hotkey("ctrl", "a")
                pyperclip.copy("")
                gui.hotkey("ctrl", "c")

                clipboard = pyperclip.paste()

                # Quando este erro acontence ele tende a ferrar tudo.
                # Por isso, estamos fechando o bravos por completo e continuando as proximas lojas como normal.
                if "error 103." in clipboard:
                    dealer_msg.set_color("red")
                    dealer_msg.add_text("Erro no bravos!", tag="h2")
                    dealer_msg.add_text(
                        f"Pulando {dealer['unit']} por conta de um erro 103..."
                    )
                    dealer_msg.add_img(gui.screenshot())
                    dealer_msg.add_text("Reinciando Bravos...")
                    bravos = self.acquire_bravos()
                    self.m_queue.push(dealer_msg)
                    continue

                logs = clipboard.split("\n")
                copyof_logs = list()
                for index in range(len(logs)):
                    logs[index] = (
                        logs[index].replace("\n", "").replace("\r", "").strip()
                    )
                    if not (logs[index] == " " or logs[index] == ""):
                        copyof_logs.append(logs[index])
                logs = copyof_logs
                (
                    dealer_msg.set_color("red" if "ERRO" in logs else "green")
                    .insert_raw("<br>")
                    .add_text(
                        f"    > Resultados da verificação de {dealer["unit"]} retorna: ",
                        tag="h2",
                    )
                    .insert_raw("<br>")
                    .add_text([f'arquivo: "{dealer["file_name"]}"'] + logs)
                )

                gui.hotkey("alt", "A")
                gui.hotkey("alt", "T")
                gui.press("tab", presses=2)
                gui.press("space")

                print(
                    f" > Hibernado por {self.WAITING_PERIOD_LENGTH} segundos de modo a esperar o download concluir."
                )
                for sec in range(self.WAITING_PERIOD_LENGTH):
                    while True:
                        try:
                            popup: window.Win32Window = window.getWindowsWithTitle(
                                "Confirm"
                            )[0]
                            popup.activate()
                            gui.press("enter")
                            gui.sleep(1)
                        except:
                            break
                    gui.sleep(1)
                    print(f"\r > Segundos totais: {sec}       ", end="")
                print("")

                (
                    dealer_msg.set_color("green")
                    .insert_raw("<br>")
                    .add_text(f"    > Ultima tela de {dealer["unit"]} foi:", tag="h2")
                    .insert_raw("<br>")
                    .add_img(gui.screenshot())
                )

                self.kill_warning(3)
                gui.hotkey("alt", "L")
                gui.hotkey("ctrl", "a")
                pyperclip.copy("")
                gui.hotkey("ctrl", "c")

                logs = pyperclip.paste().split("\n")
                copyof_logs = list()
                for index in range(len(logs)):
                    logs[index] = (
                        logs[index].replace("\n", "").replace("\r", "").strip()
                    )
                    if not (logs[index] == " " or logs[index] == ""):
                        copyof_logs.append(logs[index])
                logs: list[str] = copyof_logs
                (
                    dealer_msg.set_color(
                        "red"
                        if ("ERRO" in logs or "não encontrado" in logs)
                        else "green"
                    )
                    .insert_raw("<br>")
                    .add_text(
                        f"    > Logs de execução de {dealer["unit"]} retorna:", tag="h2"
                    )
                    .insert_raw("<br>")
                    .add_text(logs)
                )

                for log in logs:
                    self.v_error_spotter(log)
                if len(self.v_message_bundle) > 1:
                    self.v_queue.push(
                        self.v_queue.make_message()
                        .add_text(f"{dealer['unit']} ({dealer['er']}) ->", tag="h2")
                        .add_text(self.v_message_bundle)
                    )
                    self.v_message_bundle = list()
                dealer["upl_done"] = True
            except:
                dealer_msg.set_color("red")
                dealer_msg.add_text("Erro no bravos!", tag="h2")
                dealer_msg.add_text(self.m_queue.print_trace(), tag="pre")
                self.m_queue.push(dealer_msg)
                continue
            self.m_queue.push(dealer_msg)

        try:
            if len(self.v_queue._body_queue) > 0:
                self.v_queue.flush()
        except:
            self.m_queue.push(
                self.m_queue.make_message().add_text(
                    [
                        "Erro no v_eh!" "Mensagem não enviada.",
                        self.m_queue.print_trace(),
                    ]
                )
            )
        return True


# inicializa o arquivo
if __name__ == "__main__":
    dealer = dict()
    dealer["file_name"] = "C:\\Arquivos Diarios\\08-07-24\\CS08072024.pk"
    dealer["er"] = "2.5"
    dealer["unit"] = "CSS"

    import os

    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    config = yaml.safe_load(open(".\\data\\config.yaml", "r").read())
    eh = DANImail.Queue(yaml.safe_load(open(config["dani_ti"], "r")))
    eh._IS_DEBUG = True

    eh.push(eh.make_message().set_color("green").add_text("Rodando em modo de teste"))
    bravos(config, eh).upload([dealer])
    eh.flush()