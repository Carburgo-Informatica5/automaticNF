import smtplib, ssl, enum, io
from email.mime.multipart import MIMEMultipart
from email.mime.nonmultipart import MIMENonMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from PIL.Image import Image
from os import path, makedirs
from datetime import datetime
import random
from typing import *
from email.message import EmailMessage
import traceback


class WriteTo(enum.Enum):
    header = 1
    body = 2


class Message:
    def __init__(self, queue) -> None:
        self.color: str = "red"
        self._inner_html_queue: list[str] = []
        self._inner_attachment_queue: list[MIMENonMultipart] = []
        self._start: str = lambda: "" 
        self._end: str = lambda: ""

    def __get_random_cid(self) -> str:
        return "attachment-" + "".join(
            random.choices("abcdefghijklmnopqrstuvwxyzç1234567890", k=64)
        )

    def set_color(self, color: str) -> "Message":
        self.color = color
        return self

    def insert_raw(self, content: str | list[str]) -> "Message":
        if isinstance(content, str):
            self._inner_html_queue.append(content)
        else:
            self._inner_html_queue.extend(content)
        return self

    def add_text(
        self, content: str | list[str], tag: str = "p", tag_end: str = None
    ) -> "Message":
        tag_end = tag_end or tag

        def html(string: str):
            string = (
                string.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace('"', "&quot;")
                .replace("'", "&#39;")
            )
            return f"<{tag}>{string}</{tag_end}>"

        if isinstance(content, str):
            self._inner_html_queue.append(html(content))
        else:
            self._inner_html_queue.extend([html(text) for text in content])
        return self

    def add_img(self, img: Image | bytes) -> "Message":
        if isinstance(img, Image):
            buffer = io.BytesIO()
            img.save(buffer, "JPEG")
            image = MIMEImage(buffer.getvalue())
        elif isinstance(img, bytes):
            image = MIMEImage(img)
        else:
            raise ValueError("Image must be a PIL.Image or bytes.")
        image_name = self.__get_random_cid()
        image.add_header("Content-ID", f"<{image_name}>")
        self._inner_attachment_queue.append(image)
        self._inner_html_queue.append(f'<img src="cid:{image_name}"><br>')
        return self

    def _compose(self) -> tuple[str, list[MIMENonMultipart]]:
        result = self._start()
        result += "".join(self._inner_html_queue)
        result += self._end()
        return result, self._inner_attachment_queue


class Queue:
    def __init__(self, config: dict, is_debug: bool = False) -> None:
        self._queue: list[EmailMessage] = []
        self._config = config
        self._IS_DEBUG = is_debug
        self._header_queue: str = ""
        self._body_queue: str = ""
        self._attachments_queue: list[MIMENonMultipart] = []
        self._has_error: bool = False

        # Corrige o assunto para usar strftime se necessário
        subject = self._config.get("subject", "Sistema de lancamento de notas fiscais")
        try:
            self._config["subject"] = datetime.now().strftime(subject)
        except Exception:
            pass  # Se não for um formato válido, mantém como está

    def make_message(self, message: Type[Message] = Message) -> Message:
        return message(self)

    def push(self, message: Message, at: WriteTo = WriteTo.body):
        text, attachments = message._compose()
        self._attachments_queue += attachments
        if at == WriteTo.body:
            self._body_queue += text
        else:
            self._header_queue += text
        return self

    def flush(self):
        if not self._body_queue and not self._header_queue:
            return
        html = self._build_html()
        if self._IS_DEBUG:
            print(
                "\n\n-> [MODO DEBUG ATIVADO] O e-mail não será enviado. Conteúdo abaixo:"
            )
            print(html)
            self._clear_queues()
            return
        msg = MIMEMultipart("related")
        msg["From"] = self._config["from"]
        msg["Subject"] = self._config["subject"]
        msg["To"] = (
            ", ".join(self._config["to"])
            if isinstance(self._config["to"], list)
            else self._config["to"]
        )
        msg.attach(MIMEText(html, "html"))
        for attachment in self._attachments_queue:
            msg.attach(attachment)
        try:
            server = smtplib.SMTP(self._config["smtp_sv"], self._config["smtp_prt"])
            server.set_debuglevel(0)
            server.starttls(context=ssl.create_default_context())
            server.login(user=self._config["from"], password=self._config["pswd"])
            server.send_message(msg=msg)
            server.close()
            print("-> E-mail enviado com sucesso!")
            self._save_email_spool(str(msg))
        except (smtplib.SMTPException, EOFError, ConnectionRefusedError) as e:
            print("!!! ERRO AO ENVIAR E-MAIL:", e)
        finally:
            self._clear_queues()

    def _build_html(self) -> str:
        def text_sanitize(string):
            return " ".join(
                string.replace("\n", "")
                .replace("\r", "")
                .replace("    ", " ")
                .strip()
                .split()
            )

        style = ""
        if "style" in self._config and path.isfile(self._config["style"]):
            with open(self._config["style"], "r") as f:
                style = text_sanitize(f.read())
        signature = ""
        if "signature" in self._config and path.isfile(self._config["signature"]):
            try:
                with open(self._config["signature"], "rb") as f:
                    signature = f.read().decode()
            except Exception as e:
                print(f"Aviso: Não foi possível ler o arquivo de assinatura: {e}")
        html = f"""<html>
            <head><style>{style}</style></head>
            <header>{signature}{self._header_queue}</header>
            <body>{self._body_queue}</body>
        </html>"""
        return html

    def _save_email_spool(self, email_content: str):
        try:
            file_name = self._config["subject"]
            for char in "\\/:*<>|":
                file_name = file_name.replace(char, " ")
            count = 0
            makedirs(self._config["spool"], exist_ok=True)
            file_path_lambda = (
                lambda: f"{self._config['spool']}\\{file_name}{'' if count < 1 else f'({count})'}.eml"
            )
            while path.isfile(file_path_lambda()):
                count += 1
            with open(file_path_lambda(), "w", encoding="utf-8") as file:
                file.write(email_content)
            print(f"-> Cópia do e-mail salva em: {file_path_lambda()}")
        except Exception as e:
            print(f"Aviso: Falha ao salvar cópia do e-mail: {e}")

    def _clear_queues(self):
        self._body_queue = ""
        self._header_queue = ""
        self._attachments_queue = []

    def print_trace(self) -> str:
        return traceback.format_exc()


# Inicializa o arquivo
if __name__ == "__main__":
    import yaml
    import pyautogui

    config = yaml.safe_load(open("./data/dani_ti.yaml", "r").read())
    dani = Queue(config)

    # Testes de mensagens
    message = (
        dani.make_message()
        .set_color("green")
        .add_text("Isto é um teste com imagem")
        .add_img(pyautogui.screenshot())
        .add_text(
            [
                "Isto é uma mensagem de teste",
                "Isso támbem é um teste",
                "Isso não é um teste, favor entrar em panico de imediato.",
            ]
        )
        .add_img(pyautogui.screenshot())
        .add_text(["Main falhou silenciosamente.", dani.print_trace()])
    )

    message2 = (
        dani.make_message()
        .set_color("red")
        .add_text(
            ["Este teste prova se a fila foi resetada corretamente.", "eu espero"]
        )
        .add_text("ola! isso é um teste de titulo", tag="h1")
        .add_img(pyautogui.screenshot())
        .insert_raw("<h3><b><i>isto prova que é possivel inserir html</b></i></h3>")
        .add_text("<br><i>Isto prova que não é possivel injetar html arbitrario</i>")
    )

    (
        dani.push(message)
        .push(
            dani.make_message()
            .set_color("green")
            .add_text("Isto é uma mensagem de cabeçalho")
        )
        .flush()
        .push(message2)
        .flush()
    )
