import enum
import os
import random
import smtplib
import ssl
import yaml
from email.mime.multipart import MIMEMultipart
from email.mime.nonmultipart import MIMENonMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from PIL.Image import Image
from datetime import datetime
from typing import *
from email.message import EmailMessage


class WriteTo(enum.Enum):
    header = 1
    body = 2


class Message:
    def __init__(self, queue) -> None:
        self.color: str = "red"
        self._inner_html_queue: list[str] = []
        self._inner_attachment_queue: list[MIMENonMultipart] = []
        self._start: Callable[[], str] = lambda: ""
        self._end: Callable[[], str] = lambda: ""

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
            return f"<{tag} class='{self.color}'>{string}</{tag_end}>"

        if isinstance(content, str):
            self._inner_html_queue.append(html(content))
        else:
            for c in content:
                self._inner_html_queue.append(html(c))
        return self

    def add_img(self, img: Image | bytes) -> "Message":
        if isinstance(img, Image):
            import io

            buf = io.BytesIO()
            img.save(buf, format="PNG")
            buf.seek(0)
            image = MIMEImage(buf.read(), _subtype="png")
        elif isinstance(img, bytes):
            image = MIMEImage(img, _subtype="png")
        else:
            return self
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
            self._config["subject"] = subject.format(
                data=datetime.now().strftime("%d/%m/%Y - %H:%M")
            )
        except Exception:
            self._config["subject"] = subject

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
            context = ssl.create_default_context()
            server = smtplib.SMTP(self._config["smtp_sv"], self._config["smtp_prt"])
            server.set_debuglevel(0)  
            server.starttls(context=context)
            server.login(self._config["from"], self._config["pswd"])
            server.send_message(msg=msg)
            server.close()
            print("-> E-mail enviado com sucesso!")
            self._save_email_spool(str(msg))
        except (smtplib.SMTPException, EOFError, ConnectionRefusedError) as e:
            print(f"Erro ao enviar e-mail: {e}")
        finally:
            self._clear_queues()

    def _build_html(self) -> str:
        # Inclui o CSS externo se existir
        css_path = os.path.join(os.path.dirname(__file__), "data", "style.css")
        css = ""
        if os.path.exists(css_path):
            with open(css_path, "r", encoding="utf-8") as f:
                css = f"<style>{f.read()}</style>"
        html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
{css}
</head>
<body>
{self._header_queue}
{self._body_queue}
</body>
</html>
"""
        return html

    def _save_email_spool(self, email_content: str):
        spool_dir = os.path.join(os.path.dirname(__file__), "spool")
        if not os.path.exists(spool_dir):
            os.makedirs(spool_dir)
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"email_{now}.eml"
        with open(os.path.join(spool_dir, filename), "w", encoding="utf-8") as f:
            f.write(email_content)

    def _clear_queues(self):
        self._header_queue = ""
        self._body_queue = ""
        self._attachments_queue = []

    def print_trace(self) -> str:
        import traceback

        return traceback.format_exc()


# Inicializa o arquivo
if __name__ == "__main__":
    import pyautogui

    config = yaml.safe_load(open("./data/dani_ti.yaml", "r", encoding="utf-8").read())
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
