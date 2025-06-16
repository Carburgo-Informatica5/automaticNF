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
        self.status: str = "default"
        self._inner_html_queue: list[str] = []
        self._inner_attachment_queue: list[MIMENonMultipart] = []
        self._start: Callable[[], str] = lambda: ""
        self._end: Callable[[], str] = lambda: ""

    def __get_random_cid(self) -> str:
        return "attachment-" + "".join(
            random.choices("abcdefghijklmnopqrstuvwxyzç1234567890", k=64)
        )

    def set_status(self, status: str) -> "Message":
        """Define o status da mensagem (ex: 'success', 'error') para aplicar o estilo CSS correto."""
        self.status = status
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

        status_to_class = {
            "success": "success-message",
            "error": "error-message",
            "default": "p",
        }
        css_class = status_to_class.get(self.status, "p")

        def is_traceback(text: str) -> bool:
            return "Traceback (most recent call last):" in text

        def html(string: str):
            # Formata tracebacks automaticamente com a classe 'code_block'
            if tag == "p" and is_traceback(string):
                return f'<pre class="code_block">{string}</pre>'

            # Aplica a classe de status se não for 'default'
            if css_class != "p":
                return f'<{tag} class="{css_class}">{string}</{tag_end}>'

            # Retorna texto padrão
            return f"<{tag}>{string}</{tag_end}>"

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
            with open("email_preview.html", "w", encoding="utf-8") as f:
                f.write(html)
            print("-> E-mail em modo debug. Preview salvo como 'email_preview.html'")
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
        css_path = os.path.join(os.path.dirname(__file__), "data", "style.css")
        css = ""
        if os.path.exists(css_path):
            with open(css_path, "r", encoding="utf-8") as f:
                css = f"<style>{f.read()}</style>"

        email_title = self._config.get("subject", "Relatório do Sistema")

        html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
{css}
</head>
<body>
    <div class="email-container">
        <div class="email-header">
            {email_title}
        </div>
        
        <div class="email-body">
            {self._header_queue}
            {self._body_queue}
        </div>

        <div class="email-footer">
            <p>Este é um e-mail automático gerado pelo sistema DANI.</p>
            <p>{datetime.now().strftime("%d/%m/%Y, %H:%M:%S")}</p>
        </div>
    </div>
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


if __name__ == "__main__":
    import pyautogui

    if not os.path.exists("./data/dani_ti.yaml"):
        print("Erro: Arquivo de configuração './data/dani_ti.yaml' não encontrado.")
    else:
        config = yaml.safe_load(
            open("./data/dani_ti.yaml", "r", encoding="utf-8").read()
        )

        dani = Queue(config, is_debug=True)

        message_success = (
            dani.make_message()
            .set_status("success")
            .add_text("Processo concluído com sucesso!")
            .add_text("Todos os relatórios foram processados e validados.")
            .add_img(pyautogui.screenshot())
        )

        try:
            _ = 1 / 0
        except Exception:
            traceback_info = dani.print_trace()

        message_error = (
            dani.make_message()
            .set_status("error")
            .add_text("Falha Crítica no Sistema", tag="h2")
            .add_text("Ocorreu um erro inesperado durante a execução do processo 'X'.")
            .add_text("Detalhes do erro:")
            .add_text(traceback_info)
        )

        (dani.push(message_success).push(message_error).flush())
