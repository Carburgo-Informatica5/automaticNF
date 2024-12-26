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

#- Class Defition -#
class WriteTo(enum.Enum):
    header = 1
    body = 2

# Coloca em uma caixa a mensagem que ela gera
class Message:
    def __init__(self, queue) -> None:
        self.color: str = "red"
        self._inner_html_queue: list[str] = list()
        self._inner_attachment_queue: list[MIMENonMultipart] = list()
        self._start: str = lambda: f"<div class=\"code_block {self.color}\"><h2> -#- Start -#- </h2><br>"
        self._end: str = lambda: f"<br><h2> -#- End -#- </h2></div>"

# Pega os caracteres da mensagem e a forma
    def __get_random_cid(self) -> str:
        result: str = "attachment-"
        for _ in range(64): result += random.choice('abcdefghijklmnopqrstuvwxyz1234567890')
        return result

# Define a cor dos textos
    def set_color(self, color: str) -> Self:
        self.color = color
        return self

    def insert_raw(self, content: str|list[str]) -> Self:
        if type(content) is str: 
            self._inner_html_queue.append(" ".join(content.split()))
            print(content)
        else:
            for text in content:
                self._inner_html_queue.append(" ".join(text.split()))
                print(text)
        return self

# Adiciona o texto com base em HTML criando a mensagem e e tags da semântica
    def add_text(self, content: str|list[str], tag:str = "p", tag_end: str = None) -> Self:
        def html(string: str):
            string = ( string
            .replace('&', "&amp;")
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('\"', '&quot;')
            .replace('\'', "&#39;")
            )
            return f"<{tag}>{string}</{tag_end}>"
        
        if tag_end is None: tag_end = tag
        if type(content) is str: 
            self._inner_html_queue.append(html(content))
            print(content)

        else: 
            for text in content: 
                self._inner_html_queue.append(html(text))
                print(text)
        return self

# Adiciona a imagem que é gerada, aceita somente em JPEG
    def add_img(self, img: Image|bytes) -> Self:
        if type(img) is Image:
            buffer = io.BytesIO()
            img.save(buffer, "JPEG")
            image = MIMEImage(buffer.getvalue())
        elif type(img) is bytes:
            image = MIMEImage(img)
        else:
            raise ValueError("Image must be a Image from pillow or bytes.")
        
        image_name = self.__get_random_cid()
        image.add_header('Content-ID', f"<{image_name}>")

        self._inner_attachment_queue.append(image)
        self._inner_html_queue += f"<img src=\"cid:{image_name}\"><br>"
        return self

# Juntas os elementos da Mensagem e a forma
    def _compose(self) -> tuple[str, list[MIMENonMultipart]]:
        result: str = "" 
        result += self._start()
        for message in self._inner_html_queue:
            result += message
        result += self._end()
        return (result, self._inner_attachment_queue)


class Queue:
    def __init__(self, config:dict) -> None:
        self._config = config
        self._config["subject"] = datetime.today().strftime(self._config["subject"])
        while self._config["spool"][-1] == '\\':
            self._config["spool"] = self._config["spool"][-1:]

    _IS_DEBUG = False
    _header_queue: str = ""
    _body_queue: str = ""
    _attachments_queue: list[MIMENonMultipart] = list()
    _has_error: bool = False

    def sendto_queue(self, message_bundle: list[str], screenshot: Image = None, is_info: bool = False) -> None:
        print("AVISO: essa função foi deprecada e deve ser substituida.")
        message = (
            self.make_message()
            .set_color("green" if is_info else "red")
            .add_text(message_bundle)
            )
        if screenshot is not None: message.add_img(screenshot)
        self.push(message)

# Faz a mensagem
    def make_message(self, message: Message = Message) -> Message:
        return message(self)

# Escreve o corpo da mensagem e a envia
    def push(self, message: Message, at:WriteTo = WriteTo.body):
        text, attachments = message._compose()
        self._attachments_queue += attachments
        match at: 
            case WriteTo.body: self._body_queue += text
            case WriteTo.header: self._header_queue += text
        return self

    def flush(self):
        if self._body_queue == "" and self._header_queue == "": return
        
        if self._IS_DEBUG:
            print("\n \n -> [DEBUG IS ON] Dumping queue to console. No emails sent.")
            if len(self._header_queue) > 0: print('\n', self._header_queue)
            if len(self._body_queue) > 0: print('\n', self._body_queue)
            if len(self._header_queue) < 1 and len(self._body_queue) < 1: print("\n The Queue is empty, therefore, no lines will be printed.")
            else: print("\n")
            return

        print("\n \n -> flushing queue...")
        msg = MIMEMultipart('related')
        msg['From'] = self._config["from"]
        msg['Subject'] = self._config["subject"]
        if type(self._config["to"]) is list: msg['To'] = ", ".join(self._config["to"])
        else: msg['To'] = self._config["to"]

        if self._has_error and self._body_queue.count("code_block") > 6:
            self._body_queue = f"<div class=\"code_block red\"><h2>Foram encontrados erros!</h2></div>" + self._body_queue

        # Monta o HTML
        text_sanitize = lambda string: ' '.join(string.replace('\n','').replace('\r', '').replace('    ', ' ').strip().split())
        style = text_sanitize(open(self._config["style"], 'r').read())
    
        html = "<html>"
        html += f"<style>{style}</style>"
        html += "<header>"
        try: html += text_sanitize(open(self._config["signature"], 'rb').read().decode())
        except: pass
        html += self._header_queue
        html += "</header>"
        html += "<body>"
        html += self._body_queue
        html += "</body>"
        html += "</html>"

        msg.attach(MIMEText(html, 'html'))
        for attachment in self._attachments_queue: msg.attach(attachment)

        server = smtplib.SMTP(self._config["smtp_sv"], self._config["smtp_prt"])
        server.starttls(context=ssl.create_default_context())
        server.login(user=self._config["from"], password=self._config["pswd"])
        server.send_message(msg=msg)
        server.close()

        self._body_queue = ""
        self._header_queue = ""
        self._attachments_queue = list()

        #print(msg)
        file_name:str = self._config["subject"]
        for char in "\\/:*<>|": file_name = file_name.replace(char, ' ')

        count: int = 0
        makedirs(self._config["spool"], exist_ok=True)
        file_path = lambda: f"{self._config["spool"]}\\{file_name}{"" if count < 1 else f"({count})"}.eml"
        while path.isfile(file_path()): count += 1
        with open(file_path(), 'w') as file: file.write(str(msg))

        return self

    #- aliases -#
    def print_trace(self) -> str:
        import traceback
        return traceback.format_exc()

# Inicializa o arquivo
if __name__ == '__main__':
    import yaml
    import pyautogui

    config = yaml.safe_load(open("./data/dani_ti.yaml", 'r').read())
    dani = Queue(config)

# Testes de mensagens
    message = (   
        dani.make_message()
        .set_color("green")
        .add_text("Isto é um teste com imagem")
        .add_img(pyautogui.screenshot())
        .add_text(["Isto é uma mensagem de teste", "Isso támbem é um teste", "Isso não é um teste, favor entrar em panico de imediato."])
        .add_img(pyautogui.screenshot())
        .add_text(["Main falhou silenciosamente.", dani.print_trace()])
    )

    message2 = (
        dani.make_message()
        .set_color("red")
        .add_text(["Este teste prova se a fila foi resetada corretamente.", "eu espero"])
        .add_text("ola! isso é um teste de titulo", tag="h1")
        .add_img(pyautogui.screenshot())
        .insert_raw("<h3><b><i>isto prova que é possivel inserir html</b></i></h3>")
        .add_text("<br><i>Isto prova que não é possivel injetar html arbitrario</i>")
    )

    (   dani
        .push(message)
        .push(dani.make_message().set_color("green").add_text("Isto é uma mensagem de cabeçalho"))
        .flush()
        .push(message2)
        .flush()
    )