import os
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from dotenv import load_dotenv

load_dotenv()

EMAIL_REMETENTE = os.getenv("email")
EMAIL_SENHA = os.getenv("senha")
CAMINHO_IMAGEM_ASSINATURA = r"C:\Projects\Vamos_EnvioMedicao\imagem (5).png"

async def enviar_email(destinatario: str, nome_cliente: str, caminho_anexo: str, link: str, validacao: str):
    try:
        mes_ano = datetime.now().strftime('%B/%Y').capitalize()
        link_limpo = str(link).strip()

        assinatura_html = """
            <br>
            <img src="cid:assinatura_img" style="width:300px;" alt="Assinatura">
        """

        if validacao:
            corpo = f"""
                <html>
                <body>
                    <p>Prezado {nome_cliente},</p>
                    <p>Segue medição para validação, referente {mes_ano}:</p>
                    <p><a href="{link_limpo}" target="_blank">{nome_cliente}-{mes_ano}</a></p>
                    <p><strong>Estipulamos o prazo de 5 dias úteis para a análise e validação, caso não haja retorno, as faturas serão emitidas por nosso sistema.</strong></p>
                    <p>Qualquer dúvida, fico à disposição.</p>
                    <p>Atenciosamente,</p>
                    {assinatura_html}
                </body>
                </html>
            """
        else:
            corpo = f"""
                <html>
                <body>
                    <p>Prezado {nome_cliente},</p>
                    <p>Segue medição, nota de débito e boleto de multas referente a {mes_ano}:</p>
                    <p><a href="{link_limpo}" target="_blank">{nome_cliente}-{mes_ano}</a></p>
                    <p>Qualquer dúvida, fico à disposição.</p>
                    <p>Atenciosamente,</p>
                    {assinatura_html}
                </body>
                </html>
            """

        # ✉️ Monta o e-mail
        msg = MIMEMultipart('related')
        msg['From'] = EMAIL_REMETENTE
        if isinstance(destinatario, str):
            destinatarios = [destinatario]
        else:
            destinatarios = destinatario

# Adiciona e-mails fixos
            destinatarios.extend([
    'gvm.multas@grupovamos.com.br',
    'elita.cruz@grupovamos.com.br',
    'laura.silvasantos@grupovamos.com.br'
            ])

# Remove duplicados e espaços
        destinatarios = list(set(e.strip() for e in destinatarios if e.strip()))
        msg['To'] = ", ".join(destinatarios)
        msg['Subject'] = f"Medição do cliente {nome_cliente}"

        # Anexa o corpo HTML
        msg_alternativo = MIMEMultipart('alternative')
        msg.attach(msg_alternativo)
        msg_alternativo.attach(MIMEText(corpo, 'html'))

        # Adiciona imagem da assinatura como inline (cid:assinatura_img)
        if os.path.exists(CAMINHO_IMAGEM_ASSINATURA):
            with open(CAMINHO_IMAGEM_ASSINATURA, 'rb') as img:
                mime_image = MIMEImage(img.read())
                mime_image.add_header('Content-ID', '<assinatura_img>')
                msg.attach(mime_image)

        # Anexa o arquivo, se existir
        if caminho_anexo and os.path.exists(caminho_anexo):
            with open(caminho_anexo, 'rb') as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(caminho_anexo))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(caminho_anexo)}"'
                msg.attach(part)

        # Envia e-mail via Outlook SMTP
        with smtplib.SMTP("smtp.office365.com", 587) as server:
            server.starttls()
            server.login(EMAIL_REMETENTE, EMAIL_SENHA)
            server.send_message(msg)

        print(f"✅ E-mail enviado com sucesso para {destinatario}")

    except Exception as e:
        print(f"❌ Erro ao enviar e-mail para {destinatario}: {e}")
