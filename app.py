import win32com.client
import csv
import os
import re
from datetime import datetime
import pythoncom
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


def send_error_email(error_message):
    """Envia um e-mail com a mensagem de erro."""
    try:
        sender_email = "guilherme.meijomil@sqltech.com.br"
        recipient_email = "guilherme.meijomil@sqltech.com.br"
        subject = "Erro no processo de extração de e-mails"
        body = f"Ocorreu um erro durante a execução do script de extração de e-mails:\n\n{error_message}"

        # Configuração do servidor SMTP (substitua pelos dados do seu servidor)
        smtp_server = "smtp.sqltech.com.br"  # Atualize para o servidor SMTP correto
        smtp_port = 587  # Porta geralmente usada para SMTP
        smtp_username = "guilherme.meijomil@sqltech.com.br"
        smtp_password = "231297Gui@"  # Senha fornecida

        # Monta o e-mail
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # Envia o e-mail
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            text = msg.as_string()
            server.sendmail(sender_email, recipient_email, text)

        print("E-mail de erro enviado com sucesso.")
    except Exception as e:
        print(f"Erro ao tentar enviar o e-mail: {e}")


def extract_emails(output_folder):
    """Extrai os e-mails recebidos no dia atual e salva em um arquivo CSV."""
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch(
        "Outlook.Application").GetNamespace("MAPI")

    # Obtém a pasta "Caixa de Entrada"
    inbox = outlook.GetDefaultFolder(6)  # 6 é a pasta "Caixa de Entrada"

    # Acessa a subpasta "Receitas.new"
    receitas_folder = inbox.Folders["Receitas.new"]

    today = datetime.now().date()

    # Formata corretamente a data para o filtro
    filter_condition = f"[ReceivedTime] >= '{today.strftime('%d/%m/%Y')} 12:00 AM'"

    try:
        # Obtém os e-mails da subpasta "Receitas.new"
        messages = receitas_folder.Items
        messages = messages.Restrict(filter_condition)

        # Agora, vamos filtrar o assunto manualmente
        emails_data = []
        for message in messages:
            if message.Class == 43:  # Verifica se é um e-mail
                try:
                    # Remove espaços extras no início e final do assunto
                    subject = message.Subject.strip()

                    # Ignora e-mails com "Re:" ou "Fwd:" no início, independentemente de maiúsculas ou espaços extras
                    if not re.match(r'^(Re:|Fwd:)\s?', subject, re.IGNORECASE):
                        # Acessa o remetente
                        sender = message.Sender
                        sender_name = sender.Name.strip()  # Nome do remetente
                        sender_email = sender.Address.strip()  # E-mail real do remetente

                        # Se o e-mail do remetente for um identificador do Exchange, tenta corrigir
                        if "/o=" in sender_email:
                            # O remetente é um usuário interno do Exchange, tenta obter o e-mail correto
                            try:
                                sender_email = sender.GetExchangeUser().PrimarySmtpAddress
                            except Exception as e:
                                print(
                                    f"Erro ao tentar obter o e-mail do Exchange para {sender_name}: {e}")

                        received_time = message.ReceivedTime.strftime(
                            "%Y-%m-%d %H:%M:%S")
                        emails_data.append(
                            [sender_name, sender_email, subject, received_time])
                except Exception as e:
                    print(f"Erro ao processar mensagem: {e}")
                    send_error_email(str(e))  # Envia e-mail em caso de erro

        # Garante que a pasta de saída exista
        os.makedirs(output_folder, exist_ok=True)

        # Define o caminho do arquivo CSV
        output_file = os.path.join(output_folder, "emails.csv")

        # Salva os dados em um arquivo CSV
        with open(output_file, "w", newline="", encoding="utf-8") as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow(
                ["Remetente", "E-mail", "Assunto", "Data de Recebimento"])
            writer.writerows(emails_data)

        print(f"Arquivo salvo em: {output_file}")
    except Exception as e:
        print(f"Erro ao processar os e-mails: {e}")
        send_error_email(str(e))  # Envia e-mail em caso de erro

    pythoncom.CoUninitialize()


if __name__ == "__main__":
    # Caminho para salvar o CSV na rede
    output_folder = r"\\sqlsrv23\e$\VEDDARA\RECEITAS"
    extract_emails(output_folder)
