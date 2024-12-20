import win32com.client
import csv
import os
import re
from datetime import datetime
import pythoncom


def clean_sender_email(sender_email):
    """Limpa o e-mail do remetente para uma forma mais legível."""
    exchange_regex = re.compile(r"/O=[^/]+/OU=[^/]+/CN=RECIPIENTS/CN=[^@]+")
    match = exchange_regex.match(sender_email)
    if match:
        return match.group().split("=")[-1]
    return sender_email


def extract_emails(output_folder):
    """Extrai os e-mails recebidos no dia atual e salva em um arquivo CSV."""
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch(
        "Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 é a pasta "Caixa de Entrada"

    today = datetime.now().date()

    # Filtra para pegar apenas os e-mails recebidos hoje
    filter_condition = f"[ReceivedTime] >= '{
        today.strftime('%m/%d/%Y')} 12:00 AM'"

    # Obtém os e-mails da caixa de entrada
    messages = inbox.Items
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
                    sender = clean_sender_email(message.SenderEmailAddress)
                    received_time = message.ReceivedTime.strftime(
                        "%Y-%m-%d %H:%M:%S")
                    emails_data.append([sender, subject, received_time])
            except Exception as e:
                print(f"Erro ao processar mensagem: {e}")

    # Garante que a pasta de saída exista
    os.makedirs(output_folder, exist_ok=True)

    # Define o caminho do arquivo CSV
    output_file = os.path.join(output_folder, "emails.csv")

    # Salva os dados em um arquivo CSV
    with open(output_file, "w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(["Remetente", "Assunto", "Data de Recebimento"])
        writer.writerows(emails_data)

    print(f"Arquivo salvo em: {output_file}")
    pythoncom.CoUninitialize()


if __name__ == "__main__":
    # Caminho para salvar o CSV na rede
    output_folder = r"\\sqlsrv23\e$\VEDDARA\RECEITAS"
    extract_emails(output_folder)
