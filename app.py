import win32com.client
import csv
import os
import re
from datetime import datetime
import pythoncom


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
    filter_condition = f"[ReceivedTime] >= '{
        today.strftime('%d/%m/%Y')} 12:00 AM'"

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
    pythoncom.CoUninitialize()


if __name__ == "__main__":
    # Caminho para salvar o CSV na rede
    output_folder = r"\\sqlsrv23\e$\VEDDARA\RECEITAS"
    extract_emails(output_folder)
