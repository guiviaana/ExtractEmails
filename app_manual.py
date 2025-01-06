import win32com.client
import csv
import os
import re
from datetime import datetime
import pythoncom


def extract_emails_by_date_range(output_folder, start_date, end_date):
    """
    Extrai os e-mails recebidos dentro de um intervalo de datas e salva em um arquivo CSV.
    """
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch(
        "Outlook.Application").GetNamespace("MAPI")

    # Obtém a pasta "Caixa de Entrada"
    inbox = outlook.GetDefaultFolder(6)  # 6 é a pasta "Caixa de Entrada"

    # Acessa a subpasta "Receitas.new"
    receitas_folder = inbox.Folders["Receitas.new"]

    # Formata a condição de filtro para o intervalo de datas
    filter_condition = f"[ReceivedTime] >= '{start_date.strftime(
        '%d/%m/%Y')} 12:00 AM' AND [ReceivedTime] <= '{end_date.strftime('%d/%m/%Y')} 11:59 PM'"

    # Obtém os e-mails da subpasta "Receitas.new" com o filtro aplicado
    messages = receitas_folder.Items
    messages = messages.Restrict(filter_condition)

    emails_data = []
    for message in messages:
        if message.Class == 43:  # Verifica se é um e-mail
            try:
                # Remove espaços extras no início e final do assunto
                subject = message.Subject.strip()

                # Ignora e-mails com "Re:" ou "Fwd:" no início
                if not re.match(r'^(Re:|Fwd:)\s?', subject, re.IGNORECASE):
                    # Acessa o remetente
                    sender = message.Sender
                    sender_name = sender.Name.strip()  # Nome do remetente
                    sender_email = sender.Address.strip()  # E-mail do remetente

                    # Corrige e-mail do remetente interno do Exchange, se necessário
                    if "/o=" in sender_email:
                        try:
                            sender_email = sender.GetExchangeUser().PrimarySmtpAddress
                        except Exception as e:
                            print(
                                f"Erro ao tentar obter o e-mail do Exchange para {sender_name}: {e}")

                    # Obtém a data de recebimento
                    received_time = message.ReceivedTime.strftime(
                        "%Y-%m-%d %H:%M:%S")
                    emails_data.append(
                        [sender_name, sender_email, subject, received_time])
            except Exception as e:
                print(f"Erro ao processar mensagem: {e}")

    # Garante que a pasta de saída exista
    os.makedirs(output_folder, exist_ok=True)

    # Define o caminho do arquivo CSV
    output_file = os.path.join(output_folder, "emails_manual.csv")

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

    # Intervalo de datas fixo (única execução)
    start_date = datetime.strptime("2025-01-05", "%Y-%m-%d")
    end_date = datetime.strptime("2025-01-06", "%Y-%m-%d")

    print(f"Iniciando extração de e-mails do período de {
          start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')}.")
    extract_emails_by_date_range(output_folder, start_date, end_date)
    print("Execução concluída.")
