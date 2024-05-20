import win32com.client
import csv
from datetime import datetime
import re

def find_subfolder(root_folder, subfolder_name):
    for folder in root_folder.Folders:
        if folder.Name.lower() == subfolder_name.lower():
            return folder
        # Se a pasta atual tiver subpastas, verifica recursivamente
        if folder.Folders.Count > 0:
            subfolder = find_subfolder(folder, subfolder_name)
            if subfolder:
                return subfolder
    return None

def clean_sender_email(sender_email):
    # Verifica se o remetente é um endereço interno do Exchange Server
    exchange_regex = re.compile(r"/O=[^/]+/OU=[^/]+/CN=RECIPIENTS/CN=[^@]+@[^@]+")
    if exchange_regex.match(sender_email):
        # Extrai apenas o endereço de e-mail visível
        return sender_email.split("@")[-1]
    return sender_email

def export_emails_to_csv(email_address, subfolder_name=None, start_date=None, end_date=None):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Acessa todas as pastas de todas as contas configuradas no Outlook
    for account in outlook.Folders:
        # Verifica se a conta corresponde ao endereço de e-mail especificado
        if account.Name.lower() == email_address.lower():
            # Acessa a caixa de entrada
            inbox_folder = account.Folders("Caixa de Entrada")
            
            # Se subfolder_name for fornecido, procura a subpasta desejada
            if subfolder_name:
                subfolder = find_subfolder(inbox_folder, subfolder_name)
                if subfolder is None:
                    print(f"A subpasta '{subfolder_name}' não foi encontrada na caixa de entrada da conta de e-mail {email_address}.")
                    return
                inbox_folder = subfolder

            # Define o filtro de data para a busca
            if start_date and end_date:
                filter_str = "[ReceivedTime] >= '{}' AND [ReceivedTime] <= '{}'".format(start_date.strftime('%m/%d/%Y'), end_date.strftime('%m/%d/%Y'))
                filtered_emails = inbox_folder.Items.Restrict(filter_str)
            else:
                filtered_emails = inbox_folder.Items

            # Cria um arquivo CSV para salvar os dados
            with open('emails.csv', 'w', newline='', encoding='utf-8') as csvfile:
                csv_writer = csv.writer(csvfile)
                csv_writer.writerow(['Assunto', 'Nome do Remetente', 'Endereço do Remetente', 'Data e Hora'])

                # Itera sobre os e-mails filtrados e grava as informações no CSV
                for email in filtered_emails:
                    sender_name = email.SenderName
                    sender_email = clean_sender_email(email.SenderEmailAddress)
                    csv_writer.writerow([email.Subject, sender_name, sender_email, email.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")])

            if subfolder_name:
                print(f"E-mails exportados da subpasta '{subfolder_name}' dentro da caixa de entrada para emails.csv com sucesso!")
            else:
                print("E-mails exportados da caixa de entrada para emails.csv com sucesso!")
            return

    print(f"A conta de e-mail {email_address} não foi encontrada no Outlook.")

# Exemplo de uso
if __name__ == "__main__":
    email_address = "guilherme.meijomil@sqltech.com.br" # Email que será buscado
    subfolder_name = "Teste"  # Nome da subpasta dentro da caixa de entrada (opcional)
    start_date = datetime(2024, 3, 16)  # Data de início (opcional)
    end_date = datetime(2024, 3, 18)   # Data de término (opcional)

    export_emails_to_csv(email_address, subfolder_name, start_date, end_date)
