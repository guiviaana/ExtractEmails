# Importa as bibliotecas necessárias
from flask import Flask, render_template, request, redirect, url_for, flash, make_response
from datetime import datetime, timedelta
import win32com.client  # Biblioteca para acessar o Outlook
import csv  # Biblioteca para manipular arquivos CSV
import re  # Biblioteca para expressões regulares
import pythoncom  # Biblioteca para inicializar COM para acesso ao Outlook
import io  # Biblioteca para operações de entrada/saída

# Cria uma instância do aplicativo Flask
app = Flask(__name__)
app.secret_key = 'some_secret_key'  # Chave secreta usada para gerenciar a sessão no Flask

def clean_sender_email(sender_email):
    # Função que limpa o e-mail do remetente para uma forma mais legível
    exchange_regex = re.compile(r"/O=[^/]+/OU=[^/]+/CN=RECIPIENTS/CN=[^@]+@[^@]+")
    if exchange_regex.match(sender_email):
        # Se o e-mail do remetente estiver no formato Exchange, simplifica o endereço
        return sender_email.split("@")[-1]
    return sender_email

def export_emails_to_csv(email_address, subfolder_name=None, start_date=None, end_date=None):
    # Inicializa a biblioteca COM para uso do Outlook
    pythoncom.CoInitialize()

    # Acessa o namespace do Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Procura a conta de e-mail especificada
    for account in outlook.Folders:
        if account.Name.lower() == email_address.lower():
            # Acessa a pasta "Caixa de Entrada" da conta de e-mail
            inbox_folder = account.Folders("Caixa de Entrada")
            
            if subfolder_name:
                # Se o nome da subpasta foi fornecido, procura a subpasta especificada
                subfolder = find_subfolder(inbox_folder, subfolder_name)
                if subfolder is None:
                    # Se a subpasta não for encontrada, exibe uma mensagem de erro
                    flash(f"A subpasta '{subfolder_name}' não foi encontrada na caixa de entrada da conta de e-mail {email_address}.", "error")
                    return None
                inbox_folder = subfolder

            if start_date and end_date:
                # Se as datas de início e término forem fornecidas, filtra os e-mails por data
                end_date = end_date + timedelta(days=1)  # Inclui o dia final no filtro
                filter_str = "[ReceivedTime] >= '{}' AND [ReceivedTime] < '{}'".format(
                    start_date.strftime('%m/%d/%Y'),
                    end_date.strftime('%m/%d/%Y')
                )
                filtered_emails = inbox_folder.Items.Restrict(filter_str)
            else:
                # Se não, seleciona todos os e-mails da pasta
                filtered_emails = inbox_folder.Items

            # Cria um objeto StringIO para armazenar o conteúdo CSV
            csv_content = io.StringIO()
            # Cria um escritor CSV
            csv_writer = csv.writer(csv_content)
            # Escreve o cabeçalho do CSV
            csv_writer.writerow(['Assunto', 'Nome do Remetente', 'Endereço do Remetente', 'Data e Hora'])
            for email in filtered_emails:
                # Para cada e-mail, obtém os detalhes e os escreve no CSV
                sender_name = email.SenderName
                sender_email = clean_sender_email(email.SenderEmailAddress)
                csv_writer.writerow([email.Subject, sender_name, sender_email, email.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")])

            # Cria uma resposta HTTP com o conteúdo CSV
            response = make_response(csv_content.getvalue())
            response.headers['Content-Type'] = 'text/csv'
            response.headers['Content-Disposition'] = 'attachment; filename=emails.csv'
            return response

    # Se a conta de e-mail não for encontrada, exibe uma mensagem de erro
    flash(f"A conta de e-mail {email_address} não foi encontrada no Outlook.", "error")
    return None

def find_subfolder(folder, subfolder_name):
    # Função que encontra uma subpasta específica na pasta fornecida
    for subfolder in folder.Folders:
        if subfolder.Name.lower() == subfolder_name.lower():
            return subfolder
    return None

@app.route('/', methods=['GET', 'POST'])
def index():
    # Rota principal que trata requisições GET e POST
    if request.method == 'POST':
        # Se a requisição for POST, obtém os dados do formulário
        email_address = request.form['email_address']
        subfolder_name = request.form['subfolder_name']
        start_date_str = request.form['start_date']
        end_date_str = request.form['end_date']

        try:
            # Converte as datas de string para objetos datetime
            start_date = datetime.strptime(start_date_str, '%d-%m-%Y')
            end_date = datetime.strptime(end_date_str, '%d-%m-%Y')
        except ValueError:
            # Se a conversão falhar, exibe uma mensagem de erro
            flash("Formato de data inválido. Use o formato DD-MM-YYYY.", "error")
            return redirect(url_for('index'))

        # Exporta os e-mails para um arquivo CSV
        response = export_emails_to_csv(email_address, subfolder_name, start_date, end_date)
        if response:
            return response

    # Renderiza o template HTML para a página principal
    return render_template('index.html')

if __name__ == '__main__':
    # Inicia o servidor Flask
    app.run(host='0.0.0.0', port=5000)
