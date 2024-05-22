from flask import Flask, render_template, request, redirect, url_for, flash, make_response
from datetime import datetime, timedelta
import win32com.client
import csv
import re
import pythoncom
import io

app = Flask(__name__)
app.secret_key = 'some_secret_key'

def clean_sender_email(sender_email):
    exchange_regex = re.compile(r"/O=[^/]+/OU=[^/]+/CN=RECIPIENTS/CN=[^@]+@[^@]+")
    if exchange_regex.match(sender_email):
        return sender_email.split("@")[-1]
    return sender_email

def export_emails_to_csv(email_address, subfolder_name=None, start_date=None, end_date=None):
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except Exception as e:
        print(f"Erro ao inicializar Outlook: {e}")
        flash(f"Erro ao inicializar Outlook: {e}", "error")
        return None

    try:
        for account in outlook.Folders:
            if account.Name.lower() == email_address.lower():
                inbox_folder = account.Folders("Caixa de Entrada")
                if subfolder_name:
                    subfolder = find_subfolder(inbox_folder, subfolder_name)
                    if subfolder is None:
                        flash(f"A subpasta '{subfolder_name}' não foi encontrada na caixa de entrada da conta de e-mail {email_address}.", "error")
                        return None
                    inbox_folder = subfolder

                if start_date and end_date:
                    end_date += timedelta(days=1)  # Inclui o dia final completo
                    filter_str = "[ReceivedTime] >= '{}' AND [ReceivedTime] < '{}'".format(start_date.strftime('%m/%d/%Y'), end_date.strftime('%m/%d/%Y'))
                    filtered_emails = inbox_folder.Items.Restrict(filter_str)
                else:
                    filtered_emails = inbox_folder.Items

                csv_content = io.StringIO()
                csv_writer = csv.writer(csv_content)
                csv_writer.writerow(['Assunto', 'Nome do Remetente', 'Endereço do Remetente', 'Data e Hora'])

                for email in filtered_emails:
                    sender_name = email.SenderName
                    sender_email = clean_sender_email(email.SenderEmailAddress)
                    csv_writer.writerow([email.Subject, sender_name, sender_email, email.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")])

                response = make_response(csv_content.getvalue())
                response.headers['Content-Type'] = 'text/csv'
                response.headers['Content-Disposition'] = 'attachment; filename=emails.csv'
                return response

        flash(f"A conta de e-mail {email_address} não foi encontrada no Outlook.", "error")
        return None

    except Exception as e:
        print(f"Erro ao acessar a conta de e-mail ou exportar e-mails: {e}")
        flash(f"Erro ao acessar a conta de e-mail ou exportar e-mails: {e}", "error")
        return None

def find_subfolder(folder, subfolder_name):
    for subfolder in folder.Folders:
        if subfolder.Name.lower() == subfolder_name.lower():
            return subfolder
    return None

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        email_address = request.form['email_address']
        subfolder_name = request.form['subfolder_name']
        start_date = request.form['start_date']
        end_date = request.form['end_date']

        if not email_address:
            flash("O campo 'Endereço de E-mail' é obrigatório.", "error")
            return redirect(url_for('index'))

        try:
            start_date = datetime.strptime(start_date, '%Y-%m-%d')
            end_date = datetime.strptime(end_date, '%Y-%m-%d')
        except ValueError:
            flash("Formato de data inválido. Use o formato YYYY-MM-DD.", "error")
            return redirect(url_for('index'))

        response = export_emails_to_csv(email_address, subfolder_name, start_date, end_date)
        if response:
            return response
        else:
            return redirect(url_for('index'))

    return render_template('index.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
