import schedule
import time
import os
import psutil
import subprocess
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


def send_execution_logs():
    """Envia as últimas 6 linhas do terminal como e-mail."""
    sender_email = "guilherme.meijomil@sqltech.com.br"
    sender_password = "231297Gui@"
    recipient_email = "guilherme.meijomil@sqltech.com.br"

    subject = "Logs de execução do Scheduler"
    log_file_path = r"C:\Users\guilherme.meijomil\Documents\ExtractEmails\execution_log.txt"

    try:
        if os.path.exists(log_file_path):
            try:
                with open(log_file_path, "r", encoding="utf-8") as log_file:
                    lines = log_file.readlines()
            except UnicodeDecodeError:
                with open(log_file_path, "r", encoding="latin-1") as log_file:
                    lines = log_file.readlines()

            last_lines = lines[-6:] if len(lines) >= 6 else lines
            log_content = "".join(last_lines)
        else:
            log_content = "Arquivo de log não encontrado."

        body = f"As últimas 6 linhas do log do Scheduler são:\n\n{log_content}"
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        with smtplib.SMTP('smtp.office365.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)

        print("E-mail de logs enviado com sucesso.")
    except Exception as e:
        print(f"Erro ao enviar os logs por e-mail: {e}")


def run_script():
    """Executa o script app.py e registra logs corretamente."""
    script_path = r"C:\Users\guilherme.meijomil\Documents\ExtractEmails\app.py"
    log_file_path = r"C:\Users\guilherme.meijomil\Documents\ExtractEmails\execution_log.txt"

    for proc in psutil.process_iter(['cmdline']):
        try:
            if proc.info['cmdline'] and "app.py" in proc.info['cmdline']:
                print("app.py já está em execução.")
                return
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

    try:
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_message = f"Iniciando execução do script às {current_time}...\n"

        with open(log_file_path, "a", encoding="utf-8-sig") as log_file:
            log_file.write(log_message)

        result = subprocess.run(
            ["python", script_path], capture_output=True, text=True, encoding="utf-8")

        log_message += result.stdout + "\n" + result.stderr
        log_message += f"Script concluído às {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.\n"

        with open(log_file_path, "a", encoding="utf-8-sig") as log_file:
            log_file.write(log_message)
    except Exception as e:
        with open(log_file_path, "a", encoding="utf-8-sig") as log_file:
            log_file.write(f"Erro ao executar o script: {e}\n")


def display_schedule():
    current_time = datetime.now().strftime('%H:%M:%S')
    next_run_time = (datetime.now().replace(minute=0, second=0,
                     microsecond=0) + timedelta(hours=1)).strftime('%H:%M:%S')
    print(
        f"Agendador iniciado ({current_time}). Próxima execução: ({next_run_time}).")


schedule.every().hour.at(":00").do(run_script)
for day in ["monday", "tuesday", "wednesday", "thursday", "friday"]:
    for hour in ["10:00", "15:00", "19:00"]:
        getattr(schedule.every(), day).at(hour).do(send_execution_logs)

display_schedule()
while True:
    schedule.run_pending()
    time.sleep(1)
