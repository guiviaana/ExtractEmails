import schedule
import time
import os
import psutil
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

    # Caminho do arquivo de log (se houver um arquivo específico, atualize aqui)
    log_file_path = r"C:\Users\guilherme.meijomil\Documents\ExtractEmails\execution_log.txt"

    try:
        # Lê as últimas 6 linhas do arquivo de log
        if os.path.exists(log_file_path):
            with open(log_file_path, "r", encoding="utf-8") as log_file:
                lines = log_file.readlines()
                last_lines = lines[-6:] if len(lines) >= 6 else lines
                log_content = "".join(last_lines)
        else:
            log_content = "Arquivo de log não encontrado."

        # Corpo do e-mail com os logs
        body = f"As últimas 6 linhas do log do Scheduler são:\n\n{log_content}"

        # Configuração da mensagem
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # Configuração do servidor SMTP
        with smtplib.SMTP('smtp.office365.com', 587) as server:
            server.starttls()  # Ativa a segurança TLS
            server.login(sender_email, sender_password)  # Realiza o login
            server.send_message(msg)  # Envia a mensagem

        print("E-mail de logs enviado com sucesso.")
    except Exception as e:
        print(f"Erro ao enviar os logs por e-mail: {e}")


def run_script():
    """Executa o script app.py localizado no diretório especificado."""
    script_name = 'app.py'
    log_file_path = r"C:\Users\guilherme.meijomil\Documents\ExtractEmails\execution_log.txt"

    # Verifica se o script já está em execução
    for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            if proc.info['cmdline'] and script_name in proc.info['cmdline']:
                print(f"{script_name} já está em execução.")
                return
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

    try:
        # Captura o horário atual de execução
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_message = f"Iniciando execução do script {
            script_name} às {current_time}...\n"

        # Executa o script e redireciona a saída para o log
        os.system(
            f"python C:\\Users\\guilherme.meijomil\\Documents\\ExtractEmails\\app.py >> {log_file_path} 2>&1")

        log_message += f"Script {
            script_name} concluído com sucesso às {current_time}.\n"

        # Adiciona a mensagem ao arquivo de log
        with open(log_file_path, "a", encoding="utf-8") as log_file:
            log_file.write(log_message)
    except Exception as e:
        error_message = f"Erro ao executar o script: {e}\n"
        with open(log_file_path, "a", encoding="utf-8") as log_file:
            log_file.write(error_message)


def display_schedule():
    current_time = datetime.now().strftime('%H:%M:%S')
    next_run_time = (datetime.now().replace(minute=0, second=0,
                     microsecond=0) + timedelta(hours=1)).strftime('%H:%M:%S')
    print(f"Agendador iniciado ({
          current_time}). Próxima execução: ({next_run_time}).")


# Agenda a execução do script a cada 1 hora
schedule.every().hour.at(":00").do(run_script)

# Agenda o envio dos logs de execução às 10h, 15h e 19h durante os dias úteis
schedule.every().monday.at("10:00").do(send_execution_logs)
schedule.every().monday.at("15:00").do(send_execution_logs)
schedule.every().monday.at("19:00").do(send_execution_logs)

schedule.every().tuesday.at("10:00").do(send_execution_logs)
schedule.every().tuesday.at("15:00").do(send_execution_logs)
schedule.every().tuesday.at("19:00").do(send_execution_logs)

schedule.every().wednesday.at("10:00").do(send_execution_logs)
schedule.every().wednesday.at("15:00").do(send_execution_logs)
schedule.every().wednesday.at("19:00").do(send_execution_logs)

schedule.every().thursday.at("10:00").do(send_execution_logs)
schedule.every().thursday.at("15:00").do(send_execution_logs)
schedule.every().thursday.at("19:00").do(send_execution_logs)

schedule.every().friday.at("10:00").do(send_execution_logs)
schedule.every().friday.at("15:00").do(send_execution_logs)
schedule.every().friday.at("19:00").do(send_execution_logs)

display_schedule()

while True:
    schedule.run_pending()
    time.sleep(1)
