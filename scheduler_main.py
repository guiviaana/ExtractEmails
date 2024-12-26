import schedule
import time
import os
import psutil
from datetime import datetime, timedelta


def run_script():
    """Executa o script app.py localizado no diretório especificado."""
    script_name = 'app.py'

    # Verifica se o script já está em execução
    for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            # Verifica se o processo possui cmdline e se contém o script_name
            if proc.info['cmdline'] and script_name in proc.info['cmdline']:
                print(f"{script_name} já está em execução.")
                return
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            # Ignora processos que já terminaram ou que não podemos acessar
            pass

    # Se não encontrar o processo em execução, executa o script
    try:
        # Captura o horário atual de execução
        current_time = datetime.now().strftime('%H:%M:%S')
        print(f"Iniciando execução do script {
              script_name} às {current_time}...")

        # Executa o script
        os.system(
            r"python C:\Users\guilherme.meijomil\Documents\ExtractEmails\app.py")

        # Após a execução, imprime o horário de conclusão
        print(f"Script {script_name} concluído com sucesso às {current_time}.")
    except Exception as e:
        print(f"Erro ao executar o script: {e}")

# Função para exibir os horários


def display_schedule():
    # Captura o horário atual
    current_time = datetime.now().strftime('%H:%M:%S')

    # Captura o próximo horário cheio (ex: 12:00, 13:00, 14:00)
    next_run_time = (datetime.now().replace(minute=0, second=0,
                     microsecond=0) + timedelta(hours=1)).strftime('%H:%M:%S')

    print(f"Agendador de execução iniciado ({
          current_time}). O script app.py será executado a cada 1 hora, próxima execução: ({next_run_time}).")


# Agenda a execução a cada 1 hora, nos horários cheios (ex: 12:00, 13:00, 14:00, etc)
schedule.every().hour.at(":00").do(run_script)

# Exibe a mensagem inicial com o horário de início e a próxima execução
display_schedule()

while True:
    # Verifica e executa tarefas pendentes
    schedule.run_pending()
    time.sleep(1)
