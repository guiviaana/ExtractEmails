import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def send_test_email():
    """Envia um e-mail de teste."""
    from_email = "guilherme.meijomil@sqltech.com.br"  # Seu e-mail
    password = "231297Gui@"  # Sua senha
    to_email = "guilherme.meijomil@sqltech.com.br"  # Destinatário do e-mail

    # Configuração da mensagem
    subject = "Teste de Envio de E-mail"
    body = "Este é um e-mail de teste para verificar o envio via Python."

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        # Configuração do servidor SMTP
        with smtplib.SMTP('smtp.office365.com', 587) as server:
            server.starttls()  # Segurança (TLS)
            server.login(from_email, password)  # Login no servidor
            server.send_message(msg)  # Envia o e-mail
        print(f"E-mail de teste enviado com sucesso para {to_email}.")
    except Exception as e:
        print(f"Erro ao enviar o e-mail: {e}")

# Executa a função de envio
if __name__ == "__main__":
    send_test_email()
