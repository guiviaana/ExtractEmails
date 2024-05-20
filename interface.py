import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from extrair_emails_subpasta import export_emails_to_csv

def export_emails():
    email_address = email_entry.get()
    subfolder_name = subfolder_entry.get()
    start_date = start_date_entry.get()
    end_date = end_date_entry.get()

    # Conversão das datas para o formato datetime
    try:
        start_date = datetime.strptime(start_date, '%Y-%m-%d')
        end_date = datetime.strptime(end_date, '%Y-%m-%d')
    except ValueError:
        messagebox.showerror("Erro", "Formato de data inválido. Use o formato YYYY-MM-DD.")
        return

    # Chama a função para exportar e-mails
    export_emails_to_csv(email_address, subfolder_name, start_date, end_date)
    messagebox.showinfo("Sucesso", f"E-mails exportados da subpasta '{subfolder_name}' dentro da caixa de entrada para emails.csv com sucesso!")

# Criação da janela principal
window = tk.Tk()
window.title("Exportar E-mails")
window.geometry("470x200")  # Tamanho da janela

# Estilo dos widgets
style = ttk.Style(window)
style.theme_use("clam")  # Escolha de tema (pode ser alterado para "winnative", "vista", "xpnative", etc.)

# Componentes da interface
label_font = ('Helvetica', 12)
entry_font = ('Helvetica', 12)
button_font = ('Helvetica', 12, 'bold')

# Configuração dos rótulos
tk.Label(window, text="Endereço de E-mail:", font=label_font).grid(row=0, column=0, sticky="w", padx=10, pady=5)
tk.Label(window, text="Subpasta:", font=label_font).grid(row=1, column=0, sticky="w", padx=10, pady=5)
tk.Label(window, text="Data de Início (YYYY-MM-DD):", font=label_font).grid(row=2, column=0, sticky="w", padx=10, pady=5)
tk.Label(window, text="Data de Término (YYYY-MM-DD):", font=label_font).grid(row=3, column=0, sticky="w", padx=10, pady=5)

# Configuração das entradas
email_entry = tk.Entry(window, font=entry_font)
email_entry.grid(row=0, column=1, padx=10, pady=5)
subfolder_entry = tk.Entry(window, font=entry_font)
subfolder_entry.grid(row=1, column=1, padx=10, pady=5)
start_date_entry = tk.Entry(window, font=entry_font)
start_date_entry.grid(row=2, column=1, padx=10, pady=5)
end_date_entry = tk.Entry(window, font=entry_font)
end_date_entry.grid(row=3, column=1, padx=10, pady=5)

# Configuração do botão de exportação
export_button = tk.Button(window, text="Exportar", font=button_font, command=export_emails)
export_button.grid(row=4, column=0, columnspan=2, pady=10)

# Inicia o loop da interface gráfica
window.mainloop()
