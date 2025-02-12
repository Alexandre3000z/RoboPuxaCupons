import customtkinter as ctk
from tkinter import messagebox

# Dados de autenticação (usuário e senha)
USUARIO_CORRETO = "admin"
SENHA_CORRETA = "1234"

# Função para verificar o login
def verificar_login():
    usuario = email_entry.get()
    senha = senha_entry.get()
    
    if usuario == USUARIO_CORRETO and senha == SENHA_CORRETA:
        abrir_pagina_principal()
    else:
        messagebox.showerror("Erro de Login", "Usuário ou senha incorretos!")

# Função para abrir a página principal e fechar a janela de login
def abrir_pagina_principal():
    # Fecha a janela de login
    root.destroy()

    # Cria a nova janela principal
    pagina_principal = ctk.CTk()
    pagina_principal.title("Página Principal")
    pagina_principal.geometry("800x600")
    pagina_principal.configure(fg_color="#f5f5f5")

    # Texto de boas-vindas
    boas_vindas = ctk.CTkLabel(pagina_principal, text="Bem-vindo à Página Principal!", 
                               font=("Arial", 24, "bold"), text_color="#1e3d2f")
    boas_vindas.pack(pady=50)

    # Botão para sair do app
    sair_button = ctk.CTkButton(pagina_principal, text="Sair", command=pagina_principal.quit)
    sair_button.pack(pady=20)

    # Executa a nova janela
    pagina_principal.mainloop()

# Configurações da Janela de Login
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("green")

root = ctk.CTk()
root.title("Login")
root.geometry("400x300")
root.configure(fg_color="#f5f5f5")

# Frame de login
login_frame = ctk.CTkFrame(root, width=350, height=250, corner_radius=20, fg_color="#1e3d2f")
login_frame.pack(pady=20)
login_frame.pack_propagate(False)

# Título de Login
login_label = ctk.CTkLabel(login_frame, text="Login", font=("Arial", 24, "bold"), text_color="white")
login_label.pack(pady=(20, 10))

# Campo de Usuário
email_entry = ctk.CTkEntry(login_frame, placeholder_text="Usuário", width=250, height=40, corner_radius=10,
                           fg_color="#edf6ff", text_color="black", placeholder_text_color="#a0a0a0")
email_entry.pack(pady=10)

# Campo de Senha
senha_entry = ctk.CTkEntry(login_frame, placeholder_text="Senha", show="*", width=250, height=40, corner_radius=10,
                           fg_color="#edf6ff", text_color="black", placeholder_text_color="#a0a0a0")
senha_entry.pack(pady=10)

# Botão de Login
login_button = ctk.CTkButton(login_frame, text="Entrar", width=150, height=40, corner_radius=10, fg_color="#ffffff",
                             text_color="#1e3d2f", hover_color="#dcdcdc", command=verificar_login)
login_button.pack(pady=20)

# Executa a janela de login
root.mainloop()
