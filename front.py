import customtkinter as ctk
from PIL import Image
from tkinter import messagebox

# Dados de autenticação (usuário e senha)
USUARIO_CORRETO = "admin"
SENHA_CORRETA = "1234"

# Função para verificar o login
def checkLogin():
    usuario = email_entry.get()
    senha = senha_entry.get()
    
    if usuario == USUARIO_CORRETO and senha == SENHA_CORRETA:
        openMainPage()
    else:
        messagebox.showerror("Erro de Login", "Usuário ou senha incorretos!")


def openMainPage():
    # Fecha a janela de login
    root.destroy()

    # Cria a nova janela principal
    mainPage = ctk.CTk()
    mainPage.title("Main Page")
    mainPage.geometry("1280x720")
    mainPage.configure(fg_color="#25412D")

    # Texto de boas-vindas
    welcomeLabel = ctk.CTkLabel(mainPage, text="Automação de Cupons", 
                                font=("Arial", 38, "bold"), text_color="white")
    welcomeLabel.pack(pady=50)
    
   # Frame principal que segura os dois lados
    mainPage_frame = ctk.CTkFrame(mainPage, width=1200, height=600, corner_radius=20, fg_color='#25412D')
    mainPage_frame.pack(fill="both", expand=True, padx=20, pady=20)

    # Frame da esquerda (não visível, mas organiza o conteúdo)
    MainLeft_frame = ctk.CTkFrame(mainPage_frame, fg_color="transparent")
    MainLeft_frame.pack(side="left", fill="both", expand=True, padx=20, pady=120)

    # Frame da direita (não visível, mas organiza o conteúdo)
    MainRight_frame = ctk.CTkFrame(mainPage_frame, fg_color="transparent")
    MainRight_frame.pack(side="right", fill="both", expand=True, padx=20, pady=20)
    # Botão para sair do app
    MainexitButton = ctk.CTkButton(mainPage, text="Exit", command=mainPage.quit)
    MainexitButton.pack(pady=20)

    # Executa a nova janela
    mainPage.mainloop()




#------------------PAGINA DE LOGIN------------------#


# Configurações iniciais
ctk.set_appearance_mode("light")  # Modo claro
ctk.set_default_color_theme("green")  # Tema de cores (pode personalizar depois)

# Janela principal
root = ctk.CTk()
root.title("Login")
root.geometry("1280x720")
root.configure(fg_color="#f5f5f5")  # Cor de fundo clara

# Frame principal que segura os dois lados
main_frame = ctk.CTkFrame(root, width=1200, height=600, corner_radius=20)
main_frame.pack(fill="both", expand=True, padx=20, pady=20)

#Frame que segura os dois lados
login_frame = ctk.CTkFrame(main_frame, width=1000, height=500, corner_radius=20, fg_color="#1e3d2f")
login_frame.pack(side="left", fill="both", expand=True)
login_frame.pack_propagate(False)

# Frame da esquerda (não visível, mas organiza o conteúdo)
frame_esquerda = ctk.CTkFrame(login_frame, fg_color="transparent")
frame_esquerda.pack(side="left", fill="both", expand=True, padx=20, pady=120)

# Frame da direita (não visível, mas organiza o conteúdo)
frame_direita = ctk.CTkFrame(login_frame, fg_color="transparent")
frame_direita.pack(side="right", fill="both", expand=True, padx=20, pady=20)

#-----------LADO ESQUERDO-----------#

# Título de Login
login_label = ctk.CTkLabel(frame_esquerda, text="Login", font=("Arial", 28, "bold"), text_color="white")
login_label.pack(pady=(40, 10))


# Subtítulo
subtitle_label = ctk.CTkLabel(frame_esquerda, text="Use seu email e senha para acessar a plataforma", 
                              font=("Arial", 12), text_color="white")
subtitle_label.pack(pady=(0, 20))

# Campos de Entrada (Email e Senha)
email_entry = ctk.CTkEntry(frame_esquerda, placeholder_text="Email", width=300, height=40, corner_radius=10,
                           fg_color="#edf6ff", text_color="black", placeholder_text_color="#a0a0a0")
email_entry.pack(pady=10)

senha_entry = ctk.CTkEntry(frame_esquerda, placeholder_text="Senha", show="*", width=300, height=40, corner_radius=10,
                           fg_color="#edf6ff", text_color="black", placeholder_text_color="#a0a0a0")
senha_entry.pack(pady=10)

# Link "Esqueceu sua Senha?"
forgot_password = ctk.CTkLabel( frame_esquerda, text="Esqueceu sua Senha?", font=("Arial", 10), text_color="#e0e0e0",
                               cursor="hand2")
forgot_password.pack(pady=(5, 20))

# Botão de Login
login_button = ctk.CTkButton(frame_esquerda, text="Entrar", width=150, height=40, corner_radius=10, fg_color="#ffffff",
                             text_color="#1e3d2f", hover_color="#dcdcdc", command=checkLogin)
login_button.pack(pady=10)

#-----------LADO DIREITO(logo)-----------#
logo_frame = ctk.CTkFrame(frame_direita, width=550, height=500, corner_radius=100, fg_color="white")
logo_frame.pack(side="right", fill="both", expand=True)
logo_frame.pack_propagate(False)

try:
    logo_image = ctk.CTkImage(dark_image=Image.open("image.png"), size=(500,500 ))
    logo_label = ctk.CTkLabel(logo_frame, image=logo_image, text="")
    logo_label.pack(expand=True)
except:
    # Caso não tenha o logo, exibe texto como placeholder
    logo_label = ctk.CTkLabel(logo_frame, text="BUSINESS PRO\nCONTÁBIL", font=("Arial", 24, "bold"), text_color="#1e3d2f")
    logo_label.pack(expand=True)

# Inicia a aplicação
root.mainloop()
