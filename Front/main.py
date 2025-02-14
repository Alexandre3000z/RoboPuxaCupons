import customtkinter as ctk
from PIL import Image
from tkinter import messagebox
from Front.app_state import app_state

def openMainPage(lastPage):
    
    # Função para permitir que só uma caixa seja selecionada
    def checkbox_event(checkbox):
        if checkbox == "A":
            if checkbox_a.get() == 1:
                checkbox_b.deselect()
        elif checkbox == "B":
            if checkbox_b.get() == 1:
                checkbox_a.deselect()

    def save_and_close():
        """Captura os valores dos campos e armazena no estado global antes de fechar a janela."""
        selected_process = 1 if checkbox_a.get() == 1 else 2
        selected_cupom = selected_option.get()
        inscricao_estadual = Ie_entry.get()
        mes = Month_entry.get()
        ano = Year_entry.get()

        # Atualiza os dados globais
        app_state.set_data(selected_process, selected_cupom, inscricao_estadual, mes, ano)

        # Fecha a janela principal
        mainPage.destroy()
        
        
    # Fecha a janela de login
    lastPage.destroy()

    # Cria a nova janela principal
    mainPage = ctk.CTk()
    mainPage.title("Main Page")
    mainPage.geometry("1280x720")
    mainPage.configure(fg_color="#25412D")

    # Texto de boas-vindas
    welcomeLabel = ctk.CTkLabel(mainPage, 
                                text="AUTOMAÇÃO DE CUPONS", 
                                font=("Consolas", 38, "bold"), 
                                text_color="white")
    welcomeLabel.pack(pady=50)
    
   # Frame principal que segura os dois lados
    mainPage_frame = ctk.CTkFrame(mainPage, 
                                  width=1200, 
                                  height=600, 
                                  corner_radius=20, 
                                  fg_color='#25412D',
                                  )
    
    mainPage_frame.pack(fill="both", 
                        expand=True, 
                        padx=100, pady=(10,0))

#------------------------------------------------------------------#
#--------------------------LADO ESQUERDO---------------------------#
#------------------------------------------------------------------#

    # Frame da esquerda (não visível, mas organiza o conteúdo)
    MainLeft_frame = ctk.CTkFrame(mainPage_frame, 
                                  fg_color="transparent",
                                  
                                  )
    
    MainLeft_frame.pack(side="left", 
                        fill="both", 
                        expand=True,
                        padx=20, 
                        pady=20)

#-------------------- CHECK BOXES --------------------#

    # Titulo
    IeTitleLabel = ctk.CTkLabel(MainLeft_frame, 
                                text="SELECIONE O TIPO DE PROCESSO:", 
                                font=("Consolas", 20, "bold"), 
                                text_color="white")
    
    IeTitleLabel.pack(anchor='w', pady=(0, 5))
    

    # CheckBoxes
    checkbox_a = ctk.CTkCheckBox(MainLeft_frame, 
                                 text_color='white',
                                 font=("Consolas", 18, "bold"), 
                                 text="Processo Automático (Com procuração)", 
                                 hover_color='white',
                                 border_color='white',
                                 command=lambda: checkbox_event("A"))
    
    checkbox_b = ctk.CTkCheckBox(MainLeft_frame, 
                                 text_color='white',
                                 font=("Consolas", 18, "bold"),
                                 text="Processo Manual (Sem procuração)", 
                                 hover_color='white',
                                 border_color='white',
                                 command=lambda: checkbox_event("B"))
    
    # Posicionar as CheckBoxes
    checkbox_a.pack(anchor='w', pady=(20,10), padx=20)
    checkbox_b.pack(anchor='w', pady=10, padx=20)
    
#-------------- CAMPO DE INSCRIÇÃO ESTADUAL --------------#

    # Titulo de Inscrição
    IeLabel = ctk.CTkLabel(MainLeft_frame,
                           text="INSCRIÇÃO ESTADUAL:",
                           font=("Consolas", 20, "bold"), 
                           text_color="white")
    
    IeLabel.pack(anchor='w', pady=(30, 5))
    
    # Campo para inscrição
    Ie_entry = ctk.CTkEntry(MainLeft_frame, 
                            text_color='black',
                            width=175, 
                            font=("Consolas", 18, "bold"), 
                            border_color='white')
    
    Ie_entry.pack(anchor='w', pady=0, padx=20)

#-------------- FRAME PARA ALINHAR MÊS E ANO ----------------#

    date_frame = ctk.CTkFrame(MainLeft_frame, fg_color="transparent")
    date_frame.pack(anchor='w', pady=(30, 5))

    #-------------- CAMPO MÊS --------------# 
   
    # Frame individual para o Mês
    month_frame = ctk.CTkFrame(date_frame, fg_color="transparent")
    month_frame.pack(side="left", padx=0)
    
    # Título Mês
    MonthLabel = ctk.CTkLabel(month_frame, 
                              text="MÊS:", 
                              font=("Consolas", 20, "bold"), 
                              text_color="white")
    
    MonthLabel.pack(anchor='w')
    
    # Campo para Mês
    Month_entry = ctk.CTkEntry(month_frame, 
                               text_color='black',
                               width=100, 
                               font=("Consolas", 18, "bold"), 
                               border_color='white',)
    
    Month_entry.pack(anchor='w', pady=5, padx=20)
    
    #-------------- CAMPO ANO --------------# 
    
    # Frame individual para o Ano
    year_frame = ctk.CTkFrame(date_frame, fg_color="transparent")
    year_frame.pack(side="left", padx=0)

    # Título Ano
    YearLabel = ctk.CTkLabel(year_frame, 
                             text="ANO:", 
                             font=("Consolas", 20, "bold"), 
                             text_color="white")
    
    YearLabel.pack(anchor='w')
    
    # Campo para Ano
    Year_entry = ctk.CTkEntry(year_frame, 
                              text_color='black',
                              width=100, 
                              font=("Consolas", 18, "bold"), 
                              border_color='white')
    
    Year_entry.pack(anchor='w', pady=5, padx=20)

#------------------------------------------------------------------#
#---------------------------LADO DIREITO---------------------------#
#------------------------------------------------------------------#

       
    # Frame da direita (não visível, mas organiza o conteúdo)
    MainRight_frame = ctk.CTkFrame(mainPage_frame, 
                                   fg_color="transparent",
                                   )
    
    MainRight_frame.pack(side="right", 
                         fill="both", 
                         expand=True, 
                         padx=20, pady=20)
    
#-----------------------OPÇÕES DE CUPONS-------------------------------#
    
    # Titulo
    OptionsTittleLabel = ctk.CTkLabel(MainRight_frame, 
                                text="SELECIONE OS TIPOS DE CUPONS DESEJADOS:", 
                                font=("Consolas", 20, "bold"), 
                                text_color="white")
    
    OptionsTittleLabel.pack(anchor='w', pady=(0, 5))
    
    #----------CheckBoxes-----------#
    
    #Modelo CheckBoxes
    def CheckBoxModel(text, checkboxes, var):
        modelo = ctk.CTkCheckBox(MainRight_frame, 
                             text_color='white',
                             font=("Consolas", 18, "bold"), 
                             text=text, 
                             hover_color='white',
                             border_color='white',
                             command=lambda: checkbox_event2(modelo, checkboxes, var))
        return modelo
    
    # Função para garantir que apenas uma CheckBox esteja selecionada
    def checkbox_event2(checkbox, checkboxes, var):
        for cb in checkboxes:
            if cb != checkbox:
                cb.deselect()
        var.set(checkbox.cget("text"))  # Atualiza a variável com o nome do item selecionado
    
    
    # Variável para armazenar a opção selecionada
    selected_option = ctk.StringVar(value="")
    
    # Lista para armazenar todas as checkboxes
    checkboxes = []
        
    # Criar CheckBoxes
    checkbox_1 = CheckBoxModel('Todos os cupons', checkboxes, selected_option)
    checkbox_2 = CheckBoxModel('Autorizados e Cancelados', checkboxes, selected_option)
    checkbox_3 = CheckBoxModel('Cancelados e Cancelamentos', checkboxes, selected_option)
    checkbox_4 = CheckBoxModel('Autorizados e Cancelamentos', checkboxes, selected_option)
    checkbox_5 = CheckBoxModel('Autorizados', checkboxes, selected_option)
    checkbox_6 = CheckBoxModel('Cancelados', checkboxes, selected_option)
    checkbox_7 = CheckBoxModel('Cancelamentos', checkboxes, selected_option)
    
    # Adicionar todas as checkboxes à lista para controle
    checkboxes.extend([checkbox_1, checkbox_2, checkbox_3, checkbox_4, checkbox_5, checkbox_6, checkbox_7])

    # Posicionar as CheckBoxes
    for checkbox in checkboxes:
        checkbox.pack(anchor='w', pady=10, padx=20)
    
    
    # Botão para sair do app
    MainexitButton = ctk.CTkButton(mainPage,
                                   font=("Consolas", 20, "bold"), 
                                   text_color='black', 
                                   text="EXECUTAR", 
                                   fg_color='white', 
                                   command=save_and_close)
    
    MainexitButton.pack(pady=0, side='bottom' , expand=True)

    # Executa a nova janela
    mainPage.mainloop()


