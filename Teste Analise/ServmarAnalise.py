import customtkinter
import tkinter as tk
from tkinter import font
from PIL import Image
from tkinter import filedialog
import pandas as pd
import threading
import os

def Thread():
    thread = threading.Thread(target=analise)
    thread.start()

def carregar_excel():
    global file_path
    global novo_caminho
    file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
    if file_path:
        global tabela
        tabela = pd.read_excel(file_path)
        status_label.config(text="Arquivo Excel Carregado.", foreground="black")
        print("Arquivo Excel Carregado.")
        print("-" * 100)
        # Salve o DataFrame no mesmo arquivo Excel
        tabela.to_excel(file_path, index=False)

    print(file_path)

    dialog = customtkinter.CTkInputDialog(title="Caixa de dialogo", text="Digite o nome do novo arquivo:")
    linha_text = dialog.get_input()
    print(linha_text)

    # Substitua o nome do arquivo inteiro pelo valor de linha_text
    novo_caminho = file_path.replace(os.path.basename(file_path), f'{linha_text}.xlsx')
    tabela.to_excel(novo_caminho, index=False)
    
    print(f"Arquivo salvo em: {novo_caminho}")

    status_label.config(text="Pronto para rodar")
    janelaprincipal.update_idletasks()

def analise():
        if quantidade_analise == 3 and escolha == "Ceimic":
            status_label.config(text="Rodando Verificador")
            janelaprincipal.update_idletasks()
            import Ceimic
            Ceimic.main(file_path, novo_caminho)  # Passe o caminho do arquivo como argumento
            status_label.config(text="Rodando Organizador")
            janelaprincipal.update_idletasks()
            import Organizar
            Organizar.main(novo_caminho)  # Passe o caminho do arquivo como argumento
            status_label.config(text="Rodando Analise")
            janelaprincipal.update_idletasks()
            import Analise3
            Analise3.main(novo_caminho)
            status_label.config(text="Terminou a Analise")
            janelaprincipal.update_idletasks()

        elif quantidade_analise == 2 and escolha == "Ceimic":
            status_label.config(text="Rodando Verificador")
            janelaprincipal.update_idletasks()
            import Ceimic
            Ceimic.main(file_path, novo_caminho)  # Passe o caminho do arquivo como argumento
            status_label.config(text="Rodando Organizador")
            janelaprincipal.update_idletasks()
            import Organizar
            Organizar.main(novo_caminho)  # Passe o caminho do arquivo como argumento
            status_label.config(text="Rodando Analise")
            janelaprincipal.update_idletasks()
            import Analise2
            Analise2.main(novo_caminho)
            status_label.config(text="Terminou a Analise")
            janelaprincipal.update_idletasks()

def quantidade_VO():
    global quantidade_analise
    # Criar uma nova janela para os Radiobuttons
    nova_janela_radiobutton = customtkinter.CTkToplevel()
    nova_janela_radiobutton.geometry("235x175")
    nova_janela_radiobutton.title("SERVMAR")
    nova_janela_radiobutton.resizable(width=False, height=False)

    # Variável para armazenar a escolha do usuário
    escolha_var = tk.StringVar(value=1)
    # Função para obter a escolha do usuário e atribuir à variável valor_Primario
    def obter_escolha():
        global quantidade_analise
        escolha = escolha_var.get()

        if escolha == "2 Valores Orientadores":
            quantidade_analise = 2
        elif escolha == "3 Valores Orientadores":
            quantidade_analise = 3

        print("Valor selecionado:", quantidade_analise)

        nova_janela_radiobutton.destroy()
        return quantidade_analise

    opcoes = ["2 Valores Orientadores", "3 Valores Orientadores"]

    label = customtkinter.CTkLabel(nova_janela_radiobutton, text="Escolha a quantidade de\nValores Orientadores")
    label.pack(pady=10)  # Ajuste a distância vertical entre a label e os Radiobuttons

    for opcao in opcoes:
        customtkinter.CTkRadioButton(nova_janela_radiobutton, text=opcao, variable=escolha_var, value=opcao).pack(anchor='center', pady=5)  # Ajuste a distância vertical entre os Radiobuttons

    # Botão para confirmar a escolha
    botao_confirmar = customtkinter.CTkButton(nova_janela_radiobutton, text="Confirmar", command=obter_escolha)
    botao_confirmar.pack(pady=10, padx=10)  # Ajuste a distância vertical e horizontal entre os Radiobuttons e o botão de confirmar

    nova_janela_radiobutton.wait_window()

    Thread()

def lab():

        global escolha
        # Criar uma janela modal
        nova_janela = customtkinter.CTkToplevel()
        nova_janela.title("SERVMAR")
        nova_janela.geometry("235x175")
        nova_janela.grab_set()
        nova_janela.resizable(width=False, height=False)

        # Variável para armazenar a escolha do usuário
        ordem_planilhas = tk.StringVar()

        # Função para imprimir a opção escolhida no terminal
        def imprimir_escolha():
            global escolha
            escolha = ordem_planilhas.get()
            if escolha:
                print("Opção escolhida:", escolha)
                nova_janela.destroy()
            else:
                aviso_label.pack()


        # Adicionar Radiobuttons personalizados
        opcoes = ['Ceimic']

        label = customtkinter.CTkLabel(nova_janela, text="Escolha de qual laboratorio a\n Analise deve ser feita")
        label.pack(pady=10)  # Ajuste a distância vertical entre a label e os Radiobuttons

        for opcao in opcoes:
            customtkinter.CTkRadioButton(nova_janela, text=opcao, variable=ordem_planilhas, value=opcao).pack(anchor='center')

        # Label para mostrar a mensagem de aviso
        aviso_label = customtkinter.CTkLabel(nova_janela, text="Escolha uma opção!")
        aviso_label.pack_forget()  # Oculta a label inicialmente

        # Botão para confirmar a escolha
        botao_confirmar = customtkinter.CTkButton(nova_janela, text="Confirmar", command=imprimir_escolha)
        botao_confirmar.pack(pady=20, padx=10)

        nova_janela.wait_window()

        print("oi")
        print(escolha)

        quantidade_VO()

# Obtém o diretório atual do script
script_dir = os.path.dirname(os.path.abspath(__file__))

janelaprincipal = customtkinter.CTk()
janelaprincipal.geometry("345x200")
janelaprincipal.title("SERVMAR")
janelaprincipal.resizable(width=False, height=False)

customtkinter.set_appearance_mode("light")
customtkinter.set_default_color_theme("dark-blue")

# Usa caminhos relativos para o ícone
icone_path = os.path.join(script_dir, "servmarico.ico")
janelaprincipal.iconbitmap(icone_path)

# Evite conflito de nomes de botões
botao_carregar= customtkinter.CTkButton(janelaprincipal, text="Carregar Excel", command=carregar_excel).place(x=15, y=10)
botao_analise= customtkinter.CTkButton(janelaprincipal, text="Fazer Analise", command=lab).place(x=190, y=10)

status_label = tk.Label(janelaprincipal, text="", bg="#ebebeb")
status_label.place(relx=0.5, rely=0.3, anchor="center")
projeto = customtkinter.CTkLabel(janelaprincipal, text="Projeto Canhadas", font=("Helvetica", 25), text_color="blue")
# Configurando o posicionamento para centralizar horizontalmente
projeto.place(relx=0.5, rely=0.5, anchor="center")
fonte = font.nametofont("TkDefaultFont")
fonte.configure(underline=True)
fonte.configure(size=15)
status_label.config(font=fonte)

# Usa caminhos relativos para a imagem
imagem_path = os.path.join(script_dir, "servmarlogo.png")
img = customtkinter.CTkImage(light_image=Image.open(imagem_path), size=(345, 50))
imagem = customtkinter.CTkLabel(janelaprincipal, text="", image=img).place(x=5, y=120)

janelaprincipal.mainloop()
