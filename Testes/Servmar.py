import customtkinter
import tkinter as tk
from tkinter import font
from PIL import Image
import threading
import os

def relatorio():
    thread = threading.Thread(target=rodar_relatorio)
    thread.start()
    status_label.config(text="")

def rodar_relatorio():
    status_label.config(text="Abrindo Gerar Relatório")
    import Mark4  # Importe o segundo programa como um módulo
    Mark4.main()
    status_label.config(text="")
    janelaprincipal.update

def analise():
    status_label.config(text="Programa em Desenvolvimento")

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

gerar = customtkinter.CTkButton(janelaprincipal, text="Gerar Relatório", command=relatorio).place(x=15, y=10)
teste = customtkinter.CTkButton(janelaprincipal, text="Gerar Análise", command=analise).place(x=190, y=10)

status_label = tk.Label(janelaprincipal, text="", bg="#ebebeb", font=("Arial", 15))
status_label.pack(pady=85)
fonte = font.nametofont("TkDefaultFont")
fonte.configure(underline=True)
fonte.configure(size=17)
status_label.config(font=fonte)

# Usa caminhos relativos para a imagem
imagem_path = os.path.join(script_dir, "servmarlogo.png")
img = customtkinter.CTkImage(light_image=Image.open(imagem_path), size=(345, 50))
imagem = customtkinter.CTkLabel(janelaprincipal, text="", image=img).place(x=5, y=120)

janelaprincipal.mainloop()
