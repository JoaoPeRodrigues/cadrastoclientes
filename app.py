import customtkinter as ctk
from tkinter import *
from tkinter import messagebox
import openpyxl, xlrd
import pathlib
from openpyxl import Workbook

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.layot()
        self.aparencia()
        self.todo_sistema()

    def layot(self):
        self.title("Cadrasto de clientes")
        self.geometry("700x500")

    def aparencia(self):
        self.lb_apm = ctk.CTkLabel(self, text="Tema", bg_color='transparent', text_color=['#000', "#fff"]).place(x=50,y=430)
        self.lb_opt = ctk.CTkOptionMenu(self, values=["Dark", "light", "System"], command=self.change_apm).place(x=50,y=460)



    def todo_sistema(self):

        frame = ctk.CTkFrame(self, width=700, height=50, corner_radius=0, bg_color="teal", fg_color="teal")
        frame.place(x=0, y=10)
        title = ctk.CTkLabel(frame, text="Sistema de gestão de clientes", font=("Century Gothic Bold", 24), fg_color="transparent", text_color="#fff").place(x=190,y=10)
        span = ctk.CTkLabel(self, text="Por Favor , Preencha todos os dados do formulário", font=("Century Gothic Bold", 16), text_color=["#000","#fff"]).place(x=50,y=70)

        ficheiro = pathlib.Path('Clientes.xlsx')
        if ficheiro.exists():
            pass
        else:
            ficheiro = Workbook()
            folha = ficheiro.active
            folha['A1']="Nome completo"
            folha['B1']="Contato"
            folha['C1']="Idade"
            folha['D1']="Genero"
            folha['E1']="Endereço"
            folha['F1']="Observações"

            ficheiro.save("Clientes.xlsx")
        
        def salvar():

            nome = nome_valor.get()
            contato = contato_valor.get()
            idade = idade_valor.get()
            endereco = endereco_valor.get()
            genero = genero_combox.get()
            obs = obs_entry.get()

            if(nome=="" or contato=="" or idade=="" or endereco=="" or obs==""):
                messagebox.showinfo('Sistema', 'ERRO!\nPreencha todos os dados.')
            else:
                ficheiro = openpyxl.load_workbook('Clientes.xlsx')
                folha = ficheiro.active
                folha.cell(column=1, row=folha.max_row+1, value=nome)
                folha.cell(column=2, row=folha.max_row, value=contato)
                folha.cell(column=3, row=folha.max_row, value=idade)
                folha.cell(column=4, row=folha.max_row, value=genero)
                folha.cell(column=5, row=folha.max_row, value=endereco)
                folha.cell(column=6, row=folha.max_row, value=obs)

                ficheiro.save(r'Clientes.xlsx')
                messagebox.showinfo("Sistema", "Dados salvos com sucesso!")

                limpar()

        def limpar():
            nome_valor.set("")
            contato_valor.set("")
            idade_valor.set("")
            endereco_valor.set("")
        

        nome_valor = StringVar()
        contato_valor = StringVar()
        idade_valor = StringVar()
        endereco_valor = StringVar()


        nome_entry = ctk.CTkEntry(self, width=300, textvariable=nome_valor, font=("Century Gothic Bold", 16), fg_color="transparent", placeholder_text="Nome")
        contato_entry = ctk.CTkEntry(self, width=200,textvariable=contato_valor, font=("Century Gothic Bold", 16), fg_color="transparent", placeholder_text="Contato")
        idade_entry = ctk.CTkEntry(self, width=150,textvariable=idade_valor, font=("Century Gothic Bold", 16), fg_color="transparent", placeholder_text="Idade")
        endereco_entry = ctk.CTkEntry(self, width=200,textvariable=endereco_valor, font=("Century Gothic Bold", 16), fg_color="transparent", placeholder_text="Endereço")
        obs_entry = ctk.CTkEntry(self, width=450, height=150,  font=('Century Gothic Bold', 16), border_color="#aaa", border_width=2, fg_color="transparent")

        genero_combox = ctk.CTkComboBox(self, values=["feminino","Masculino"], font=('Century Gothic Bold', 14), width=150)
        genero_combox.set("Masculino")


        lb_nome = ctk.CTkLabel(self, text="Nome: ", font=("Century Gothic Bold", 16), text_color=["#000","#fff"])
        lb_contato = ctk.CTkLabel(self, text="Contato", font=("Century Gothic Bold", 16), text_color=["#000","#fff"])
        lb_idade = ctk.CTkLabel(self, text="Idade", font=("Century Gothic Bold", 16), text_color=["#000","#fff"])
        lb_genero = ctk.CTkLabel(self, text="Genero", font=("Century Gothic Bold", 16), text_color=["#000","#fff"])
        lb_endereco = ctk.CTkLabel(self, text="Endereco", font=("Century Gothic Bold", 16), text_color=["#000","#fff"])
        lb_obs = ctk.CTkLabel(self, text="Observações", font=("Century Gothic Bold", 16), text_color=["#000","#fff"])

        btn_submit = ctk.CTkButton(self, text="Salvar Dados".upper(), command=salvar, fg_color="#151", hover_color="#131").place(x=300,y=420)
        btn_clear = ctk.CTkButton(self, text="Limpar Dados".upper(), command=limpar, fg_color="#555", hover_color="#333").place(x=470, y=420)
        
        lb_nome.place(x=50,y=120)
        nome_entry.place(x=50, y=150) 
        
        lb_contato.place(x=450, y=120)
        contato_entry.place(x=450,y=150)

        lb_idade.place(x=300,y=190)
        idade_entry.place(x=300,y=220)

        lb_genero.place(x=500,y=190)
        genero_combox.place(x=500,y=220)

        lb_endereco.place(x=50,y=190)
        endereco_entry.place(x=50,y=220)

        lb_obs.place(x=50,y=260)
        obs_entry.place(x=160,y=260)



    def change_apm(self, nova_aparencia):
        ctk.set_appearance_mode(nova_aparencia)
    

if __name__=="__main__":
    app = App()
    app.mainloop()