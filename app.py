import customtkinter as ctk
from tkinter import *
from tkinter import messagebox
import openpyxl
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
        frame = ctk.CTkFrame(self, width=700, height=50, corner_radius=0, bg_color="teal", fg_color="teal").place(x=0,y=10)

    def change_apm(self, nova_aparencia):
        ctk.set_appearance_mode(nova_aparencia)

if __name__=="__main__":
    app = App()
    app.mainloop()