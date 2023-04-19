import tkinter as tk
import re
import os.path
from openpyxl import Workbook, load_workbook
from datetime import datetime

class FullscreenToolbar(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.attributes("-fullscreen", True)
        self.create_widgets()

    def create_widgets(self):
        self.toolbar = tk.Frame(self.master, bg='grey', height=50)
        self.toolbar.pack(side='top', fill='x')

        self.exit = tk.Button(self.toolbar, text='Exit', command=self.button1_action)
        self.button1.pack(side='left', padx=10)

        self.button2 = tk.Button(self.toolbar, text='Button 2', command=self.button2_action)
        self.button2.pack(side='left', padx=10)

    def button1_action(self):
        self.master.quit()

    def button2_action(self):
        print('Button 2 clicked')

# Fonction pour enregistrer les données du formulaire


def main():
    root = tk.Tk()
    app = FullscreenToolbar(root)
    
    root.mainloop()

if __name__ == '__main__':
    main()