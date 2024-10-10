import tkinter as tk
from tkinter import messagebox
from modules.pdf_processor import PDFProcessor

class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Aplicação de Automação")
        self.root.geometry("500x400")
        
        tk.Label(self.root, text="Bem-vindo à Aplicação de Automação!").pack(pady=20)
        
        # Botão para acessar o módulo de processamento de PDFs
        tk.Button(self.root, text="Processar PDFs", command=self.open_pdf_processor).pack(pady=10)
        
        # Futuramente, outros módulos podem ser adicionados aqui
        tk.Button(self.root, text="Outros Módulos", command=self.show_other_modules).pack(pady=10)
        
    def open_pdf_processor(self):
        processor_window = tk.Toplevel(self.root)
        PDFProcessor(processor_window)
    
    def show_other_modules(self):
        messagebox.showinfo("Em breve", "Mais funcionalidades serão adicionadas futuramente!")
