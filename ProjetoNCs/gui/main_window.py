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
        # Abre uma nova janela para processar PDFs
        processor_window = tk.Toplevel(self.root)
        processor_window.transient(self.root)  # Define a janela principal como "pai"
        processor_window.grab_set()  # Bloqueia interação com a janela principal até que a nova seja fechada
        processor_window.focus_set()  # Foca na nova janela
        PDFProcessor(processor_window)
    
    def show_other_modules(self):
        messagebox.showinfo("Em breve", "Mais funcionalidades serão adicionadas futuramente!")

if __name__ == "__main__":
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()
