import tkinter as tk
from tkinter import messagebox
import requests
import subprocess
import webbrowser  # Importa para abrir o link no navegador, se necessário
from modules.pdf_processor import PDFProcessor
from modules.nc_extractor import NCExtractorApp

class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Aplicação de Automação")
        self.root.geometry("500x400")
        
        tk.Label(self.root, text="Bem-vindo à Aplicação de Automação!").pack(pady=20)

        self.check_for_updates()  # Verifica se há atualizações disponíveis

        # Botão para acessar o módulo de processamento de PDFs
        tk.Button(self.root, text="Processar PDFs", command=self.open_pdf_processor).pack(pady=10)

        # Botão para extrair NCs Consolidadas
        tk.Button(self.root, text="Extrair NCs Consolidadas", command=self.open_nc_extractor).pack(pady=10)

        # Botão para extrair Respostas
        tk.Button(self.root, text="Extrair Respostas", command=self.extract_responses).pack(pady=10)

        # Futuramente, outros módulos podem ser adicionados aqui
        tk.Button(self.root, text="Outros Módulos", command=self.show_other_modules).pack(pady=10)

    def open_pdf_processor(self):
        processor_window = tk.Toplevel(self.root)
        processor_window.transient(self.root)
        processor_window.grab_set()
        processor_window.focus_set()
        PDFProcessor(processor_window)
    
    def open_nc_extractor(self):
        nc_window = tk.Toplevel(self.root)
        nc_window.transient(self.root)
        nc_window.grab_set()
        nc_window.focus_set()
        NCExtractorApp(nc_window)

    def extract_responses(self):
        messagebox.showinfo("Extrair Respostas", "Função de Extração de Respostas em desenvolvimento.")
    
    def show_other_modules(self):
        messagebox.showinfo("Em breve", "Mais funcionalidades serão adicionadas futuramente!")

    def check_for_updates(self):
        current_version = "1.0.5"
        version_file_url = "https://raw.githubusercontent.com/Allan-Cesar/Hand_Helper_EGIS/main/version.txt"
        headers = {
            "Authorization": "token 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"  # Seu token
        }
        
        try:
            response = requests.get(version_file_url, headers=headers)
            response.raise_for_status()
            latest_version = response.text.strip()  # Remove espaços em branco extras

            from packaging.version import Version
            if Version(current_version) < Version(latest_version):
                update_available = messagebox.askyesno(
                    "Atualização Disponível",
                    f"Uma nova versão {latest_version} está disponível. Deseja atualizar?"
                )
                if update_available:
                    # Monta a URL de download com a versão correta
                    download_url = f"https://github.com/Allan-Cesar/Hand_Helper_EGIS/releases/download/{latest_version}/HandHelper.exe"
                    self.download_and_install_update(download_url)
        except requests.exceptions.HTTPError as err:
            messagebox.showerror("Erro ao verificar atualizações", f"Erro HTTP: {err}")
        except Exception as e:
            messagebox.showerror("Erro ao verificar atualizações", str(e))



    def download_and_install_update(self, download_url):
        try:
            response = requests.get(download_url)
            if response.status_code == 200:
                with open("Hand_Helper_Updated.exe", "wb") as file:
                    file.write(response.content)
                subprocess.Popen("Hand_Helper_Updated.exe")
                messagebox.showinfo("Atualização", "A atualização foi baixada e será instalada. Reinicie o aplicativo para aplicar as alterações.")
                self.root.destroy()
            else:
                # Abre a URL no navegador se o download falhar
                messagebox.showwarning("Download falhou", "Tentando abrir o link de atualização no navegador.")
                webbrowser.open(download_url)
        except Exception as e:
            messagebox.showerror("Erro ao baixar atualização", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()
