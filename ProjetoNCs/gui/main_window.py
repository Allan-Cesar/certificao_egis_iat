import tkinter as tk
from tkinter import messagebox
import requests
import subprocess
import webbrowser
from modules.pdf_processor import PDFProcessor
from modules.nc_extractor import NCExtractorApp
from modules.respostas_extractor import ResponseExtractorApp

class MainWindow:
    
    version = "1.2.0"  # Versão atual do app
    
    def __init__(self, root):
        self.root = root
        self.root.title(f"Hand Helper v.{MainWindow.version}")
        self.root.geometry("500x400")
        
        tk.Label(self.root, text=f"Bem vindo à aplicação de Automação!").pack(pady=20)
        
        # Verifica se há atualizações disponíveis
        self.check_for_updates()

        # Botões para acessar os módulos
        tk.Button(self.root, text="Processar PDFs", command=self.open_pdf_processor).pack(pady=10)
        tk.Button(self.root, text="Extrair NCs Consolidadas", command=self.open_nc_extractor).pack(pady=10)
        tk.Button(self.root, text="Extrair Respostas", command=self.open_response_extractor).pack(pady=10)
        #tk.Button(self.root, text="Outros Módulos", command=self.show_other_modules).pack(pady=10)

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
        current_version = MainWindow.version
        version_file_url = "https://raw.githubusercontent.com/Allan-Cesar/Hand_Helper_EGIS/main/version.txt"
        headers = {
            "Authorization": "token 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"  # Seu token
        }

        try:
            response = requests.get(version_file_url, headers=headers)
            response.raise_for_status()  # Levanta erro se a resposta não for bem-sucedida
            latest_version = response.text.strip()  # Remove espaços extras

            print(f"Versão atual: {current_version}")  # Debug
            print(f"Versão mais recente: {latest_version}")  # Debug

            from packaging.version import Version
            if Version(current_version) < Version(latest_version):
                update_available = messagebox.askyesno(
                    "Atualização Disponível",
                    
                    f"Uma nova versão {latest_version} está disponível.\n\n"
                    f"Versão atual {current_version}.\n"
                    f"\nDeseja baixar a mais recente?"
                )
                if update_available:
                    # Busca pela última release usando a API do GitHub
                    latest_release_url = "https://api.github.com/repos/Allan-Cesar/Hand_Helper_EGIS/releases/latest"
                    print(f"Buscando release na URL: {latest_release_url}")  # Debug
                    release_response = requests.get(latest_release_url, headers=headers)
                    release_response.raise_for_status()  # Levanta erro se não for bem-sucedido

                    print(f"Resposta da API: {release_response.text}")  # Debug
                    latest_release = release_response.json()
                    download_url = latest_release['assets'][0]['browser_download_url']  # Pega o primeiro arquivo da release
                    print(f"URL de download: {download_url}")  # Debug
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
                messagebox.showinfo("Iniciando Download", "Tentando abrir o link de atualização no navegador.")
                webbrowser.open(download_url)
        except Exception as e:
            messagebox.showerror("Erro ao baixar atualização", str(e))
            
    def open_response_extractor(self):
        response_window = tk.Toplevel(self.root)
        response_window.transient(self.root)
        response_window.grab_set()
        response_window.focus_set()
        ResponseExtractorApp(response_window)


if __name__ == "__main__":
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()
