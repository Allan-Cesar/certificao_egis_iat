import pdfplumber
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from copy import deepcopy

def extract_ncs_from_pdf(pdf_path):
    # Abrir o PDF
    with pdfplumber.open(pdf_path) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text()

    # Separar as linhas e buscar pelas NCs numeradas
    lines = text.split("\n")
    ncs = []
    current_nc = []
    collecting = False

    # Palavras-chave que indicam o fim de uma NC e início de informações irrelevantes
    end_keywords = ["Relatório", "Cliente", "Folha", "Concessionária", "Num:", "Etapa", "OM", "Oportunidade de Melhorias", "FOR-713"]

    for line in lines:
        line = line.strip()

        # Início da captura das NCs numeradas ou requisitos não atendidos
        if line.startswith(tuple(f"{i}." for i in range(1, 100))) and "Requisito não atendido:" in line:
            if collecting:
                # Se já está coletando, adiciona a NC anterior antes de iniciar uma nova
                ncs.append("\n".join(current_nc))
            collecting = True  # Sinaliza que vamos começar a capturar a NC
            current_nc = [line]  # Começa uma nova NC
        
        elif collecting and current_nc and not any(keyword in line for keyword in end_keywords):
            # Continuar capturando as linhas da NC enquanto não encontrar uma palavra-chave que indica fim de NC
            current_nc.append(line)
        
        elif collecting and any(keyword in line for keyword in end_keywords):
            # Se encontrar uma palavra que indica fim da NC ou conteúdo irrelevante, parar a captura da NC atual
            if current_nc:
                ncs.append("\n".join(current_nc))
                current_nc = []
            collecting = False

    # Adiciona a última NC capturada, se houver
    if current_nc:
        ncs.append("\n".join(current_nc))

    return ncs

def copy_images(source_sheet, target_sheet):
    """Copia todas as imagens de uma planilha para outra."""
    if source_sheet._images:
        for img in source_sheet._images:
            # Criar uma cópia da imagem
            new_img = deepcopy(img)
            # Adicionar a imagem na nova aba
            target_sheet.add_image(new_img)

def save_ncs_to_template_excel(ncs, template_path, excel_path):
    # Carregar o arquivo de modelo existente
    wb = load_workbook(template_path)
    template_sheet = wb.active  # Supondo que o layout base está na primeira aba

    for i, nc in enumerate(ncs, start=1):
        # Criar uma nova aba duplicada para cada NC
        new_sheet = wb.copy_worksheet(template_sheet)
        new_sheet.title = f"NC{i}"  # Nome da aba com o número da NC

        # Copiar as imagens da aba de modelo
        copy_images(template_sheet, new_sheet)

        # Adicionar o texto da NC na célula A14
        new_sheet.cell(row=14, column=1, value=nc)

    # Remover a aba original se não for necessário (opcional)
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
    
    # Salvar o arquivo Excel com as NCs adicionadas
    wb.save(excel_path)

# Funções da interface gráfica
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def select_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        pdf_path_var.set(file_path)

def select_output_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        output_folder_var.set(folder_path)

def select_template():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        template_path_var.set(file_path)

def process_pdf():
    pdf_path = pdf_path_var.get()
    output_folder = output_folder_var.get()
    template_path = template_path_var.get()

    if not pdf_path or not output_folder or not template_path:
        messagebox.showwarning("Erro", "Por favor, selecione um arquivo PDF, um modelo de Excel e uma pasta de saída.")
        return
    
    try:
        ncs = extract_ncs_from_pdf(pdf_path)
        if ncs:
            output_excel_path = os.path.join(output_folder, "NCs_output.xlsx")
            save_ncs_to_template_excel(ncs, template_path, output_excel_path)
            messagebox.showinfo("Sucesso", f"Arquivo Excel criado com {len(ncs)} NCs, cada uma em uma aba.")
        else:
            messagebox.showwarning("Nenhuma NC", "Nenhuma NC foi encontrada no arquivo PDF.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

# Configuração da interface gráfica
app = tk.Tk()
app.title("Processador de NCs")
app.geometry("500x400")

pdf_path_var = tk.StringVar()
output_folder_var = tk.StringVar()
template_path_var = tk.StringVar()

# Elementos da interface
tk.Label(app, text="Selecione o PDF com as NCs:").pack(pady=10)
tk.Entry(app, textvariable=pdf_path_var, width=50, state="readonly").pack(pady=5)
tk.Button(app, text="Selecionar PDF", command=select_pdf).pack(pady=5)

tk.Label(app, text="Selecione a pasta de saída:").pack(pady=10)
tk.Entry(app, textvariable=output_folder_var, width=50, state="readonly").pack(pady=5)
tk.Button(app, text="Selecionar Pasta", command=select_output_folder).pack(pady=5)

tk.Label(app, text="Selecione o modelo de Excel:").pack(pady=10)
tk.Entry(app, textvariable=template_path_var, width=50, state="readonly").pack(pady=5)
tk.Button(app, text="Selecionar Modelo", command=select_template).pack(pady=5)

tk.Button(app, text="Iniciar Processamento", command=process_pdf, bg="green", fg="white").pack(pady=20)

# Iniciar a interface
app.mainloop()
