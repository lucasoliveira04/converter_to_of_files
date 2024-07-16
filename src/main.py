import subprocess
import tkinter as tk
from tkinter import filedialog
from docx2pdf import convert as docx_to_pdf_convert
from tqdm import tqdm
from tkinter import ttk
import time
from pdf2docx import Converter
import shutil
import os
import pathlib 

def converter_docx_para_pdf(caminho_docx):
    progresso_variavel.set(0)  
    progresso_barra["value"] = 0
    janela.update_idletasks()

    try:
        for _ in tqdm(range(100), desc="Convertendo DOCX para PDF", unit="%", dynamic_ncols=True):
            time.sleep(0.02)  
            progresso_variavel.set(progresso_variavel.get() + 1)
            progresso_barra["value"] = progresso_variavel.get()
            janela.update_idletasks()

        docx_to_pdf_convert(caminho_docx)
        pdf_path = f"{caminho_docx[:-5]}.pdf"
        print(f"Conversão concluída. O arquivo PDF foi salvo em: {pdf_path}")
        abrir_explorer(pathlib.Path.home() / "Downloads")  
    except Exception as e:
        print(f"Erro na conversão: {e}")

def converter_pdf_para_docx(caminho_pdf):
    progresso_variavel.set(0)  
    progresso_barra["value"] = 0
    janela.update_idletasks()

    try:
        for _ in tqdm(range(100), desc="Convertendo PDF para DOCX", unit="%", dynamic_ncols=True):
            time.sleep(0.02)  
            progresso_variavel.set(progresso_variavel.get() + 1)
            progresso_barra["value"] = progresso_variavel.get()
            janela.update_idletasks()

        cv = Converter(caminho_pdf)
        docx_path = f"{caminho_pdf[:-4]}.docx"
        cv.convert(docx_path, start=0, end=None)
        cv.close()
        print(f"Conversão concluída. O arquivo DOCX foi salvo em: {docx_path}")
        abrir_explorer(pathlib.Path.home() / "Downloads") 
    except Exception as e:
        print(f"Erro na conversão: {e}")

def selecionar_arquivo_docx():
    caminho_docx = filedialog.askopenfilename(filetypes=[("Arquivos DOCX", "*.docx")])
    if caminho_docx:
        converter_docx_para_pdf(caminho_docx)

def selecionar_arquivo_pdf():
    caminho_pdf = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
    if caminho_pdf:
        converter_pdf_para_docx(caminho_pdf)

def local_arquivos():
    caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos DOCX", "*")])
    if caminho_arquivo:
        abrir_explorer(os.path.dirname(caminho_arquivo))

def abrir_explorer(diretorio):
    if os.path.exists(diretorio):
        subprocess.Popen(f'explorer "{os.path.abspath(diretorio)}"')
    else:
        print(f"Diretório não encontrado: {diretorio}")

janela = tk.Tk()
janela.geometry("300x300")
janela.title("Conversor DOCX para PDF e PDF para DOCX")

botao_selecionar_docx = tk.Button(janela, text="Selecionar Arquivo DOCX", command=selecionar_arquivo_docx)
botao_selecionar_docx.pack(pady=20)

botao_selecionar_pdf = tk.Button(janela, text="Selecionar Arquivo PDF", command=selecionar_arquivo_pdf)
botao_selecionar_pdf.pack(pady=20)

progresso_variavel = tk.DoubleVar()
progresso_barra = ttk.Progressbar(janela, variable=progresso_variavel, maximum=100)
progresso_barra.pack(pady=10)

button_open_local = tk.Button(janela, text="Local do arquivo", command=local_arquivos)
button_open_local.pack(pady=5)

janela.mainloop()
