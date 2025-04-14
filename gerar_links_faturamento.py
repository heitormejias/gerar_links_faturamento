import fitz  # PyMuPDF
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter

temas = [
    "EXTRATO DA DECLARAÇÃO DE IMPORTAÇÃO",
    "BILL OF LADING",
    "AIR WAYBILL",
    "Fatura Comercial",
    "Invoice",
    "FACTURA",
    "PACKING LIST",
    "PACKING SLIP",
    "ALBARN",
    "GUIA ARRECADAÇÃO ESTADUAL",
    "GUIA PARA LIBERAÇÃO DE MERCADORIA ESTRANGEIRA",
    "DOCUMENTO DE ARRECADAÇÃO ESTADUAL",
    "DEPARTAMENTO DA MARINHA MERCANTE",
    "DAI - DOCUMENTO DE ARRECADAÇÃO DE IMPORTAÇÃO",
    "DACTE",
    "GUIA DE RECOLHIMENTO DE HONORARIOS - GRH",
    "DANFE",
    "COMPROVANTE IMPORTAÇÃO",
    "Extrato de Licença de Importação",
    "LAVAGEM",
    "DEMURRAGE",
    "RETIFICAÇÃO",

    "ICMS"
]

ignored_files = ['temp_file.pdf']

name_of_new_file = 'AAAAA.pdf'

def criar_links(pdf_path):
    try:
        print(f"\nAbrindo: {pdf_path}")
        doc = fitz.open(pdf_path)

        pagina_indice = doc[0]
        encontrados = []

        for tema in temas:
            print(f"🔍 Buscando: {tema}")
            for i in range(1, len(doc)):
                texto = doc[i].get_text()
                if tema.lower() in texto.lower():
                    encontrados.append((tema, i))
                    posicoes = pagina_indice.search_for(tema)
                    if posicoes:
                        pagina_indice.insert_link({
                            "from": posicoes[0],
                            "kind": fitz.LINK_GOTO,
                            "page": i
                        })
                    print(f"✅ Tema '{tema}' encontrado na página {i+1}")
                    break

        if encontrados:
            novo_arquivo = os.path.join(os.path.dirname(pdf_path), name_of_new_file)
            doc.save(novo_arquivo)
            doc.close()
            print(f"\n✅ Arquivo salvo com links: {novo_arquivo}")
            messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso:\n{novo_arquivo}")
        else:
            print("\n❌ Nenhum tema encontrado.")
            messagebox.showinfo("Aviso", "Nenhum tema encontrado no documento.")

    except Exception as e:
        print(f"❌ ERRO: {e}")
        messagebox.showerror("Erro", str(e))


# Function to find pdf files
def merge_files(file_path):
    lfiles = []
    files_path = os.path.join(os.path.dirname(file_path),"anexos")
    rootPath = os.path.dirname(file_path)

    for file in os.listdir(files_path):
        if str(file).lower().endswith(".pdf") and file not in ignored_files:
            lfiles.append(os.path.join(files_path, file))

    writer = PdfWriter()
    reader = PdfReader(file_path)
    for page in reader.pages:
        writer.add_page(page)

    for file_path in lfiles:
        reader = PdfReader(file_path)
        for page in reader.pages:
            writer.add_page(page)

    output_path = os.path.join(rootPath, "temp_file.pdf")
    delete_temp_file(output_path) # delete if exists
    with open(output_path, "wb") as f_out:
        writer.write(f_out)
    return output_path

def delete_temp_file(merged_pdf_path):
    if os.path.exists(merged_pdf_path):
        os.remove(merged_pdf_path)

def escolher_pdf():
    file_path = filedialog.askopenfilename(title="Selecione o arquivo FATURAMENTO.pdf", filetypes=[("PDF", "*.pdf")])
    if file_path:
        new_temp_file = merge_files(file_path)
        print(f"\n✅ Novo arquivo temp: {new_temp_file}")
        criar_links(new_temp_file)
        delete_temp_file(new_temp_file)

janela = tk.Tk()
janela.title("GERAR LINKS NOVO FATURAMENTO")
janela.geometry("400x400")
janela.configure(bg="#00BCD4")

botao = tk.Button(janela, text="Buscar PDF e Gerar Links", command=escolher_pdf, width=30, height=2)
botao.pack(pady=150)

janela.mainloop()