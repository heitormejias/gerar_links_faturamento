import fitz  # PyMuPDF
import os
import tkinter as tk
from tkinter import filedialog, messagebox, StringVar
from PyPDF2 import PdfReader, PdfWriter
from dotenv import load_dotenv

def criar_links(pdf_path):
    try:
        print(f"\nAbrindo: {pdf_path}")
        doc = fitz.open(pdf_path)

        pagina_indice = doc[0]
        encontrados = []
        temas = os.getenv("PALAVRAS_CHAVES").split(",")

        global logs_str
        logs_str.set("")
        for tema in temas:
            print(f"🔍 Buscando: {tema}")
            logs_str.set(logs_str.get() + f"\n🔍 Buscando: {tema}")
            # pula uma pagina, comeca na 1.
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
                        print(f"✅ Tema '{tema}' encontrado na página {i + 1}")
                        logs_str.set(logs_str.get()+ f"\n✅ Tema '{tema}' encontrado na página {i + 1}")
                    break

        if encontrados:
            novo_arquivo = os.path.join(os.path.dirname(pdf_path), os.getenv("NOME_ARQUIVO_SAIDA", "ARQUIVO_SAIDA.pdf"))
            doc.save(novo_arquivo)
            doc.close()
            print(f"\n✅ Arquivo salvo com links: {novo_arquivo}")
            logs_str.set(logs_str.get() + f"\n\nArquivo salvo com sucesso:\n{novo_arquivo}")
            messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso:\n{novo_arquivo}")
        else:
            print("\n❌ Nenhum tema encontrado.")
            messagebox.showinfo("Aviso", "Nenhum tema encontrado no documento.")

    except Exception as e:
        print(f"❌ ERRO: {e}")
        messagebox.showerror("Erro", str(e))


# Function to find pdf files
def merge_files(file_path, folder_path):
    lfiles = []
    rootPath = os.path.dirname(file_path)
    ignored_files = os.getenv("ARQUIVO_TEMP").split(",")
    for file in os.listdir(folder_path):
        if str(file).lower().endswith(".pdf") and file not in ignored_files:
            lfiles.append(os.path.join(folder_path, file))

    writer = PdfWriter()
    reader = PdfReader(file_path)
    for page in reader.pages:
        writer.add_page(page)

    for file_pdf in lfiles:
        reader = PdfReader(file_pdf)
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
        files_pdfs_path = filedialog.askdirectory(title="Selecione o pasta com os anexos (.pdf)", initialdir=".")
        if files_pdfs_path:
            new_temp_file = merge_files(file_path, files_pdfs_path)
            print(f"\n✅ Novo arquivo temp: {new_temp_file}")
            criar_links(new_temp_file)
            delete_temp_file(new_temp_file)

load_dotenv()  # take environment variables

janela = tk.Tk()
janela.title("GERAR LINKS NOVO FATURAMENTO")
janela.geometry("400x400")
janela.configure(bg="#00BCD4")

botao = tk.Button(janela, text="Buscar PDF e Gerar Links", command=escolher_pdf, width=30, height=2)
botao.pack(pady=6)

logs_str = StringVar(value="Aqui sera apresentado os logs")
label = tk.Label(janela,textvariable=logs_str)
label.pack(pady=6)


janela.mainloop()