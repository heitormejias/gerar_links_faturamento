import fitz
import os
import tkinter as tk
from tkinter import filedialog, messagebox, StringVar
from PyPDF2 import PdfReader, PdfWriter

def criar_links(pdf_path):
    try:
        print(f"\nAbrindo: {pdf_path}")
        doc = fitz.open(pdf_path)

        pagina_indice = doc[0]
        encontrados = []

        global logs_str,themes,filename_out
        logs_str.set("")
        for tema in themes:
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
            novo_arquivo = os.path.join(os.path.dirname(pdf_path), filename_out)
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
    global file_temp
    ignored_files = file_temp.split(",")
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

    output_path = os.path.join(rootPath, file_temp)
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


janela = tk.Tk()
janela.title("GERAR LINKS NOVO FATURAMENTO")
janela.geometry("400x400")
janela.configure(bg="#00BCD4")

# variables
themes = StringVar(value="DAI,DACTE,GRH,DANFE,ICMS")
filename_out = StringVar(value="DEMONSTRATIVO DE DESPESAS.pdf")
logs_str = StringVar(value="Aqui sera apresentado os logs")
file_temp: str='temp_file.pdf'

# Create and place the filename_out label and entry
filename_out_label = tk.Label(janela, text="Arquivo de saida:")
filename_out_label.configure(bg="#00BCD4")
filename_out_label.pack()
filename_out_entry = tk.Entry(janela, textvariable=filename_out, width=40)
filename_out_entry.pack(pady=6)

# Create and place the PALAVRAS_CHAVES label and entry
temas_label = tk.Label(janela, text="Temas:")
temas_label.configure(bg="#00BCD4")
temas_label.pack()
temas_entry = tk.Entry(janela, textvariable=themes, width=40)
temas_entry.pack(pady=6)

# Create and place the Button
botao = tk.Button(janela, text="Buscar PDF e Gerar Links", command=escolher_pdf, width=30, height=2)
botao.pack(pady=6)

# Create and place the Logs
label = tk.Label(janela,textvariable=logs_str, width=40, height=5)
label.configure(bg="#ffffff")
label.pack(pady=12)

# main loop
janela.mainloop()