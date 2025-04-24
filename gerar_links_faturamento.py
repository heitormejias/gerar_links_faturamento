import fitz
import os
import tkinter as tk
from PyPDF2 import PdfReader, PdfWriter
from tkinter import ttk, filedialog, messagebox, StringVar, scrolledtext


default_bgcolor = "#00BCD4"
default_font = "Helvetica"
default_font_size = 11


def criar_links(pdf_path):
    try:
        print(f"\nAbrindo: {pdf_path}")
        doc = fitz.open(pdf_path)

        pagina_indice = doc[0]
        encontrados = []

        global logs_str,themes,filename_out
        logs_str.set("")
        for tema in themes.get():
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
            novo_arquivo = os.path.join(os.path.dirname(pdf_path), filename_out.get())
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
    #if os.path.exists(merged_pdf_path):
    #    os.remove(merged_pdf_path)
    return None

# def escolher_pdf():
#     file_path = filedialog.askopenfilename(title="Selecione o arquivo FATURAMENTO.pdf", filetypes=[("PDF", "*.pdf")])
#     if file_path:
#         files_pdfs_path = filedialog.askdirectory(title="Selecione o pasta com os anexos (.pdf)", initialdir=".")
#         if files_pdfs_path:
#             new_temp_file = merge_files(file_path, files_pdfs_path)
#             print(f"\n✅ Novo arquivo temp: {new_temp_file}")
#             criar_links(new_temp_file)
#             delete_temp_file(new_temp_file)

def cmd_demonstrative_file():
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo demonstrativo (.pdf)",
        filetypes=[("PDF", "*.pdf")])
    
    if file_path:
        demonstrative_entry.delete(0, tk.END)
        demonstrative_entry.insert(0, file_path)
        add_log(f"✅ Arquivo selecionado: {file_path}")

def cmd_consolidate_folder():
    folder_path = filedialog.askdirectory(
        title="Selecione a pasta com os arquivos para consolidação (.pdf)")
    
    if folder_path:
        consolidate_entry.delete(0, tk.END)
        consolidate_entry.insert(0, folder_path)
        add_log(f"✅ Pasta selecionada: {folder_path}")

def add_log(message):
    log_text.config(state="normal")
    log_text.insert(tk.END, message + "\n")
    log_text.see(tk.END)
    log_text.config(state="disabled")

            
if __name__ == "__main__":
    window = tk.Tk()
    window.title("DEMONSTRATIVO DE DESPESAS - CONSOLIDAÇÃO")
    window.geometry("850x400")
    window.config(bg=default_bgcolor)

    # Label Font Style
    default_label_style = ttk.Style()
    default_label_style.configure("Custom.TLabel", font=(default_font, default_font_size))

    # Demonstrative file
    demonstrative_label = ttk.Label(window, text="Arquivo demonstrativo:", style="Custom.TLabel")
    demonstrative_label.configure(background=default_bgcolor)
    demonstrative_label.grid(row=0, column=0, padx=10, pady=20,sticky="w")

    demonstrative_entry = ttk.Entry(window, width=70)
    demonstrative_entry.grid(row=0, column=1, padx=10, pady=10)

    demonstrative_button = ttk.Button(window, text='Procurar...', command=cmd_demonstrative_file)
    demonstrative_button.grid(row=0, column=2, padx=10, pady=10)

    # Others files to consolidate
    consolidate_label = ttk.Label(window, text="Pasta com arquivos para consolidação:", style="Custom.TLabel")
    consolidate_label.configure(background=default_bgcolor)
    consolidate_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")

    consolidate_entry = ttk.Entry(window, width=70)
    consolidate_entry.grid(row=1, column=1, padx=10, pady=10)

    consolidate_button = ttk.Button(window, text='Procurar...', command=cmd_consolidate_folder)
    consolidate_button.grid(row=1, column=2, padx=10, pady=10)

    # Button to start consolidation
    button = tk.Button(window, text="INICIAR CONSOLIDAÇÃO", width=30, height=2)
    button.grid(row=2, columnspan=3, pady=50)

    # Logs 
    log_label = ttk.Label(window, text="Logs:", style="Custom.TLabel")
    log_label.configure(background=default_bgcolor)
    log_label.grid(row=3, column=0, padx=10, pady=2, sticky="w")

    log_text = scrolledtext.ScrolledText(window, width=98, height=6, foreground="gray", background="white", state="disabled")
    log_text.grid(row=4, columnspan=4, padx=10, pady=2, sticky="w")

    window.mainloop()





    # entrada = ttk.Entry(window)
    # entrada.pack(pady=5)

    # botao = ttk.Button(window, text="Clique aqui")
    # botao.pack(pady=10)



    # file_temp: str='temp_file.pdf'

    # logs_str = StringVar(value="Aqui sera apresentado os logs")

    # # Create and place the filename_out label and entry
    # filename_out = StringVar(value=f"DEMONSTRATIVO-DE-DESPESAS-CONSOLIDADO.pdf")
    # filename_out_label = tk.Label(window, text="Arquivo de saida:")
    # filename_out_label.configure(bg="#F6F6F6")
    # filename_out_label.pack()
    # filename_out_entry = tk.Entry(window, textvariable=filename_out, width=40)
    # filename_out_entry.pack(pady=6)

    # # Create and place the Button
    # botao = tk.Button(window, text="Buscar PDF e Gerar Links", command=escolher_pdf, width=30, height=2)
    # botao.pack(pady=6)

    # # Create and place the Logs
    # label = tk.Label(window,textvariable=logs_str, width=40, height=5)
    # label.configure(bg="#ffffff")
    # label.pack(pady=12)

    # # main loop
    # window.mainloop()