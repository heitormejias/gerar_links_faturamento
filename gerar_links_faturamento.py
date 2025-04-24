import fitz
import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, StringVar, scrolledtext
from PyPDF2 import PdfReader, PdfWriter


default_bgcolor = "#00BCD4"
default_font = "Helvetica"
default_font_size = 11
temp_file = "merged_temp_file.pdf"
ignored_files = [temp_file]


# Function to find pdf files in rootPath
def find_pdf_file(root_path, file_name):
    lfiles = []

    for file in os.listdir(root_path):
        if str(file).lower().endswith(".pdf") and file not in ignored_files:
            lfiles.append(os.path.join(root_path, file))
    return lfiles

# Function to extract text from a PDF
def extract_text_from_pdf(file_path):
    reader = PdfReader(file_path)
    text = []
    for page in reader.pages:
        text.append(page.extract_text())
    return text

# Function to identify marking inside parentheses
def find_marking_in_text(text):
    pattern = r'\((\bLink_\d{3}\b)\)'
    return re.findall(pattern, text)

# Function to merge PDF files
def merge_pdfs(rootPath, file_paths):
    writer = PdfWriter()
    for file_path in file_paths:
        reader = PdfReader(file_path)
        for page in reader.pages:
            writer.add_page(page)
    output_path = os.path.join(rootPath, temp_file)
    with open(output_path, "wb") as f_out:
        writer.write(f_out)
    return output_path

def create_pdf_with_links(rootPath, merged_pdf_path):
    # Reabrir o documento para processar apenas os textos entre parênteses no quadro de despesas
    doc = fitz.open(merged_pdf_path)

    # Extrair texto e coordenadas da primeira página (arquivo-1)
    page = doc[0]
    full_text = page.get_text()
    words = page.get_text("words")
    words.sort(key=lambda w: (w[1], w[0]))  # Ordenar por posição

    # Extrair todos os textos entre parênteses no quadro de despesas
    # Vamos considerar uma região específica da página como "quadro de despesas"
    # Supondo que o quadro de despesas esteja verticalmente entre y=180 e y=700
    despesa_words = [w for w in words if 180 <= w[1] <= 700 and re.fullmatch(r"\(\d+\)", w[4])]
    despesa_numeros = [re.search(r"\d+", w[4]).group() for w in despesa_words]

    # Mapear as páginas onde cada número é encontrado após a primeira
    num_to_page = {}
    for num in despesa_numeros:
        for i, p in enumerate(doc.pages(1, len(doc))):  # ignorar primeira
            if num in p.get_text():
                num_to_page[num] = i + 1
                break

    # Adicionar links diretamente sobre os parênteses do quadro de despesas
    linked = set()
    for w in despesa_words:
        x0, y0, x1, y1, text, *_ = w
        num = re.search(r"\d+", text).group()
        if num in num_to_page and num not in linked:
            page.insert_link({
                "from": fitz.Rect(x0, y0, x1, y1),
                "kind": fitz.LINK_GOTO,
                "page": num_to_page[num]
            })
            linked.add(num)

    # Salvar o PDF final com os links nos parênteses do quadro de despesas
    output_despesas_linkado = os.path.join(rootPath, "demostrativo-consolidado.pdf")
    doc.save(output_despesas_linkado)
    doc.close()

def delete_temp_file(merged_pdf_path):
    if os.path.exists(merged_pdf_path):
        os.remove(merged_pdf_path)


def cmd_demonstrative_file():
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo demonstrativo (.pdf)",
        filetypes=[("PDF", "*.pdf")])
    
    if file_path:
        demonstrative_entry.delete(0, tk.END)
        demonstrative_entry.insert(0, file_path)
        ignored_files.append(file_path.split('/')[-1])
        add_log(f"✅ Demonstrativo selecionado: {file_path}")

def cmd_attachments_folder():
    folder_path = filedialog.askdirectory(
        title="Selecione a pasta dos anexos (.pdf)")
    
    if folder_path:
        attachments_entry.delete(0, tk.END)
        attachments_entry.insert(0, folder_path)
        add_log(f"✅ Pasta de anexos selecionada: {folder_path}")

def add_log(message):
    log_text.config(state="normal")
    log_text.insert(tk.END, message + "\n")
    log_text.see(tk.END)
    log_text.config(state="disabled")

def start_consolidation():
    if demonstrative_entry.get() == "None":
        add_log("❌ Arquivo de demonstrativo não encontrado!")
        return None

    
            
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
    attachments_label = ttk.Label(window, text="Pasta de anexos:", style="Custom.TLabel")
    attachments_label.configure(background=default_bgcolor)
    attachments_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")

    attachments_entry = ttk.Entry(window, width=70)
    attachments_entry.grid(row=1, column=1, padx=10, pady=10)

    attachments_button = ttk.Button(window, text='Procurar...', command=cmd_attachments_folder)
    attachments_button.grid(row=1, column=2, padx=10, pady=10)

    # Button to start consolidation
    button = tk.Button(window, text="INICIAR CONSOLIDAÇÃO", command=start_consolidation, width=30, height=2)
    button.grid(row=2, columnspan=3, pady=50)

    # Logs 
    log_label = ttk.Label(window, text="Logs:", style="Custom.TLabel")
    log_label.configure(background=default_bgcolor)
    log_label.grid(row=3, column=0, padx=10, pady=2, sticky="w")

    log_text = scrolledtext.ScrolledText(window, width=98, height=6, foreground="gray", background="white", state="disabled")
    log_text.grid(row=4, columnspan=4, padx=10, pady=2, sticky="w")

    window.mainloop()
