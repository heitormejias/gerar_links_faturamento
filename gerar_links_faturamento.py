import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import threading
import os
from PyPDF2 import PdfReader, PdfWriter
import re
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from io import BytesIO
from datetime import datetime
import fitz



default_bgcolor = "#00BCD4"
default_font = "Helvetica"
default_font_size = 11
merged_temp_file = "MERGED-TEMP-FILE.pdf"
final_filename = f"DEMONSTRATIVO-DESPESAS-CONSOLIDADDO-{datetime.now().strftime("%d%m%Y")}.pdf"



# Function to add logs in log text area
def add_log(message):
    log_text.config(state="normal")
    log_text.insert(tk.END, message + "\n")
    log_text.insert(tk.END, "====================================================================================================" + "\n")
    log_text.see(tk.END)
    log_text.config(state="disabled")

# Function to identify marking inside parentheses
def find_markings_in_text(text):
    pattern = r"Link_\d{3}"
    return re.findall(pattern, text, re.IGNORECASE)

# Function to extract text from a PDF
def extract_text_from_pdf(file_path):
    reader = PdfReader(file_path)
    text = []
    for page in reader.pages:
        text.append(page.extract_text())
    return text

# Function to get marking from demonstrative pdf file
def get_markings_from_demonstrative():
    file_path = demonstrative_entry.get()
    text = ' '.join(extract_text_from_pdf(file_path))
    markings_in_text = find_markings_in_text(text)
    return markings_in_text

# Function to get attachments pdf files from markings
def find_attachments_from_marks():
    attachments_folder = attachments_entry.get()
    attachments_files = os.listdir(attachments_folder)
    markigns = get_markings_from_demonstrative()
    dmarkings = {marking: None for marking in markigns}

    if len(markigns) == 0:
        add_log("❌ Não foram encontradas marcações no arquivo de demonstrativo!")
        return None
    
    for marking in markigns:
        for file in attachments_files:
            if marking.lower() in file.lower():
                dmarkings[marking] = file
                break
    return dmarkings

# Function to delete temp file
def delete_temp_file(merged_pdf_path):
    if os.path.exists(merged_pdf_path):
        os.remove(merged_pdf_path)

# Function to create text with the file name in the first page of the merged file
def set_filename_in_pdf(text, page_width, page_height):
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=(page_width, page_height))
    can.setFillColorRGB(0,0,0)
    can.setFont("Helvetica", 5)
    can.drawString(2, 2, text.upper())
    can.save()
    packet.seek(0)
    overlay_pdf = PdfReader(packet)
    return overlay_pdf.pages[0]

# Function to merge PDF files
def merge_pdfs(rootPath, demonstrative_file, attachments_files):
    try:
        add_log("🕒 Iniciando o merge dos anexos!")
        writer = PdfWriter()

        # First, merge demonstrative file
        reader = PdfReader(demonstrative_file)
        for page in reader.pages:
            writer.add_page(page)

        # Second, merge each attachment file
        for file in attachments_files:
            if file:
                file_path = os.path.join(rootPath, file).replace("\\", "/")
                reader = PdfReader(file_path)
                for idx, page in enumerate(reader.pages):
                    if idx == 0:
                        # Primeira página -> adicionar texto no topo
                        overlay = set_filename_in_pdf(f"Arquivo: {file}",
                                                    float(page.mediabox.width),
                                                    float(page.mediabox.height))
                        page.merge_page(overlay)
                    writer.add_page(page)

        output_path = os.path.join(rootPath, merged_temp_file)

        with open(output_path, "wb") as f_out:
            writer.write(f_out)

        add_log(f"✅ Merge dos anexos efetuado com sucesso!")

    except Exception as e:
        add_log(f"❌ Erro ao efetuar o merge dos anexos: {e}")

# Function to search markings in the first page of merged temp file and create link
# to the first occurrence in others pages
def create_pdf_with_links(root_path, merged_temp_file):
    try:
        add_log(f"🕒 Iniciando a criação dos links no demonstrativo!")
        merged_temp_file_path = os.path.join(root_path, merged_temp_file).replace("\\", "/")
        doc = fitz.open(merged_temp_file_path)

        destination_links = {}

        # Search markings in pages, page 1 onwards
        for page_number in range(1, len(doc)):
            page = doc[page_number]
            page_text = page.get_text()

            found_markings = find_markings_in_text(page_text)
            for marking in found_markings:
                if marking not in destination_links:
                    destination_links[marking] = page_number
            

            # create links to return top
            text_back = 'IR PARA O TOPO'

            page_to_link = doc[page_number]
            point = fitz.Point(page_to_link.rect.width - 100, 10)
            shape = page_to_link.new_shape()
            bbox = fitz.Rect(point.x, point.y, point.x + 200, point.y + 30)  # largura e altura do retângulo
            shape.draw_rect(bbox)
            shape.finish(fill=(0.9, 0.9, 0.5))  # Cor de fundo amarelado (RGB de 0 a 1)

            # create a Shape to draw on

            shape.insert_text(point, text_back, fontsize=10, color=(1,0,0)) # red
            shape.commit()
            area_back = page_to_link.search_for(text_back)
            page_to_link.insert_link({
                "from": area_back[0],
                "kind": fitz.LINK_GOTO,
                "page": 0
            })

        # Search markings in page 0
        page_0 = doc[0]
        blocks = page_0.get_text("blocks")

        for block in blocks:
            x0, y0, x1, y1, text_block, *_ = block
            found_markings = find_markings_in_text(text_block)
            for marking in found_markings:
                if marking.upper() in destination_links:
                    destination_page = destination_links[marking.upper()]
                    pontos = page_0.search_for(marking.upper())
                    for ponto in pontos:
                        page_0.insert_link({
                            "kind": fitz.LINK_GOTO,
                            "from": ponto,
                            "page": destination_page,
                            "zoom": 0
                        })

        add_log(f"✅ Links criados no demonstrativo com sucesso!")
        add_log(f"🕒 Salvando o arquivo consolidado!")
        output_despesas_linkado = os.path.join(root_path, final_filename)
        doc.save(output_despesas_linkado)
        add_log(f"✅ Arquivo {output_despesas_linkado} salvo com sucesso!")
        doc.close()

    except Exception as e:
        add_log(f"❌ Ocorreu um erro ao criar os links no demonstrativo: {e}, por favor tente novamente!")

# Function to select demonstrative file
def cmd_demonstrative_file():
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo demonstrativo (.pdf)",
        filetypes=[("PDF", "*.pdf")])
    if file_path:
        demonstrative_entry.delete(0, tk.END)
        demonstrative_entry.insert(0, file_path)
        add_log(f"✅ Demonstrativo selecionado: {file_path}")

# Function to select attachments folder
def cmd_attachments_folder():
    folder_path = filedialog.askdirectory(
        title="Selecione a pasta dos anexos (.pdf)")
    if folder_path:
        attachments_entry.delete(0, tk.END)
        attachments_entry.insert(0, folder_path)
        add_log(f"✅ Pasta de anexos selecionada: {folder_path}")

# Function to start consolidation
def cmd_consolidation_button():
    consolidation_button.config(state=tk.DISABLED)
    def thread():
        demonstrative_file = demonstrative_entry.get()
        attachments_folder = attachments_entry.get()

        if not demonstrative_file or not attachments_folder:
            add_log(f"❌ Arquivo de demonstrativo ou pasta com anexos não encontrado!")
            return None

        dmarkings = find_attachments_from_marks()

        for marking in dmarkings.items():
            if marking[1]:
                add_log(f"✅ Arquivo encontrado: {marking[0].upper()} - {marking[1].upper()}")
            else:
                add_log(f"❌ Arquivo não encontrado para a marcação: {marking[0].upper()}")

        merge_pdfs(attachments_folder, demonstrative_file, dmarkings.values())
        consolidation_button.config(state=tk.NORMAL)

        create_pdf_with_links(attachments_folder, merged_temp_file)

    threading.Thread(target=thread).start()



if __name__ == "__main__":
    window = tk.Tk()
    window.title("DEMONSTRATIVO DE DESPESAS - CONSOLIDAÇÃO")
    window.geometry("850x600")
    window.config(bg=default_bgcolor)

    # Label Font Style
    default_label_style = ttk.Style()
    default_label_style.configure("Custom.TLabel", font=(default_font, default_font_size))

    # Demonstrative file
    demonstrative_label = ttk.Label(window, text="Arquivo de demonstrativo:", style="Custom.TLabel")
    demonstrative_label.configure(background=default_bgcolor)
    demonstrative_label.grid(row=0, column=0, padx=10, pady=20,sticky="w")

    demonstrative_entry = ttk.Entry(window, width=80)
    demonstrative_entry.grid(row=0, column=1, padx=10, pady=10)

    demonstrative_button = ttk.Button(window, text='Procurar...', command=cmd_demonstrative_file)
    demonstrative_button.grid(row=0, column=2, padx=10, pady=10)

    # Others files to consolidate
    attachments_label = ttk.Label(window, text="Pasta com anexos:", style="Custom.TLabel")
    attachments_label.configure(background=default_bgcolor)
    attachments_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")

    attachments_entry = ttk.Entry(window, width=80)
    attachments_entry.grid(row=1, column=1, padx=10, pady=10)

    attachments_button = ttk.Button(window, text='Procurar...', command=cmd_attachments_folder)
    attachments_button.grid(row=1, column=2, padx=10, pady=10)

    # Button to start consolidation
    consolidation_button = tk.Button(window, text="INICIAR CONSOLIDAÇÃO", command=cmd_consolidation_button, width=30, height=2)
    consolidation_button.grid(row=2, columnspan=3, pady=50)

    # Logs 
    log_label = ttk.Label(window, text="Logs:", style="Custom.TLabel")
    log_label.configure(background=default_bgcolor)
    log_label.grid(row=3, column=0, padx=10, pady=2, sticky="w")

    log_text = scrolledtext.ScrolledText(window, width=100, height=18, foreground="gray", background="white", state="disabled")
    log_text.grid(row=4, columnspan=6, padx=10, pady=2, sticky="w")

    window.mainloop()
