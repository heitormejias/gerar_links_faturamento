from PyPDF2 import PdfReader, PdfWriter
import fitz
import re
import os

ignored_files = ['temp_file.pdf', 'demostrativo-consolidado.pdf']

# Function to find pdf files
def find_pdf_files(rootPath):
    lfiles = []

    for file in os.listdir(rootPath):
        if str(file).lower().endswith(".pdf") and file not in ignored_files:
            lfiles.append(os.path.join(rootPath, file))
    return lfiles

# Function to extract text from a PDF
def extract_text_from_pdf(file_path):
    reader = PdfReader(file_path)
    text = []
    for page in reader.pages:
        text.append(page.extract_text())
    return text

# Function to identify expression inside parentheses
def find_expression_in_text(text):
    pattern = r'\((\d+)\)'  # Regex pattern to find expression in parentheses
    return re.findall(pattern, text)

# Function to merge PDF files
def merge_pdfs(rootPath, file_paths):
    writer = PdfWriter()
    for file_path in file_paths:
        reader = PdfReader(file_path)
        for page in reader.pages:
            writer.add_page(page)
    output_path = os.path.join(rootPath, "temp_file.pdf")
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


if __name__ == "__main__":
    # rootPath = input('Insira o caminho dos arquivos pdf: ').replace("\'\'", "\'\\'")
    # paymentStatement = input('Insira o nome do arquivo demostrativo: ').replace("\'\'", "\'\\'")


    rootPath = "/home/heitor/PycharmProjects/Exemplo/static"
    paymentStatement = "demostrativo.pdf"

    file_paths = find_pdf_files(rootPath)

    if not file_paths:
        print("### ARQUIVOS NÃO ENCONTRADOS !!! ###")
        exit()

    if not paymentStatement:
        print("### ARQUIVO DEMOSTRATIVO NÃO INFORMADO !!! ###")
        exit()

    if os.path.join(rootPath, paymentStatement) not in file_paths:
        print("### ARQUIVO DEMOSTRATIVO NÃO ENCONTRADO !!! ###")
        exit()

    # Extract text from all files
    texts = [extract_text_from_pdf(file_path) for file_path in file_paths]

    # Find all expression in parentheses in each document
    expression_in_texts = [find_expression_in_text(text[0]) for text in texts]

    # Merge the PDFs in the specified order
    merged_pdf_path = merge_pdfs(rootPath, file_paths)

    # Create the PDF with links
    create_pdf_with_links(rootPath, merged_pdf_path)

    # Delete merged pdf
    delete_temp_file(merged_pdf_path)



