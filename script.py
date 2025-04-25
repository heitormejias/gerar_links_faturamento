#import fitz
import pymupdf

import os
import re
from PyPDF2 import PdfReader, PdfWriter
from pymupdf.utils import get_textpage_ocr

temp_file = "merged_temp_file.pdf"
final_new_file = "final_pdf_file.pdf"

class Dicionario(object):
    def __init__(self,code:str,filename:str="",page:int=0):
        self.code = code
        self.filename = filename
        self.page = page

    def set_page(self,page:int=0):
        self.page = page

    def set_filename(self,filename:str=""):
        self.filename = filename


# Meu dicionario
my_dict = [
    #---DESPESAS
    Dicionario('(Link_100)'), # ICMS(Link_100)
    Dicionario('(Link_101)'), # MARINHA MERCANTE(AFRMM)(Link_101)
    Dicionario('(Link_102)'), # FRETE INTL(Link_102)
    Dicionario('(Link_103)'), # TAXA DE LIBERAÇÃO AWB(Link_103)
    Dicionario('(Link_104)'), # DESPESAS FRETE INTL(Link_104)
    Dicionario('(Link_105)'), # DESPESAS FRETE INTL - ORIGEM(Link_105)
    Dicionario('(Link_106)'), # DESPESAS FRETE INTL - DESTINO(Link_106)
    Dicionario('(Link_107)'), # ARMAZENAGEM - SANTOS(Link_107)
    Dicionario('(Link_108)'), # ARMAZENAGEM - GRU / VCP(Link_108)
    Dicionario('(Link_109)'), # ARMAZENAGEM - EADI(Link_109)
    Dicionario('(Link_110)'), # TRANSP.ROD.LIBERAÇAO(Link_110)
    Dicionario('(Link_111)'), # TRANSP.ROD.REMOÇÃO(Link_111)
    Dicionario('(Link_112)'), # SDA(Link_112)
    Dicionario('(Link_113)'), # DEMURRAGE(Link_113)
    Dicionario('(Link_114)'), # FUMIGAÇÃO(Link_114)
    Dicionario('(Link_115)'), # EMISSÃO DE LI(Link_115)
    Dicionario('(Link_116)'), # TAXA DE INSPEÇÃO MADEIRA(Link_116)
    Dicionario('(Link_117)'), # COMPLEMENTO II(Link_117)
    Dicionario('(Link_118)'), # COMPLEMENTO IPI(Link_118)
    Dicionario('(Link_119)'), # COMPLEMENTO PIS(Link_119)
    Dicionario('(Link_120)'), # COMPLEMENTO DE COFINS(Link_120)
    Dicionario('(Link_121)'), # MULTA LI(Link_121)
    Dicionario('(Link_122)'), # SERVICO * DESEMBARAÇO DI(Link_122)
    Dicionario('(Link_123)'), # TAR.DOC / TED - FRETE INTL(Link_123)
    Dicionario('(Link_124)'), # TAR.DOC / TED - ARMAZENAGEM(Link_124)
    Dicionario('(Link_125)'), # LAVAGEM QUIMICA(Link_125)
    Dicionario('(Link_126)'), # LAVAGEM SIMPLES(Link_126)
    Dicionario('(LINK_127)'), # TAR.DOC / TED - LAVAGEM CNTR(LINK_127)

    #---DOCUMENTOS ANEXOS
    Dicionario('(Link_001)'), # DI_EXTRATO DECLARAÇÃO DE IMP.(Link_001)
    Dicionario('(Link_002)'), # BL_BILL OF LADING(Link_002)
    Dicionario('(Link_003)'), # AWB_AIR WAYBILL(Link_003)
    Dicionario('(Link_004)'), # FATURA COMERCIAL / INVOICE(Link_004)
    Dicionario('(Link_005)'), # PACKING LIST / PACKING SLIP(Link_005)
    Dicionario('(Link_006)'), # GARE(Link_006)
    Dicionario('(Link_007)'), # GLME(Link_007)
    Dicionario('(Link_008)'), # DARE - ICMS(Link_008)
    Dicionario('(Link_009)'), # MARINHA MERCANTE(Link_009)
    Dicionario('(Link_010)'), # DAI - ARMAZENAGEM(Link_010)
    Dicionario('(Link_011)'), # DACTE - (Link_011)
    Dicionario('(Link_012)'), # GUIA DE RECOLHIMENTO SDA - GRH(Link_012)
    Dicionario('(Link_013)'), # DANFE(Link_013)
    Dicionario('(Link_014)'), # CI_COMPROVANTE IMPORTAÇÃO(Link_014)
    Dicionario('(Link_015)'), # LI_LICENÇA DE IMPORTÇÃO(Link_015)
    Dicionario('(Link_016)'), # RETIFICAÇÃO DI(Link_016)
]


# Function to find pdf files in folderPath that's exists in dictionary
def find_pdfs_on_folder(folderPath: str):
    new_list=[]
    count_page = 1

    for file in os.listdir(folderPath):
        #[any(char in word for char in special_characters) for word in sentence]
        if file.endswith(".pdf"): #  and any(s.code.lower() in filenamestr for s in my_dict):
            for dic in my_dict:
                if dic.code.lower() in file.lower():
                    new_list.append(Dicionario(dic.code, file, count_page))
                    count_page += 1

    return list[Dicionario](filter(lambda x: x.filename != '', new_list))


# Function to merge PDF files
def merge_pdfs(file_pdf: str, rootPath :str, my_dict_list: list[Dicionario]):
    root_file_path = os.path.dirname(file_pdf)
    output_path = os.path.join(root_file_path, temp_file)
    #clean tmp file before
    delete_temp_file(output_path)

    writer = PdfWriter()
    #write file pdf in first place
    reader = PdfReader(file_pdf)
    for page in reader.pages:
        writer.add_page(page)

    #than, merge all files founded in dict to this unique file
    for dict_item in my_dict_list:
        reader = PdfReader(os.path.join(rootPath, dict_item.filename))
        for page in reader.pages:
            writer.add_page(page)

    with open(output_path, "wb") as f_out:
        writer.write(f_out)
    return output_path


# delete file temp
def delete_temp_file(merged_pdf_path):
    if os.path.exists(merged_pdf_path):
        os.remove(merged_pdf_path)


def create_links(file_pdf: str, my_dict_list: list[Dicionario]):
    try:
        # open doc
        doc = pymupdf.open(file_pdf)
        pages_on_file = len(doc)

        # get first page
        index_page = doc[0]

        for dic in my_dict_list:
            # find dictionary code on first page text
            posicoes = index_page.search_for(dic.code)
            print(posicoes)
            if posicoes:
                print(f"✅ Link: {dic.code} to page {dic.page}")
                index_page.add_squiggly_annot(posicoes)
                index_page.insert_link({
                    "from": posicoes[0],
                    "kind": pymupdf.LINK_GOTO,
                    "page": dic.page
                })

        # create links to return top
        text_back = 'Voltar'
        for i in range(1,pages_on_file):
            page_to_link = doc[i]
            point = pymupdf.Point(page_to_link.rect.width - 45, 10)
            # create a Shape to draw on
            shape = page_to_link.new_shape()
            shape.insert_text(point, text_back, color=(1,0,0)) # red
            shape.commit()
            area_back = page_to_link.search_for(text_back)
            page_to_link.insert_link({
                "from": area_back[0],
                "kind": pymupdf.LINK_GOTO,
                "page": 0
            })



        # save doc
        my_new_file = os.path.join(os.path.dirname(file_pdf), final_new_file)
        doc.save(my_new_file)
        doc.close()
        print(f"\n✅ Arquivo salvo com links: {my_new_file}")
    except Exception as e:
        print(f"❌ ERRO: {e}")


# ----------
# ----------
# ----------


fixed_file_pdf = '/home/heitor/PycharmProjects/Exemplo/testes/DEMONSTRATIVO DE DESPESAS.pdf'
fixed_anexos_path = '/home/heitor/PycharmProjects/Exemplo/testes/anexos'

listPdfFiles = find_pdfs_on_folder(fixed_anexos_path)
print( f'Found {len(listPdfFiles)} itens' )

new_file_path = merge_pdfs(fixed_file_pdf, fixed_anexos_path, listPdfFiles)
create_links(new_file_path, listPdfFiles)

print( f'Clean tmp files...' )
delete_temp_file(new_file_path)

print( f'Done!' )