from pathlib import Path
import PySimpleGUI as sg
import os
from PyPDF2 import PdfReader, PdfWriter, PdfMerger

 
# Mescla páginas de diferentes arquivos PDF
def merge_pdf(selected_pdf_files, filesDirectory, pdf_file_name):
   
    # separa o nome de cada arquivo pdf usando ; como separador
    pdf_files = selected_pdf_files.split(";")

 
    merger = PdfMerger()
    for pdf_file in pdf_files:
        merger.append(pdf_file)
    path_to_file = "{}/{}.pdf".format(filesDirectory, pdf_file_name)
    merger.write(path_to_file)
    merger.close()
    sg.popup("Páginas unitizadas com sucesso!")

 
# Separa cada página de um arquivo PDF com diversas páginas em arquivos individuais
def split_pdf(path, pdfName, filesDirectory):

 
    pdf = PdfReader(path)
    for page in range(len(pdf.pages)):
        pdf_writer = PdfWriter()
        pdf_writer.add_page(pdf.pages[page])
        output_filename = '{}/{}_pg_{}.pdf'.format(filesDirectory,
            pdfName, page+1)
        with open(output_filename, 'wb') as out:
            pdf_writer.write(out)
    sg.popup("Separação concluída!")

 
# Protege o PDF contra impressões
def noPrint_pdf(path, npPdf, filesDirectory):
    pdf = PdfReader(path)
    pdf_writer = PdfWriter()
    for i in range(len(pdf.pages)):
        page = pdf.getPage(i)
        pdf_writer.addPage(page)
    pdf_writer.encrypt(user_pwd='', owner_pwd=None, use_128bit=True, permissions_flag=10)
    output_filename = '{}/{}.pdf'.format(filesDirectory, npPdf)
    with open(output_filename, 'wb') as file:
        pdf_writer.write(file)
    sg.popup("Bloqueio concluído!")

 
#  janela inicial
def main_window():

 
    menu_def = [["Informações",["Sobre","Ajuda","Histórico de versões"]]]
   
    layout = [
        [sg.MenubarCustom(menu_def, tearoff=False)],
        [sg.Text("Escolha uma das opções a seguir:")],
        [sg.Exit("Sair"), sg.Button("Mesclar"), sg.Button("Separar"), sg.Button("Bloquear Impressão")],
    ]

 
    window = sg.Window("PDF Manip", layout)

 
    while True:
        event, values = window.read()
        print(event, values)
        if event in (sg.WINDOW_CLOSED, "Sair"):
            break
        if event == "Sobre":
            window.disappear()
            sg.popup("PDF Manip\nVersão 2.0.0\nDesenvolvido por Bruno K. Kajita")
            window.reappear()
        if event == "Ajuda":
            window.disappear()
            sg.popup("Escolhar 'Mesclar' para unir diferentes arquivos PDF em um único.\n Escolha 'Separar' para criar um arquivo para cada página.")
            window.reappear()
        if event == "Histórico de versões":
            window.disappear()
            sg.popup("Versão 1.0.0 - Inicial\nVersão 2.0.0 - PDF não imprimível \nVersão 2.0.1 - Correções")
            window.reappear()
        if event == "Mesclar":
            merge_pdf(
                selected_pdf_files = sg.popup_get_file("Selecione os arquivos",
                    file_types=(("Arquivos PDF","*.pdf*"),), multiple_files=True),
                pdf_file_name = sg.popup_get_text("Informe o Nome do Arquivo:"),
                filesDirectory = sg.popup_get_folder("Selecione a pasta de destino:")
            )
        if event == "Separar":
            split_pdf(
                path = sg.popup_get_file("Selecione o arquivo:", file_types=(("Arquivos PDF","*.pdf"),)),
                pdfName = sg.popup_get_text("Informe o nome dos arquivos:"),
                filesDirectory = sg.popup_get_folder("Selecione a pasta de destino:")
            )
        if event == "Bloquear Impressão":
            noPrint_pdf(
                path = sg.popup_get_file("Selecione o arquivo:", file_types=(("Arquivos PDF","*.pdf"),)),
                npPdf = sg.popup_get_text("Informe o nome dos arquivos:"),
                filesDirectory = sg.popup_get_folder("Selecione a pasta de destino:")
            )
           
    window.close()


 
if __name__ == "__main__":
    sg.theme('DarkTeal10')
    main_window()