import comtypes.client # type: ignore
import os

caminho_apresentacoes = "C:/Users/ruan-d.silva/Desktop/MT-Scripts/py-pptx-pdf/apresentacoes/"  # Altere para o caminho onde estão as apresentações
caminho_pdfs = "C:/Users/ruan-d.silva/Desktop/MT-Scripts/py-pptx-pdf/pdfs/"  # Altere para o caminho onde estão os pdfs

def convert_pptx_to_pdf(input_file, output_file):
    if not os.path.isfile(input_file):
        print(f"Input file does not exist: {input_file}")
        return
    
    input_file = os.path.abspath(input_file)
    output_file = os.path.abspath(output_file)
    
    try:
        print(f"Convertendo {input_file} para {output_file}")
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1
        presentation = powerpoint.Presentations.Open(f'"{input_file}"')
        presentation.SaveAs(f'"{output_file}"', 32)  # 32 is the enum for pdf format
        presentation.Close()
        powerpoint.Quit()
        print(f"Arquivo {input_file} convertido para pdf com sucesso!")
    except Exception as e:
        print(f"An error occurred: {e}")
        quit(1)

contador_pptx = 0
for file in os.listdir(caminho_apresentacoes):
    if file.endswith(".pptx"):
        contador_pptx += 1
        convert_pptx_to_pdf(os.path.join(caminho_apresentacoes, file), os.path.join(caminho_pdfs, file.replace(".pptx", ".pdf")))
print(f"Total de apresentações convertidas: {contador_pptx}")