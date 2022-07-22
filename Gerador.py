from PyPDF2 import PdfReader
from docx import Document
from docx.shared import Pt

reader = PdfReader("etiquetas.pdf")
text = ""
pag = 0

document = Document()

for page in reader.pages:
  text+= page.extract_text() + "\n"
  pag += 1

textoApagar = "|-APAGAR-|_|-APAGAR-|_|-APAGAR-|_|-APAGAR-|"
correcoes = [
              "Locação Maq",
              "FORMALIZAÇÃO EMPRESA",
              "Alvará local.Funcion.CNPJ",
              "005 -SECFAZ/Fiscalização",
              "004 -SECFAZ/Arrecadação",
              "015-Patrulha Agricola",
              "028-FORMALIZAÇÃO EMPRESAS-REDESIM"
            ]

for string in correcoes:
  text = text.replace(string, textoApagar)

text = text.replace("Documento..: ", "")

print(text)

print(f"{text.count(textoApagar)} etiquetas foram identificadas como 'para apagar'")

style = document.styles['Normal']
font = style.font
font.name = 'Cambria (Corpo)'
font.size = Pt(10)

p = document.add_paragraph(text)

p.style = document.styles['Normal']

document.save("etiqueta.docx")
