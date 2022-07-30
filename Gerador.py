from PyPDF2 import PdfReader
from docx import Document
from docx.shared import Pt

reader = PdfReader("etiquetas.pdf")
text = ""
paginas = 0

document = Document()
style = document.styles['Normal']
font = style.font
font.name = 'Cambria (Corpo)'
font.size = Pt(10)

for page in reader.pages:
  text+= page.extract_text() + "\n"
  paginas += 1


textoApagar = "|-APAGAR-|"
correcoes = [
              "Locação Maq",
              "FORMALIZAÇÃO EMPRESA",
              "Alvará local.Funcion.CNPJ",
              "SECFAZ/Fiscalização",
              "SECFAZ/Arrecadação",
              "018 -SECFAZ / Fiscalização",
              "010-Isenção da Taxa de Entulho",
              "015-Patrulha Agricola",
              "028-FORMALIZAÇÃO EMPRESAS-REDESIM",
              "029-FORMALIZAÇÃO EMPRESA - MEI",
              "030-ALTERAÇÃO - MEI",
              "031-ALTERAÇÃO - REDESIM"
            ]

for string in correcoes:
  text = text.replace(string, textoApagar)

text = text.replace("Documento..: ", "")

# ------------Código de teste --------------- #

class Etiqueta:
  def __init__(self):
    self.listaEtiquetas = []
    self.assunto = ''+'\n'
    self.subassunto = ''+'\n'
    self.requerente = ''+'\n'
  
  def __str__(self):
    return {"assunto....":self.assunto.removeprefix("Assunto....: "), "subassunto.":self.subassunto.removeprefix("Subassunto.: "), "requerente.":self.requerente.removeprefix("Requerente.: ")}
  
  def geradorEtiqueta(self, t):
    for line in t.split('\n'):
      if line.startswith("Assu"):
        self.assunto = line
        self.listaEtiquetas.append(self.assunto)
      elif line.startswith("Suba"):
        self.subassunto = line
        self.listaEtiquetas.append(self.subassunto)
      elif line.startswith("Requ"):
        self.requerente = line
        self.listaEtiquetas.append(self.requerente)
    return self.listaEtiquetas
    
  
etiqueta = Etiqueta()
etiqueta.geradorEtiqueta(text)

indexUm = 0
indexDois = 3

while paginas >= 0:
  tag = etiqueta.listaEtiquetas[indexUm:indexDois]
  paginas-=1
  if textoApagar in str(tag):
    str(tag).replace(str(tag), "")
  else:
    tag = '\n'.join([i for i in tag[0:]])
    p = document.add_paragraph(tag)
    
  indexUm += 3
  indexDois += 3

p.style = document.styles['Normal']
document.save("etiqueta.docx")
