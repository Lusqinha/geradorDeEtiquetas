from PyPDF2 import PdfReader
from docx import Document
from docx.shared import Pt
from glob import glob


pdFile = glob('*.pdf')[0]

reader = PdfReader(pdFile)
text = ""
paginas = 0

textoApagar = "|-APAGAR-|"
correcoes = [
              "Locação Maq",
              "FORMALIZAÇÃO EMPRESA",
              "Alvará local.Funcion.CNPJ",
              "SECFAZ/Fiscalização",
              "SECFAZ/Arrecadação",
              "018 -SECFAZ / Fiscalização",
              "010-Isenção da Taxa de Entulho",
              "Taxa de Entulho",
              "AUTONOMO",
              "Construção de Tumulo",
              "015-Patrulha Agricola",
              "Patrulha",
              "028-FORMALIZAÇÃO EMPRESAS-REDESIM",
              "029-FORMALIZAÇÃO EMPRESA - MEI",
              "030-ALTERAÇÃO - MEI",
              "030 -ALTERAÇÃO",
              "031-ALTERAÇÃO - REDESIM",
              "031 -ALTERAÇÃO - REDESIM",
              "Baixa de lotação",
              "Procuradoria Geral",
              "PGM"
            ]

document = Document()
style = document.styles['Normal']
font = style.font
font.name = 'Cambria (Corpo)'
font.size = Pt(11)

for page in reader.pages:
  text+= page.extract_text() + "\n"
  paginas += 1

for string in correcoes:
  text = text.replace(string, textoApagar)

text = text.replace("Documento..: ", "")

class Etiqueta:
  def __init__(self):
    self.listaEtiquetas = []
    self.protocolo = '' + '\n'
    self.assunto = ''+'\n'
    self.subassunto = ''+'\n'
    self.requerente = ''+'\n'
  
  def geradorEtiqueta(self, t):
    for line in t.split('\n'):
      if line.startswith("2022"):
        self.protocolo = line
        self.listaEtiquetas.append(self.protocolo)
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
indexDois = 4


while paginas >= 0:
  tag = etiqueta.listaEtiquetas[indexUm:indexDois]
  paginas-=1
  if textoApagar in str(tag):
    str(tag).replace(str(tag), "")
  else:
    tag = '\n'.join([i for i in tag[0:]])
    p = document.add_paragraph(tag)
    
  indexUm += 4
  indexDois += 4

p.style = document.styles['Normal']
document.save("etiqueta.docx")
