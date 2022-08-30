from cgitb import text
from msilib.schema import tables
from docx import Document
import pandas as pd
tabela = pd.read_excel("Colaboradores.xlsx")

def NovoDado (adicionar,Df,NomeColuna,substituir):
    contador = 0
    for adicionar in Df[str(NomeColuna)]:
        documento = Document('geradorPython - {0} .docx'.format(contador))
        for paragrafo in documento.paragraphs:  
            paragrafo.text = paragrafo.text.replace(str(substituir),adicionar)
            documento.save('geradorPython - {0} .docx'.format(contador))
        contador = contador + 1 


 