import openpyxl
from docxtpl import DocxTemplate
import datetime

#Carregar os dados do arquivo excel na sua localização
path = "C:\\Users\\fgsjd\\OneDrive\\Área de Trabalho\\Developer Career\\converter_excel_docx\\student_data.xlsx" # Caminho até o arquivo excel.
workbook = openpyxl.load_workbook(path)
sheet = workbook.active

list_values = list(sheet.values)
print(list_values)
#--------------------------------------------------------------------------------------------------------
# Até esse trecho o código escrito foir para abrir o arquivo no formato excel


#variável que armazenar o modelo ou "template onde ele vai "juntar" ou como chamamos de concatenar nos espaços referidos.
doc = DocxTemplate("certificate.docx")

# # Criando um indíce para percorrer toda a tupla
for value_tuple in list_values[1:]:
    doc.render({"name": value_tuple[0],
                "course": value_tuple[1],
                "date": value_tuple[2],
                "instructor": value_tuple[3]})
    
    #Aqui ele irá gerar o arquivo contendo o formato em .docx (o famoso word)
    doc_name_final = "certificado" + value_tuple[0] + value_tuple[1] + ".docx"
    doc.save(doc_name_final) #E aqui ele salva o arquivo
    

    
    
    