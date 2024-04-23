import openpyxl
from datetime import datetime

sheetPath = '/home/szabo/Documents/Ficha Frequencia 2024 - Henrique Szabo.xlsx'
pages = ['Instruções', 'JAN', 'FEV', 'MAR', 'ABR', 'MAIO', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']
text = None
workbook = openpyxl.load_workbook(sheetPath)

for i in range(len(pages)):
    if i == int(datetime.now().strftime('%m')):
        sheet = workbook[pages[i]]

while(text == None):
    text = input("Digite o feito do dia: ")

try:
    for row in sheet.iter_rows(18, 48):
        data_celula = row[0].value
        if isinstance(data_celula, datetime):
            data_celula = data_celula.strftime('%d/%m/%Y')
        if data_celula == datetime.now().strftime('%d/%m/%Y'):
            row[1].value = text
            row[4].value = 6
except Exception as e:
    print("Erro ao salvar dados:", e)
finally:
    try:
        workbook.save(sheetPath)
    except Exception as e:
        print("Erro ao salvar workbook:", e)
    finally:
        workbook.close()