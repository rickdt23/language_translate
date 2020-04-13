
import openpyxl

from translate import Translator

wb = openpyxl.load_workbook('Python Translate.xlsx')
ws = wb['Sheet1']
A = ws['A1'].value
print(A)
translater= Translator(from_lang='English', to_lang='Spanish')
translation= translater.translate(A)
print(translation)


#print(wb)
