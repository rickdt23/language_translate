import openpyxl
from translate import Translator

#open workbook
wb = openpyxl.load_workbook('Python Translate.xlsx')
#label individual sheets
ws1 = wb['Sheet1']
ws2 = wb['Sheet2']

#record cell values

A1 = ws1['A1'].value
A2 = None
if ws1['A2'].value != 0:
    A2 = ws1['A2'].value

#print English words
print(A1)
if A2 != None:
    print(A2)

#translate text
translater = Translator(from_lang='English', to_lang='Spanish')

translation = translater.translate(A1)
if A2 != None:
    translation2 = translater.translate(A2)

#save translated values
ws2['A1'].value = translation
if A2 != None:
    ws2['A2'].value = translation2

#print Spanish words
print(translation)
if A2 != None:
    print(translation2)

#save workbook
wb.save('Python Translate.xlsx')

