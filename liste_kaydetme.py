from openpyxl import load_workbook

import pickle

som = load_workbook('sablon.xlsx')
print(som)
fileObject = open("lib.dll","wb")
pickle.dump(som, fileObject)



