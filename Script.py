from turtle import title
import openpyxl
import pandas as pd
from openpyxl import load_workbook, workbook
from openpyxl.utils import get_column_letter




archivo_excel = pd.read_excel('C:/Users/e.sarmiento/Desktop/Excel documents/team-members/team-members.xlsx')
data =(archivo_excel[['Date','User ID','Name','Email']])

data.to_excel('C:/Users/e.sarmiento/Desktop/Excel documents/master2.xlsx',startrow=0,sheet_name='Report',index=False,header=False)

#archivo_excel_master =pd.read_excel('C:/Users/e.sarmiento/Desktop/Excel documents/master.xlsx')


#wb = load_workbook('C:/Users/e.sarmiento/Desktop/Excel documents/master.xlsx')
#ws = wb.active
#print(wb)

#wb2 = load_workbook('C:/Users/e.sarmiento/Desktop/Excel documents/team-members/team-members-export-20220505142801_2022-05-04_2022-05-04.xlsx')
#ws2 = wb2['Sheet1']

#for fila in range (2, ws2.max_row+1):
 # for colunma in range (1,5):
  #    char = get_column_letter(colunma)
   #   data = (ws2[char+str (fila)].value)
    #  print(data)

