from csv import writer
from turtle import title
import openpyxl
import pandas as pd
from openpyxl import load_workbook, workbook
from openpyxl.utils import get_column_letter
import xlwings as ws


teamsExcel = 'C:/Users/e.sarmiento/Desktop/Excel documents/team-members/team-members.xlsx'
masterExcel = 'C:/Users/e.sarmiento/Desktop/Excel documents/master05.xlsx'


#se lee team members y masterExcel
readTeamsExcel = pd.read_excel(teamsExcel)
readMasterExcel = pd.read_excel(masterExcel, sheet_name='BD-Login')
#se almacena teams
dataT =pd.DataFrame(readTeamsExcel[['Date','User ID','Name','Email']])
#se lee master para obtener fila de bd-login
dataM = pd.DataFrame(readMasterExcel)
datafiltered = dataM.dropna(subset=['Date'])
lastindex =datafiltered.index[-1]
 #se pega el team member en master
dataT.to_excel(masterExcel,startrow=lastindex+2,index=False,header=False,sheet_name='BD-Login')




 




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

