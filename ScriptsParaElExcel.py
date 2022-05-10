
from cgi import print_form
from email.header import Header
from sqlite3 import DataError
import openpyxl
import pandas as pd
from openpyxl import load_workbook, workbook
from openpyxl.utils import get_column_letter
from xlsx2csv import Xlsx2csv
from io import StringIO   
 
 
 
def read_excel(path: str, sheet_index: int) -> pd.DataFrame: 
 buffer = StringIO() 
 Xlsx2csv(path, outputencoding="utf-8").convert(buffer,sheetid=sheet_index) 
 buffer.seek(0)  
 df = pd.read_csv(buffer, low_memory=False)
 return df

def bdLoginCloudTalkCopyPaste():
 print('bdLoginCloudTalk copiado y pegado ejecutado')
 #se obtienen las direcciones de los documentos a trabajar
 teamsExcel = 'C:/Users/e.sarmiento/Desktop/Excel documents/team-members/team-members.xlsx'
 masterExcel = 'C:/Users/e.sarmiento/Desktop/Excel documents/ADH Templete.xlsx'
 agentsExcel =  'C:/Users/e.sarmiento/Desktop/Excel documents/agents/agents.xlsx'

 #se lee team members,agent y masterExcel
 readTeamsExcel = pd.read_excel(teamsExcel)
 readMasterExcelBdLogin = pd.read_excel(masterExcel, sheet_name='BD-Login')
 readMasterExcelClouldtalk = pd.read_excel(masterExcel, sheet_name='Cloudtalk')
 readagentsExcel = pd.read_excel(agentsExcel)

 #se almacenan los datos y se hacen los filtros para evitar NaN
 dataT =pd.DataFrame(readTeamsExcel[['Date','User ID','Name','Email']])
 datafilteredteams = dataT.dropna(subset=['Date'])
 #print('1RA PARTE')
 #print(datafilteredteams)
 dataT2 =pd.DataFrame(readTeamsExcel[['Group','Absence','Productive time','Unproductive time','Neutral time','Total DeskTime','Offline time','Private time','Arrived','Left','Late','Total time at work','Idle time','Extra hours before work','Extra hours after work','Hourly rate']])
 datafilteredteams2 = dataT2.dropna(subset=['Group'])
 #print ('2da parte')
 #print(datafilteredteams2)
 dataAgent =pd.DataFrame(readagentsExcel[['Date','Agent','SIP Login time (sec)','Idle time (sec)','Ringing time (sec)','Talking time (sec)','Wrap up time (sec)','Inbound Calls','Outbound Calls']])
 datafilteredAgent = dataAgent.dropna(subset=['Date'])
 #print(dataAgent)

 #se lee master para obtener fila de bd-login
 dataM = pd.DataFrame(readMasterExcelBdLogin)
 datafiltered = dataM.dropna(subset=['Date'])
 
 if datafiltered.empty:
  print('empty')
  lastindex=1

 else:
  print('no empty')
  lastindex =datafiltered.index[-1] +2

 #se lee master para obtener fila de cloudtalk
 lastindexFound = pd.DataFrame(readMasterExcelClouldtalk)
 
 print(lastindexFound)
 #if datafiltered2['Date'].empty :
  #print('empty')
  #lastindexCloud = 1
 #else:
   #print('no empty')
 lastindexCloud = lastindexFound.index[-1]+2
 
  

 #se pega el team y cloud en master

 with pd.ExcelWriter (masterExcel,mode="a",
    engine="openpyxl",
    if_sheet_exists="overlay") as writer:
   datafilteredteams.to_excel(writer,startrow=lastindex,index=False,header=False,sheet_name='BD-Login')
   datafilteredteams2.to_excel(writer,startrow=lastindex,index=False,header=False,sheet_name='BD-Login',startcol=5)
   datafilteredAgent.to_excel(writer,index=False,startrow=lastindexCloud,header=False,sheet_name='Cloudtalk')
 print('bdLoginCloudTalk copiado y pegado finalizado')

def scheduleFilling():
 try:
  print('scheduleFilling copiado y pegado ejecutado')
   #se obtiene las direcciones de los documentos a trabajar
  masterExcel = 'C:/Users/e.sarmiento/Desktop/Excel documents/ADH Templete.xlsx'
  ScheduleExcel='C:/Users/e.sarmiento/Desktop/Excel documents/Cyracom Schedule/01. Cyracom Schedule May 02 - May 8.xlsb'
   
   #se lee schedules y masterExcel
  readScheduleExcel = pd.read_excel(ScheduleExcel, sheet_name='Schedule', engine='pyxlsb',skiprows=range(0,5))
  #readMasterExcelSchedule = pd.read_excel(masterExcel, sheet_name='Schedules')
   
   #se almacenan los datos y se hacen los filtros para evitar NaN
  
  dataT =pd.DataFrame(readScheduleExcel.iloc[:,[3,4,5]])

  datafilteredT = dataT.dropna(how='all')
 
  print('1ra parte')
  print(datafilteredT)
  dataT2=pd.DataFrame(readScheduleExcel.iloc[:,[7,8,9,10,11,12,13,14,15,16,17]])
  datafilteredT2 = dataT2.dropna(how='all')
  print('2da parte')
  print(datafilteredT2)					


   #se pega el schedule en master
  with pd.ExcelWriter (masterExcel,mode="a",engine="openpyxl",if_sheet_exists="overlay") as writer:
   datafilteredT.to_excel(writer,startrow=1,index=False,header=False ,sheet_name='Schedules',startcol=2)
   datafilteredT2.to_excel(writer,startrow=1,index=False,header=False,sheet_name='Schedules',startcol=6)

  print('scheduleFilling copiado y pegado finalizado')
 except DataError:
   print(DataError +'error')

def arreglarCorreoSchedule():
  print('arreglar corre scheudule ejecutado')
   #se obtiene las direcciones de los documentos a trabajar
  masterExcel = 'C:/Users/e.sarmiento/Desktop/Excel documents/master05.xlsx'
   #se lee masterExcel
  readMasterExcelSchedule = pd.read_excel(masterExcel, sheet_name='Schedules')

   #se almacenan los datos y se hacen los cambios deseados
  
  dataT =pd.DataFrame(readMasterExcelSchedule.iloc[:,[17]])

  print(dataT)

def ArreglarFechaTemp():
 print('Arrglar fecha ejecutado')
#se obtienen las direcciones de los documentos a trabajar
 teamsExcel = 'C:/Users/e.sarmiento/Desktop/Excel documents/team-members/team-members.xlsx'
 masterExcel = 'C:/Users/e.sarmiento/Desktop/Excel documents/master05.xlsx'
 agentsExcel =  'C:/Users/e.sarmiento/Desktop/Excel documents/agents/agents.xlsx'

  #se lee team members,agent y masterExcel
 readTeamsExcel = pd.read_excel(teamsExcel)
 readMasterExcelBdLogin = pd.read_excel(masterExcel, sheet_name='BD-Login')
 readMasterExcelClouldtalk = pd.read_excel(masterExcel, sheet_name='Cloudtalk')
 readagentsExcel = pd.read_excel(agentsExcel)

  #se almacenan los datos y se hacen los filtros para evitar NaN
 dataT =pd.DataFrame(readMasterExcelBdLogin[['Date','User ID','Name','Email']])
 datafilteredteams = dataT.dropna(subset=['Date'])
 datafilteredteams.style.format({'Date':'{:,2f}'})
 print(datafilteredteams)

bdLoginCloudTalkCopyPaste()






   