import os;
import pandas as pd;
import smtplib;
from pretty_html_table import build_table;
from email.message import EmailMessage;

EMAIL_ADDRESS = 'test@gmail.com'
EMAIL_PASSWORD = 'pass123'

def reportSplitIntoChunks(dataFrame):
  nameColumn = dataFrame['Name']
  nameList = list(nameColumn)
  return splitData(nameList)


def splitData(arr):
  result = []
  temp = []
  for index,element in enumerate(arr):   
    temp.append(index)
    if(pd.isna(element)):
      result.append(temp[:])
      temp = []
  return result

def handleChunk(arrOfIndexes, df):
  dataOfInterest = df[arrOfIndexes[0]:(arrOfIndexes[-1]+1)]
  sendeeEmail = df['Send To'][arrOfIndexes[0]]
  html = dataFrameToHtml(dataOfInterest[['Date', 'Name', 'Water: previous measuring', 
  'Water: current measuring', 'm3', 'Summ', 'Losses', 'Overall']])  
  server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
  server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
  sendEmail(sendeeEmail, 'Monthly Report', html, server)
  

def dataFrameToHtml(df):
  body = """
  <html>
  <head>
  </head>
  <body>
    {0}
  </body>
  </html>
  """.format(build_table(df, 'blue_light'))
  return body

def sendEmail(destEmail, subject, message, server):
  msg = EmailMessage()
  msg['Subject'] = subject
  msg['From'] = EMAIL_ADDRESS
  msg['To'] = destEmail
  msg.set_content('Please enable html in emails')
  msg.add_alternative(message, subtype='html')
  server.send_message(msg)

fileName = input('Please enter the name of excel file to be handled: ') + '.xlsx'
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
try:
  df = pd.read_excel(fileName)
  chunks = reportSplitIntoChunks(df)
  for chunk in chunks:
    handleChunk(chunk, df)
except:
  print('check if the file is in the same directory, or password, or internet or idk')
  input()
finally:
  server.quit()


