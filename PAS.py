# -*- coding: UTF-8 -*-
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import win32com
from win32com.client import Dispatch, constants
from Arrays import *
import os


all = info()
# use creds to create a client to interact with the Google Drive API
scope = ['https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)
sheet = client.open("Clinic Rule for new patient Form Responses").sheet1
max_row = len(sheet.get_all_values())
for i in range(1,max_row):
    list_of_hashes = sheet.row_values(i + 1)
    all.up_name(list_of_hashes[1])
    all.up_sig(list_of_hashes[2])
    all.up_date(list_of_hashes[3])

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.


# Extract and print all of the values
#
# max_cols = len(sheet.get_all_values()[0])
title=sheet.row_values(1)
# list_of_hashes = sheet.get_all_records()
col= sheet.col_values(1)
print(title)
print(sheet.row_values(2))
print(max_row)
cur_path = os.getcwd()
template_path = cur_path + '\PAS.docx'

w = win32com.client.Dispatch('Word.Application')
w.Visible = True

doc = w.Documents.Open(FileName=template_path)
w.Selection.Find.ClearFormatting()
w.Selection.Find.Replacement.ClearFormatting()

New_Name = "jvon_1"
New_Sig = "jvon_2"
New_Date = "jvon_3"


for i in range(1,max_row):
    #1
    Old_Name, New_Name = New_Name, all.Name[i-1]
    w.Selection.Find.Execute(Old_Name, False, False, False, False, False, True, 1, True, New_Name, 49)
    #2
    Old_Sig, New_Sig = New_Sig, all.Sig[i - 1]
    w.Selection.Find.Execute(Old_Sig, False, False, False, False, False, True, 1, True, New_Sig, 49)
    #3
    Old_Date, New_Date = New_Date, all.Date[i-1]
    w.Selection.Find.Execute(Old_Date, False,False,False,False,False,True,1,True,New_Date,49)


    Name = all.Name[i-1]
    doc.SaveAs(cur_path+ "\Jvon\Patient A Rules/"+Name+".doc")

doc.Close(0)
w.Documents.Close()
w.Quit()
