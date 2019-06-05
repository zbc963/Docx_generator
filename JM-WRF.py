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
sheet = client.open("Walk-in Registration Form Form Responses").sheet1
max_row = len(sheet.get_all_values())
for i in range(1,max_row):
    list_of_hashes = sheet.row_values(i + 1)
    all.up_date(list_of_hashes[13])
    all.up_name(list_of_hashes[1]+"    "+list_of_hashes[2])
    all.up_phone(list_of_hashes[3])
    all.up_email(list_of_hashes[4])
    all.up_address(list_of_hashes[5])
    all.up_fax(list_of_hashes[6])
    all.up_physician(list_of_hashes[7])
    all.up_phyphone(list_of_hashes[8])
    all.up_emergencycon(list_of_hashes[10])
    all.up_emergencypho(list_of_hashes[11])
    all.up_sig(list_of_hashes[12])

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
template_path = cur_path + '\WRF.docx'

w = win32com.client.Dispatch('Word.Application')
w.Visible = True

doc = w.Documents.Open(FileName=template_path)
w.Selection.Find.ClearFormatting()
w.Selection.Find.Replacement.ClearFormatting()
New_Date = "jvon__14"
New_Name = "jvon_2"
New_Phone = "jvon_4"
New_Email = "jvon_5"
New_Address = "jvon_6"
New_Physician = "jvon__8"
New_Phyphone = "jvon__9"
New_Fax = "jvon__10"
New_Emergencycon = "jvon__11"
New_Emergencypho = "jvon__12"
New_Sig = "jvon__13"



for i in range(1,max_row):

    #1
    Old_Date, New_Date = New_Date, all.Date[i-1]
    w.Selection.Find.Execute(Old_Date, False,False,False,False,False,True,1,True,New_Date,49)
    #2
    Old_Name, New_Name = New_Name, all.Name[i-1]
    w.Selection.Find.Execute(Old_Name, False, False, False, False, False, True, 1, True, New_Name, 49)
    #3

    #4
    Old_Phone, New_Phone = New_Phone, all.Phone[i - 1]
    w.Selection.Find.Execute(Old_Phone, False, False, False, False, False, True, 1, True, New_Phone, 49)
    #5
    Old_Email, New_Email = New_Email, all.Email[i - 1]
    w.Selection.Find.Execute(Old_Email, False, False, False, False, False, True, 1, True, New_Email, 49)
    #6
    Old_Address, New_Address = New_Address, all.Address[i - 1]
    w.Selection.Find.Execute(Old_Address, False, False, False, False, False, True, 1, True, New_Address, 49)

    #7

    #8
    Old_Physician, New_Physician = New_Physician, all.Physician[i - 1]
    w.Selection.Find.Execute(Old_Physician, False, False, False, False, False, True, 1, True, New_Physician, 49)
    #9
    Old_Phyphone, New_Phyphone = New_Phyphone, all.Phyphone[i - 1]
    w.Selection.Find.Execute(Old_Phyphone, False, False, False, False, False, True, 1, True, New_Phyphone, 49)
    #10
    Old_Fax, New_Fax = New_Fax, all.Fax[i - 1]
    w.Selection.Find.Execute(Old_Fax, False, False, False, False, False, True, 1, True, New_Fax, 49)
    #11
    Old_Emergencycon, New_Emergencycon = New_Emergencycon, all.Emergencycon[i - 1]
    w.Selection.Find.Execute(Old_Emergencycon, False, False, False, False, False, True, 1, True, New_Emergencycon, 49)
    #12
    Old_Emergencypho, New_Emergencypho = New_Emergencypho, all.Emergencypho[i - 1]
    w.Selection.Find.Execute(Old_Emergencypho, False, False, False, False, False, True, 1, True, New_Emergencypho, 49)

    #13
    Old_Sig, New_Sig = New_Sig, all.Sig[i - 1]
    w.Selection.Find.Execute(Old_Sig, False, False, False, False, False, True, 1, True, New_Sig, 49)

    Name = all.Name[i-1]
    doc.SaveAs(cur_path+"\Jvon\Walk-in Registration Form/"+Name+".doc")

doc.Close(0)
w.Documents.Close()
w.Quit()
