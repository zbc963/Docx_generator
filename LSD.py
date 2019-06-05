# -*- coding: UTF-8 -*-
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import win32com
from win32com.client import Dispatch, constants
from Arrays import *
import os


all = info()
all.up_date("123")
# use creds to create a client to interact with the Google Drive API
scope = ['https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)
sheet = client.open("Welcome to Lotus Smile Dental Form Responses").sheet1
max_row = len(sheet.get_all_values())
for i in range(1,max_row):
    list_of_hashes = sheet.row_values(i + 1)
    all.up_date(list_of_hashes[22])
    all.up_name(list_of_hashes[1]+"    "+list_of_hashes[2])
    all.up_gen(list_of_hashes[3])
    all.up_address(list_of_hashes[4])
    all.up_phone(list_of_hashes[5])
    all.up_email(list_of_hashes[6])
    all.up_referred(list_of_hashes[7])
    all.up_prevdent(list_of_hashes[8])
    all.up_dentalcover(list_of_hashes[9])
    all.up_sensitivity(list_of_hashes[10])
    all.up_gum_bleed(list_of_hashes[11])
    all.up_grind(list_of_hashes[12])
    all.up_cracking(list_of_hashes[13])
    all.up_emergencycon(list_of_hashes[14])
    all.up_emergencypho(list_of_hashes[15])
    all.up_relationship(list_of_hashes[16])
    all.up_physician(list_of_hashes[17])
    all.up_phyphone(list_of_hashes[18])
    all.up_medicondition(list_of_hashes[19])
    all.up_mediother(list_of_hashes[20])
    all.up_sig(list_of_hashes[21])
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
template_path = cur_path + '\LSD.docx'

w = win32com.client.Dispatch('Word.Application')
w.Visible = True

doc = w.Documents.Open(FileName=template_path)
w.Selection.Find.ClearFormatting()
w.Selection.Find.Replacement.ClearFormatting()
New_Date = "jvon_1"
New_Name = "jvon_2"
New_Gen = "jvon_3"
New_Address = "jvon_4"
New_Phone = "jvon_5"
New_Email = "jvon_6"
New_Referred = "jvon_7"
New_Prevdent = "jvon_8"
New_Dentalcover = "jvon_9"
New_Sensitivity = "jvon__10"
New_Gum_Bleed = "jvon__11"
New_Grind = "jvon__12"
New_Cracking = "jvon__13"
New_Emergencycon = "jvon__14"
New_Emergencypho = "jvon__15"
New_Relationship = "jvon__16"
New_Physician = "jvon__17"
New_Phyphone = "jvon__18"
New_Medicondition = "jvon__19"
New_Mediother = "jvon__20"
New_Sig = "jvon__21"



for i in range(1,max_row):

    #1
    Old_Date, New_Date = New_Date, all.Date[i-1]
    w.Selection.Find.Execute(Old_Date, False,False,False,False,False,True,1,True,New_Date,49)
    #2
    Old_Name, New_Name = New_Name, all.Name[i-1]
    w.Selection.Find.Execute(Old_Name, False, False, False, False, False, True, 1, True, New_Name, 49)
    #3
    Old_Gen, New_Gen = New_Gen, all.Gen[i-1]
    w.Selection.Find.Execute(Old_Gen, False, False, False, False, False, True, 1, True, New_Gen, 49)
    #4
    Old_Address, New_Address = New_Address, all.Address[i - 1]
    w.Selection.Find.Execute(Old_Address, False, False, False, False, False, True, 1, True, New_Address, 49)
    #5
    Old_Phone, New_Phone = New_Phone, all.Phone[i - 1]
    w.Selection.Find.Execute(Old_Phone, False, False, False, False, False, True, 1, True, New_Phone, 49)
    #6
    Old_Email, New_Email = New_Email, all.Email[i-1]
    w.Selection.Find.Execute(Old_Email, False, False, False, False, False, True, 1, True, New_Email, 49)
    #7
    Old_Referred, New_Referred = New_Referred, all.Referred[i - 1]
    w.Selection.Find.Execute(Old_Referred, False, False, False, False, False, True, 1, True, New_Referred, 49)
    #8
    Old_Prevdent, New_Prevdent = New_Prevdent, all.Prevdent[i - 1]
    w.Selection.Find.Execute(Old_Prevdent, False, False, False, False, False, True, 1, True, New_Prevdent, 49)
    #9
    Old_Dentalcover, New_Dentalcover = New_Dentalcover, all.Dentalcover[i - 1]
    w.Selection.Find.Execute(Old_Dentalcover, False, False, False, False, False, True, 1, True, New_Dentalcover, 49)
    #10
    Old_Sensitivity, New_Sensitivity = New_Sensitivity, all.Sensitivity[i - 1]
    w.Selection.Find.Execute(Old_Sensitivity, False, False, False, False, False, True, 1, True, New_Sensitivity, 49)
    #11
    Old_Gum_Bleed, New_Gum_Bleed = New_Gum_Bleed, all.Gum_Bleed[i - 1]
    w.Selection.Find.Execute(Old_Gum_Bleed, False, False, False, False, False, True, 1, True, New_Gum_Bleed, 49)
    #12
    Old_Grind, New_Grind = New_Grind, all.Grind[i - 1]
    w.Selection.Find.Execute(Old_Gum_Bleed, False, False, False, False, False, True, 1, True, New_Gum_Bleed, 49)
    #13
    Old_Cracking, New_Cracking = New_Cracking, all.Cracking[i - 1]
    w.Selection.Find.Execute(Old_Cracking, False, False, False, False, False, True, 1, True, New_Cracking, 49)
    #14
    Old_Emergencycon, New_Emergencycon = New_Emergencycon, all.Emergencycon[i - 1]
    w.Selection.Find.Execute(Old_Emergencycon, False, False, False, False, False, True, 1, True, New_Emergencycon, 49)
    #15
    Old_Emergencypho, New_Emergencypho = New_Emergencypho, all.Emergencypho[i - 1]
    w.Selection.Find.Execute(Old_Emergencypho, False, False, False, False, False, True, 1, True, New_Emergencypho, 49)
    #16
    Old_Relationship, New_Relationship = New_Relationship, all.Relationship[i - 1]
    w.Selection.Find.Execute(Old_Relationship, False, False, False, False, False, True, 1, True, New_Relationship, 49)
    #17
    Old_Physician, New_Physician = New_Physician, all.Physician[i - 1]
    w.Selection.Find.Execute(Old_Physician, False, False, False, False, False, True, 1, True, New_Physician, 49)
    #18
    Old_Phyphone, New_Phyphone = New_Phyphone, all.Phyphone[i - 1]
    w.Selection.Find.Execute(Old_Phyphone, False, False, False, False, False, True, 1, True, New_Phyphone, 49)
    #19
    Old_Medicondition, New_Medicondition = New_Medicondition, all.Medicondition[i - 1]
    w.Selection.Find.Execute(Old_Medicondition, False, False, False, False, False, True, 1, True, New_Medicondition, 49)
    #20
    Old_Mediother, New_Mediother = New_Mediother, all.Mediother[i - 1]
    w.Selection.Find.Execute(Old_Mediother, False, False, False, False, False, True, 1, True, New_Mediother, 49)
    #21
    Old_Sig, New_Sig = New_Sig, all.Sig[i - 1]
    w.Selection.Find.Execute(Old_Sig, False, False, False, False, False, True, 1, True, New_Sig, 49)

    Name = all.Name[i-1]
    doc.SaveAs(cur_path+"\LotusSmileDental/"+Name+".doc")

doc.Close(0)
w.Documents.Close()
w.Quit()
