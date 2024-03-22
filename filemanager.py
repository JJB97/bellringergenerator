#! /usr/bin/python3
# -*- coding: utf-8 -*-
from oauth2client.client import GoogleCredentials
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
from time import sleep
import io, os, gspread


env = {}
with open('./info_filemanage.txt', 'r') as f:
    infofile = f.read()
    infofile = infofile.strip()
    infofile = infofile.split('\n')
    for i in range(0,len(infofile)):
        infofile[i] = infofile[i].split(' = ')
        env[infofile[i][0]] = infofile[i][1]

jsonLocalPath = env['JSON_LOCAL_PATH']
sheetURL = env['SHEET_URL']
problemListSheetURL = env['PROBLEM_LIST_SHEET_URL']
sheetName = env['SHEET_NAME']

requestSheetCopyPath = './setup/request_sheet_copy.txt'
userListPath = './setup/userlist.txt'
problemDirPath = './problems/'


sheetColInfo = {'timeStamp':1, 'email':2, 'problemImage':3, 'answer':4, 'coverage':5, 'weight':6, 'organ':7}
gc = gspread.service_account(filename=jsonLocalPath)
store = file.Storage('storage.json')
creds = store.get()
drive = build('drive', 'v3', http=creds.authorize(Http()))

def newRequest():
    sheet = gc.open_by_url(sheetURL)
    worksheet = sheet.worksheet(sheetName)
    
    requestSheet = worksheet.get_all_values()
    requestSheet = requestSheet[1:]
    requestSheet.reverse()
    
    newRequest = []
    with open(requestSheetCopyPath, 'r') as f:
        requestSheetCopy = f.read()
    requestSheetCopy = requestSheetCopy.split('\n')
    requestSheetCopy.pop()
    requestSheetCopy.reverse()
    
    for i in requestSheet:
        temp = '\t'.join(i)
        if not (temp in requestSheetCopy):
            newRequest.append(i)
        else:
            break
    requestSheet.reverse()
    requestSheetCopy.reverse()
    newRequest.reverse()
    if len(newRequest)==0: 
        return newRequest
    with open(userListPath, 'r') as f:
        userList = f.read()
    userList = userList.split('\n')
    userList.pop()
    
    validNewRequest = []
    for i in range(len(newRequest)):
        if (newRequest[i][sheetColInfo['email']-1] in userList):
            validNewRequest.append(newRequest[i])
    
    return validNewRequest


def storeProblem(requestRow):
    file_id = requestRow[sheetColInfo['problemImage']-1]
    file_id = file_id.split('?id=')
    file_id = file_id[-1]
    extension = drive.files().get(fileId=file_id).execute()#이걸로 mimeType, name 알수있음
    extension = extension['name'].split('.')
    extension = extension[-1]
    
    directory = problemDirPath+requestRow[sheetColInfo['coverage']-1]+"/"
    filename = requestRow[sheetColInfo['answer']-1]+'  '+requestRow[sheetColInfo['timeStamp']-1]+'  '+requestRow[sheetColInfo['email']-1]+'.'+extension
    if not os.path.exists(directory):
        os.makedirs(directory)
    
    request = drive.files().get_media(fileId=file_id)
    fh = io.FileIO(directory+filename, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done == False:
        done = downloader.next_chunk()
    with open(directory+"fileList.txt", 'a') as f:
        f.write('\t'.join(requestRow))
        f.write('\t'+directory+filename)
        f.write('\n')
    return directory+filename


def updateProblemList(requestRow):
    sheet = gc.open_by_url(problemListSheetURL)
    try:
        worksheet = sheet.worksheet(requestRow[sheetColInfo['coverage']-1])
    except gspread.exceptions.WorksheetNotFound:
        title_name = requestRow[sheetColInfo['coverage']-1]
        worksheet = sheet.add_worksheet(title=title_name, rows=1, cols=5)
        colVal = ['정답(소문자로 기록됨)', '포함된 기관 (소문자로 기록됨)', '가중치(1~5)', '업로드된 문제 수']
        worksheet.append_row(colVal)
    cell_list = worksheet.findall(requestRow[sheetColInfo['answer']-1].lower())
    
    
    temp = []
    for i in cell_list:
        organ = requestRow[sheetColInfo['organ']-1].lower()
        if i.col==1 and i.row!=1 and worksheet.cell(i.row,2).value==organ and worksheet.cell(i.row,3).value==requestRow[sheetColInfo['weight']-1]:
            temp.append(i)
        
    cell_list = temp
    del temp
    
    if len(cell_list)==0:
        input_data = [requestRow[sheetColInfo['answer']-1].lower(), requestRow[sheetColInfo['organ']-1].lower(), requestRow[sheetColInfo['weight']-1], 1]
        worksheet.append_row(input_data)
    else:
        row = cell_list[0].row
        num_of_problems = worksheet.cell(row, 4).value
        num_of_problems = int(num_of_problems)+1
        worksheet.update_cell(row, 4, num_of_problems)
    return

def deleteDriveFile(requestRow):
    file_id = requestRow[sheetColInfo['problemImage']-1]
    file_id = file_id.split('?id=')
    file_id = file_id[-1]
    drive.files().delete(fileId=file_id).execute()
    return

def updateSheet(requestRow):
    sheet = gc.open_by_url(sheetURL)
    worksheet = sheet.worksheet(requestRow[sheetColInfo['coverage']-1])
    
    return

if __name__=='__main__':
    while True:
        newRequestList = newRequest()
        
        if len(newRequestList)!=0:
            for i in newRequestList:
                storeProblem(i)
                updateProblemList(i)
                #deleteDriveFile(i)
                #updateSheet(i)
                with open(requestSheetCopyPath, 'a') as f:
                    f.write('\t'.join(i))
                    f.write('\n')
                print(i)
        print("Done")
        sleep(70)
