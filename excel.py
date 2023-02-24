from pprint import pprint
import os
import json
#import win32com.client as win32
from openpyxl import Workbook  
import time  

def create_excel(path, name):
    wb = Workbook()  
    sheet = wb.active
    sheet.title = 'Speaker Diarisation'

    #step 1.1- Read the json file
    json_data=json.loads(open(path).read())

    #insert headers
    header_labels=('speaker', 'text' , 'emotion', 'time (ms)', 'dbfs')

    for index, val in enumerate(header_labels):
        sheet.cell(row=1, column=index+1).value = val
    
    speakers=[]
    row_num = 2
    for i, data in enumerate(json_data):
        speakers.append(data['speaker'])
        for j, dbfs in enumerate(data['dbfs']):
            sheet.cell(row=row_num, column=1).value = data['speaker']
            sheet.cell(row=row_num, column=2).value = data['text']
            sheet.cell(row=row_num, column=3).value = data['emotion']
            sheet.cell(row=row_num, column=4).value = dbfs['time']
            sheet.cell(row=row_num, column=5).value = dbfs['value']
        
            row_num=row_num+1


    # Create speaker talktime distribution worksheet
   

    unique_speakers = unique(speakers)
   
    st=[]
    for speaker in unique_speakers:
        
        speakingTime = 0

        for data in json_data:
            if data['speaker'] == speaker:
                speakingTime = speakingTime + (data['end'] - data['start'])
                
        st.append(speakingTime)
        totalTime = sum(st)

        speakingPercentage=[]
        for t in st:
            speakingPercentage.append(round(t/totalTime*100))

    print(st)
    header_label1=('speaker', 'time (%)')

    wb.create_sheet('Speaker Talktime Distribution') # Create a new sheet
    wb.active = wb['Speaker Talktime Distribution'] # Set it as active
    sheet1 = wb.active

    for index, val in enumerate(header_label1):
        sheet1.cell(row=1, column=index+1).value = val

    for row, speakerTime in enumerate(speakingPercentage):
        sheet1.cell(row+2, column=1).value = unique_speakers[row]
        sheet1.cell(row+2, column=2).value = speakerTime

    # Set the first sheets as active.
    wb.active = wb['Speaker Diarisation']

    wb.save(name + ".xlsx")  

    return True


def unique(list):
 
    # initialize a null list
    unique_list = []
 
    # traverse for all elements
    for x in list:
        # check if exists in unique_list or not
        if x not in unique_list:
            unique_list.append(x)
   
    return unique_list