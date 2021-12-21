## Requirement: install openpyxl module for this script to work
## pip install openpyxl

import openpyxl as xl
import os

# Test Location
folder_loc="C:\\Project_Files\\gladd4-verif\\VerificationEnvironment\\Tools\\Rapita\\Tests"
files=os.listdir(folder_loc)

# Placeholder for data
jama_id=[]
tp_id=[]
tc_num=[]

# Fetching traceability data from TPs
for file in files[1:]:
    filepath=os.path.join(folder_loc,file)
    print("Processing file",file,"...")
    wb=xl.load_workbook(filepath)
    sheet=wb.active
    tc=1
    for row in sheet.rows:
        for cell in row:
            if(cell.value=='Requirements'):
                for test in row:
                    if(test.value!=None and test.value!='Requirements'):
                        jama_id.append(test.value)
            if(cell.value=='Name'):
                for test in row:
                    if(test.value!=None and test.value!='Name'):
                        tp_id.append(test.value)
                        tc_num.append(tc)
                        tc+=1
                        
    print("Complete")
print('Done')

# Printing fetched matrix
print(f"Number of JAMA ID populated = {len(jama_id)}")
print(f"Number of TP ID populated = {len(tp_id)}")
print(f"Number of Test Case # populated = {len(tc_num)}")

# Creating an excel file for extracted data
wwb=xl.Workbook()
ws=wwb.active
ws.title='Test Trace'
ws.append(['Test Case Jama ID','Test Case #','Test procedure Name'])

# Appending data
for jid,tcn,tpi in zip(jama_id,tc_num,tp_id):
    ws.append([jid,tcn,tpi])

# Give any filename to save data into
wwb.save('traceability.xlsx')

