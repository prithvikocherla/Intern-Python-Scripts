#!/usr/bin/env python
# coding: utf-8

# Author - Prithvi Kocherla 
# Date - 4/24/2019
# Copyright (c) Bristol-Myers Squibb 2019

import os, json
import win32com.client
import re
from datetime import datetime
import pandas as pd
import zipfile
import numpy as np

# Environment selection, set it to 'PROD' for production or '' for testing
env = ''

def main():
    t_drive_path = r'T:\Data Sources\Automation - Do not touch\\' if env is 'PROD' else r''
    get_path = t_drive_path if env is 'PROD' else os.getcwd()

    print(f"Working Dictionary: {get_path}")
    print(f"Pandas version: {pd.__version__}")

    # Paths for final excel repots
    field_agent_path = r'Downloaded Attachments\BMS FIELD Agent Daily Stats\Field-Scientific Global Agents 2018.xlsx'
    field_calltype_path = r'Downloaded Attachments\BMS FIELD Call Type Daily Stats\Field-Scientific Global Call Type.xlsx' 
    facility_agent_path = r'Downloaded Attachments\Facilities Agent Daily Stats\Facilities Agent Daily Stats 2018.xlsx'
    facility_calltype_path = r'Downloaded Attachments\Facilities Call Type Daily\Facilities Call Type Daily Stats 2018.xlsx'
    otc_agent_path = r'Downloaded Attachments\OTC AMER-CANADA Agent Daily Stats\AGENT DAILY STATS 2018 2019_v2.xlsx'
    otc_calltype_path = r'Downloaded Attachments\OTC AMER-CANADA Call Type Daily Stats\CALL TYPE DAILY 2018_v2.xlsx'
    emea_agent_path = r'Downloaded Attachments\OTC EMEA Agent Daily\OTC EMEA AGENTS NEW.xlsx'
    emea_calltype_path = r'Downloaded Attachments\OTC EMEA Call Type Daily\OTC EMEA CALL TYPE NEW.xlsx'

    # Dictionary containing File Names and File paths
    filenames = {field_agent_path.split("\\")[1] : t_drive_path + field_agent_path,
    field_calltype_path.split("\\")[1] : t_drive_path + field_calltype_path,
    facility_agent_path.split("\\")[1] :  t_drive_path + facility_agent_path,
    facility_calltype_path.split("\\")[1] : t_drive_path + facility_calltype_path,
    otc_agent_path.split("\\")[1] : t_drive_path + otc_agent_path,
    otc_calltype_path.split("\\")[1] : t_drive_path + otc_calltype_path,
    emea_agent_path.split("\\")[1] : t_drive_path + emea_agent_path,
    emea_calltype_path.split("\\")[1] : t_drive_path + emea_calltype_path}

    print(filenames, sep=',', end='\n\n')
    # Connection to Outlook mail
    print("Connecting to Outlook")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        mailbox = outlook.Folders
    except Exception as e:
        print(e)
        raise("Outlook connection failed")

    # check to see what this mailbox object is looking for
    icount = 0
    for i in mailbox:
        icount += 1
        if re.match(r"[\w\.-]+@[\w\.-]+",str(i)):
            print("found BMS inbox (" + str(i) + ") processing...")
            bms_box = icount - 1
            break

    # Connecting to scheduled reports folder
    BMS_inbox = outlook.Folders[bms_box]
    main_inbox = BMS_inbox.Folders["Inbox"]
    try:
        reporting_inbox = main_inbox.Folders["Scheduled Reports"]
    except Exception:
        raise Exception("'Scheduled Reports' folder not found under 'inbox', please set a rule and create this folder in outlook")

    # Selecting the type of email to download the report
    print("\nPlease input the respective option for downloading the report from Email.", end="\n\n")
    
    # Dynamically creating menu
    master_dict = {}
    for key, filename in enumerate(filenames):
        print(f"{key+1} - {filename}")
        master_dict[key+1] = filename
        lastkey = key + 2

    # Storing all items into dictionary to use up when selected option is of type 'All'
    master_dict[lastkey] = filenames
    print(f"{lastkey} - All")

    while True:
        var = input('Enter Option: ')
        try:
            var = int(var)

            # Selection of file type required to be downloaded
            if master_dict[var]:
                if var == lastkey:
                    fileList = master_dict[lastkey]
                    break
                fileList = {master_dict[var] : filenames[master_dict[var]]}
                break
        except Exception as e:
            print(e)
            print(f"Incorrect choice, please choose numbers between {list(master_dict.keys())[0]} and {list(master_dict.keys())[-1]}")        
            
    for fileName, mainfile_Path in fileList.items():
        fileName = fileName.split(".")[0]
        print("Loading files from: " + fileName)
        # Loading master excel file
        print("\nLoading master excel file...")
        exceldb_master = pd.ExcelFile(mainfile_Path)

        # Loading previous write history date to limit the duplicates
        print("Loading file write history...")
        try:
            exceldb_history = pd.read_excel(exceldb_master, sheet_name='write_history')
            sheets = exceldb_master.sheet_names
            sheets.remove("write_history")
        except Exception as e:
            print("No write history found, creating new write history")
            sheets = exceldb_master.sheet_names
            exceldb_history = pd.DataFrame({"write_history": [""]})

        print("List of sheets:" + str(sheets))
        print("Parsing latest sheet..")
        lastsheet = sheets[-1]
        exceldb = exceldb_master.parse(lastsheet)

        # note: this logic will not work for 10+ sheets - will add additional logic for this later
        if exceldb.shape[0] > 900_000:
            print("Too many records in last sheet, creating new sheet...")
            exceldb = pd.DataFrame().reindex_like(exceldb).dropna()
            for i in lastsheet:
                if i.isdigit():
                    digi = int(i) + 1
                    lastsheet = 'Sheet' + str(digi)
        else:
            sheets.remove(lastsheet)

        # Attachment download, extract, concatenate, and write to final report method
        count = 0
        for i in reporting_inbox.Items:
            try:
                date = i.Subject.split()[-1]
                email_date = date.split('[')[1].split(']')[0]
                email_date = datetime.strptime(email_date, '%m/%d/%y')
            except Exception as e:
                print("Incorrect date format, skipping: " + i.Subject)
                continue

            # Checking the dates that are not in write history and email with required subject line
            if (np.datetime64(email_date) not in exceldb_history['write_history'].values) and fileName in str(i.Subject):
                print("\nEmail Received Date: {}".format(email_date.date()))
                subject = str(i.Subject)
                print('Subject: ' + subject)

                attachments = i.Attachments
                x = 1
                while x <= attachments.Count:
                    # extracting excel files from zip and moving to temp dictionary
                    attachment = attachments.Item(x)
                    x += 1

                    attachment.SaveASFile(os.path.join(get_path,attachment.FileName))
                    print("Attachment Name: " + str(attachment))

                    new_path = os.path.join(get_path, attachment.FileName)

                    zip_ref = zipfile.ZipFile(new_path, 'r')
                    zip_ref.extractall(get_path)
                    zip_ref.close()
                    os.remove(new_path)
                    print("File extracted, converting to .xlsx format...")#, end='\n\n')

                    file_path = os.path.join(get_path, (fileName + ".xls"))
                    # reading xls files into pandas
                    df = pd.read_html(file_path, header=0)  
                    df = df[0]
                    list1 = df.columns
                    list2 = df.loc[0].values
                    list1 = [x for x in list1 if x.split(" ")[0] != 'Unnamed:']
                    list1 = [x for x in list1 if x != 'Completed Tasks']
                    try:
                        # fixing column values
                        if var in [1,3,5,7]:
                            list1a = list1[:3]
                            list1b = list1[3:]
                            list2 = [x for x in list2 if str(x) != "nan"]

                        elif var in [2,4,6,8]:
                            list1 = [x for x in list1 if x != 'Tasks']
                            list1a = list1[:5]
                            list1b = list1[5:]
                            list2 = [x for x in list2 if str(x) != "nan"]
                    except:
                        raise("file selection not found, how did you get this far?")


                    final_list = list1a + list2 + list1b
                    df = df.iloc[1:]
                    df = df[pd.notnull(df['DateTime'])]
                    try:
                        '''Prithvi please review why this error is occuring'''
                        df.columns = final_list
                    except Exception:
                        '''if frame is empy, still update write history then move it to archive'''
                        print("Failed to transform file, please review manually")
                        print(str(attachment) + "    " + subject, file=open("Review-Log.txt", mode='a'))
                        continue
                    dfr = df.copy()

                    out_path = os.path.join(get_path, 'Downloaded Attachments\\' + fileName + '\Archive\\')
                    out_path = os.path.join(out_path, fileName + " " + str(email_date.date()) + ".xlsx")
                    dfr.to_excel(out_path, index = False)
                    os.remove(file_path)
                    print("File is Archived to location: {}".format(out_path))

                    history_to_write = pd.DataFrame({"write_history": [email_date]})

                    file_to_write = df.copy()
                    # concat the excel db with the file data
                    frames = [exceldb, file_to_write]
                    exceldb = pd.concat(frames, sort=False).reset_index(drop=True)
                    #print(file_result)

                    # concat write history with file date
                    history_frames = [exceldb_history, history_to_write]
                    exceldb_history = pd.concat(history_frames, sort=True)
                    exceldb_history['write_history'] = pd.to_datetime(exceldb_history['write_history'])
                    exceldb_history.sort_values('write_history',inplace=True)

                    count += 1
                
                #  Check if appending new data exceeds the max row length in excel sheet
                if exceldb.shape[0] > 1_000_000:
                    raise Exception("Rows exceeding 1 million during operation, aborting."
                                    "Please create a new sheet to add this many rows.")      

        # Check if there is any new appended data to write to excel file
        if count != 0:
            print("\nWriting to file...")
            try:
                writer = pd.ExcelWriter(exceldb_master)
                for sheet in sheets:
                    exceldb_master.parse(sheet).to_excel(writer, sheet_name=sheet, index=False)
                exceldb.to_excel(writer, sheet_name=lastsheet, index=False)
                exceldb_history.to_excel(writer, sheet_name='write_history', index=False)
                writer.save()
                print("\nSuccess!! All the available data is downloaded to main excel report file!")
            except Exception as e:
                print("Operation failed with error: " + str(e))  
        else:
            print("\nData was already Downloaded to main excel report file, aborting process!!!")

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(e)
        input("Press enter to exit")
        exit()
