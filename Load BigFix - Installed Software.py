# -*- coding: utf-8 -*-
"""
Created on Fri Aug  9 13:40:23 2019

@author: Jason Richmond
"""

from urllib.request import Request, urlopen
import json
import urllib.request
import re
import string
import time
from schema import Schema, And, Use, Optional, SchemaError
import os
import pyodbc
import win32com.client as win32
import sys
import pyodbc
import win32com.client as win32
import sys
from tabulate import tabulate
import csv
from itertools import islice


#######################################################
#######################################################
###############   Write URI results ###################
###############   to DB             ###################
#######################################################
#def write_to_db(data, split_data):
def write_installed_software_to_db():
    
    conn = pyodbc.connect('Driver={SQL Server};'
                          'Server=xxxxxxx;'
                          'Database=xxxxxxx;'
                          'Trusted_Connection=yes;')
    cursor = conn.cursor()

    try:
        
        # Kick of SQL Server processes to populate reporting tables
        #cursor.execute("EXEC [dbo].[proc_xxxxxxx];")
        #cursor.commit()
        
        yourcsv = r'Z:FILEPATH'
        #output_file = open("sqlfile.txt","w") 
        table = 'xxxxxxx'
        row_cnt = 0
        db_vals = ''
        
        
        cursor.execute("DELETE FROM xxxxxxx;")
        cursor.commit()
        #INSERT SOURCE RECORDS TO DESTINATION
        with open(yourcsv) as csvfile:
            csvFile = csv.reader(csvfile, delimiter=',')
            headers = ['ComputerName','UserName', 'InstalledApplications','SplitFlag','ComputerSerialNumber','ComputerModel','OS','IP_Address','LastReportTime']
            insert = 'INSERT INTO {} ('.format(table) + ', '.join(headers) + ') VALUES '
            next(csvfile)
            
            for row in csvFile:
            #for row in islice(csv.reader(csvFile), 500, None):
                #print(row)
                row_cnt += 1 
                ComputerName = row[0]
                UserName = row[1]
                InstalledApplications = row[2]
                ComputerSerialNumber = row[3]
                ComputerModel = row[4]
                OS = row[5]
                IP_Address = row[6]
                LastReportTime = row[7]
                
                
                IP_Address = IP_Address.rstrip('\n').replace('\r', '').replace('\n', ' ').replace('\'', '`')
                InstalledApplications = re.sub(r'[^{0}\n]'.format(string.printable), '', InstalledApplications)
                InstalledApplications = re.sub(r'[^\x00-\x7f]',r'', InstalledApplications) 
                InstalledApplications = InstalledApplications.replace('\'', '`')
                
                col_len = len(InstalledApplications)
                
                InstalledApplications = InstalledApplications.replace("\n", "~")
                db_vals = '(\'' + xxxxxxx + '\', \'' + xxxxxxx + '\', \'' + xxxxxxx + '\', \'N\', \'' + xxxxxxx + '\', \'' + xxxxxxx + '\', \'' + xxxxxxx + '\', \'' + xxxxxxx + '\', \'' + xxxxxxx + '\');' 
                cursor.execute(insert + db_vals)
                cursor.commit()
                
            
        print(row_cnt)

    
    except:
        print(row_cnt)
        print(db_vals)
        print("Unexpected error:", sys.exc_info()[0])
        pass
        


def write_approve_software_to_db():
    
    conn = pyodbc.connect('Driver={SQL Server};'
                          'Server=xxxxxxx;'
                          'Database=xxxxxxx;'
                          'Trusted_Connection=yes;')
    cursor = conn.cursor()

    try:
        
        # Kick of SQL Server processes to populate reporting tables
        #cursor.execute("EXEC [dbo].[proc_xxxxxxx];")
        #cursor.commit()
        
        yourcsv = r'Z:FILEPATH'
        #output_file = open("sqlfile.txt","w") 
        table = 'xxxxxxx'
        row_cnt = 0
        
        
        
        cursor.execute("DELETE FROM xxxxxxx;")
        cursor.commit()
        with open(yourcsv) as csvfile:
            csvFile = csv.reader(csvfile, delimiter=',')
            #headers = next(csvFile)
            headers = ['ProductVersion','ApplicationName','Category','PublisherName','ID','Restricted','RestrictedUse','SoftwareRestriction','ApprovalType','Inactive','ItemType','Path']
            insert = 'INSERT INTO {} ('.format(table) + ', '.join(headers) + ') VALUES '
            next(csvfile) # skip first row of CSV file
            for row in csvFile: 
                row_cnt += 1 
                ProductVersion = row[0].rstrip(' ').lstrip(' ')
                ApplicationName = row[1].rstrip(' ').lstrip(' ')
                Category = row[2].rstrip(' ').lstrip(' ')
                PublisherName = row[3].rstrip(' ').lstrip(' ')
                ID = row[4].rstrip(' ').lstrip(' ')
                Restricted = row[5].rstrip(' ').lstrip(' ')
                RestrictedUse = row[6].rstrip(' ').lstrip(' ')
                SoftwareRestriction = ''#row[7].rstrip(' ').lstrip(' ')
                ApprovalType = row[8].rstrip(' ').lstrip(' ')
                Inactive = row[9].rstrip(' ').lstrip(' ')
                ItemType = row[10].rstrip(' ').lstrip(' ')
                Path = row[11].rstrip(' ').lstrip(' ')
                
                ApplicationName = re.sub(r'[^{0}\n]'.format(string.printable), '', ApplicationName)
                ApplicationName = re.sub(r'[^\x00-\x7f]',r'', ApplicationName) 
                ApplicationName = ApplicationName.replace('\'', '`')
                
                PublisherName = re.sub(r'[^{0}\n]'.format(string.printable), '', PublisherName)
                PublisherName = re.sub(r'[^\x00-\x7f]',r'', PublisherName) 
                PublisherName = PublisherName.replace('\'', '`')
                #ApplicationName = set(ApplicationName.printable)
                
                col_len = len(SoftwareRestriction)
                #print(col_len)
                
                if col_len < 4000:
                    db_vals = '(\'' + ProductVersion + '\', \'' + ApplicationName + '\', \'' + Category + '\', \'' + PublisherName + '\', \'' + ID + '\', \'' + Restricted + '\', \'' + RestrictedUse + '\', \'' + SoftwareRestriction + '\', \'' + ApprovalType + '\', \'' + Inactive + '\', \'' + ItemType + '\', \'' + Path + '\');' 
                    #print(insert + db_vals)
                    cursor.execute(insert + db_vals)
                    cursor.commit()
                
                elif col_len > 0:
                    SoftwareRestriction = SoftwareRestriction[0:4999]
                    db_vals = '(\'' + ProductVersion + '\', \'' + ApplicationName + '\', \'' + Category + '\', \'' + PublisherName + '\', \'' + ID + '\', \'' + Restricted + '\', \'' + RestrictedUse + '\', \'' + SoftwareRestriction + '\', \'' + ApprovalType + '\', \'' + Inactive + '\', \'' + ItemType + '\', \'' + Path + '\');' 
                    #print(insert + db_vals)
                    cursor.execute(insert + db_vals)
                    cursor.commit()
                    
                else:
                    quit
                
        print(row_cnt)
     
        #emailTable = tabulate(rows, headers=['Load Date', 'Process', 'Start', 'End', 'Status', 'Loaded #'], tablefmt='simple')
        #print(emailTable)

    
    except:
        print(row_cnt)
        print(SoftwareRestriction)
        print(col_len)
        print("Unexpected error:", sys.exc_info()[0])
        pass


#######################################################
#######################################################
##################  SEND EMAIL ########################
#######################################################
#######################################################
def send_email(mail_to, mail_subject, mail_body):     
    
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = mail_to #'jason.richmond@bkfs.com'
    mail.Subject = mail_subject#'Cofense info'
    mail.Body = mail_body# 'Missing fields: ' + str(list_of_missing_fields) + '\nNew Fields: ' + str(list_of_new_fields)
    #mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional
    # To attach a file to the email (optional):
    #attachment  = "Path to the attachment"
    #mail.Attachments.Add(attachment)
    mail.Send()



write_installed_software_to_db()
write_approve_software_to_db()
