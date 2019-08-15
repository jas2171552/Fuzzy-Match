# -*- coding: utf-8 -*-
"""
Created on Tue Aug 13 14:22:52 2019

@author: e1208569
"""
# pip install fuzzywuzzy 
#pip install python-Levenshtein

import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
from fuzzywuzzy.fuzz import partial_ratio
import pyodbc

#df = pd.read_csv('room_type.csv')
#df.head(10)


def match_name(applicationCombined, approveSoftware, min_score=0):
    # -1 score incase we don't get any matches
    max_score = -1
    # Returning empty name for no match as well
    max_name = ""
    approvedKey = ""
    # Iternating over all names in the other
    for name2 in approveSoftware:     
        score  = fuzz.token_set_ratio(applicationName, name2)
        score2 = fuzz.WRatio(applicationName, name2)
        score3 = fuzz.ratio(applicationName, name2)
        score4 = fuzz.partial_ratio(applicationName, name2)
        
        # Checking if we are above our threshold and have a better score
        if (score > min_score) & (score > max_score):
            max_name = name2
            #max_score = score
            max_score = max(score, score2, score3, score4)
            
    return (max_name, max_score, approvedKey)



conn = pyodbc.connect('Driver={SQL Server};'
                          'Server=JTCTLSWICSDB03,28001;'
                          'Database=ERMReporting;'
                          'Trusted_Connection=yes;')
cursor = conn.cursor()

approvedCursor = conn.cursor()
approvedCursor.execute('''
       SELECT CAST(ApprovedSoftwareKey AS VARCHAR(10)) ApprovedSoftwareKey, ApplicationName, ApplicationVersion FROM reporting.BigFix_ApprovedSoftware ORDER BY ApplicationName, ApplicationVersion;
       ''')
approveSoftware = approvedCursor.fetchall()
#print(approveSoftware)

cursor.execute('''
       SELECT InstalledSoftwareKey, ApplicationName, ApplicationVersion FROM reporting.BigFix_InstalledSoftware;
       ''')
installedSoftware = cursor.fetchall()
#print(installedSoftware) 

for row in installedSoftware:
    #print(row)
    
    installedApplicationKey = str(row[0])
    applicationName = row[1]
    applicationVersion = row[2]
    applicationCombined = applicationName + ' ' + applicationVersion
    
    matchPct = match_name(applicationCombined, approveSoftware, min_score=0)    
    if len(matchPct[0]) > 0:
        aprvdAppKey = matchPct[0][0]
        aprvdApp = matchPct[0][1]
        aprvdAppVer = matchPct[0][2]
        matchPct = str(matchPct[1])
        
        key = conn.cursor()
        sql = 'UPDATE reporting.BigFix_InstalledSoftware SET ApprovedSoftwareKey = \'' +  aprvdAppKey + ' \', MatchPCT = \'' + matchPct + '\' WHERE InstalledSoftwareKey = \'' + installedApplicationKey + '\';'
        key.execute(sql)
        key.commit()
        
        

conn.close()  

