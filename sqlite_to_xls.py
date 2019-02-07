#! /usr/bin/env python
# -*- coding: utf-8 -*-

import codecs
import shutil
import sqlite3
import os,os.path
import re
import sys
from xlsxwriter.workbook import Workbook


def checkparam():
    if (len(sys.argv) != 1):
        print "\n\tRequires three arguments:"        
        print "\n\t\export_excel.py {databasename} {tablename} {output excel name}\n\n"
        sys.exit()

# Passing param 1 as SQlite DB
db = sys.argv[1]


#Extracting db filename
outputbook = db.split(".")[0] + ".xlsx"

#connect to the database

conn = sqlite3.connect(db)
cur = conn.cursor()

workbook = Workbook(outputbook)

#func to extract sqlite table in a worksheet

def ext_dbtbl(i,sheetname): 

    i = i + 1

    worksheet = workbook.add_worksheet(sheetname)

    #Exract Table Headers 1st    
    headers = cur.execute("""SELECT sql FROM sqlite_master WHERE tbl_name = '""" + str(sheetname) + """' AND type = 'table';""")    
    for hk, hrow in enumerate(headers):       
        for hj, value in enumerate(hrow):
            headerstr = hrow[hj].strip('CREATE TABLE "').strip(str(sheetname)).strip('" (').strip(')')
            col = 0 
            for headr in headerstr.split(','):
                worksheet.write(0, col, headr)
                col += 1

    #Exract Table content then
    mysel=cur.execute("select * from `" + str(sheetname) + "`")
    for k, row in enumerate(mysel):       
        for j, value in enumerate(row):
            worksheet.write(k + 1, j, row[j])

#function to query database
def queryfunc():  
    #Query to count the number of table in the db
    querystr = "SELECT tbl_name FROM sqlite_master WHERE type='table';"  
    result = cur.execute(querystr)
    rows = result.fetchall()
    for i,a in enumerate(rows):
        sheetname = str(a).strip("(").strip("u'").strip("')',")
        ext_dbtbl(i,sheetname)

# Call the above function
queryfunc()

workbook.close()
os.system('chmod -R 777 ' + outputbook)





















