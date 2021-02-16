# Open the msaccess database with the real estate iq files linked
# https://www.geeksforgeeks.org/reading-excel-file-using-python/
#
#
#
#
#
#
#
#



#from __future__ import print_function

import os
import sys
import pyodbc
import pandas as pd
import numpy as np
import xlsxwriter
import datetime
#import time
#import argparse
#import sys
from urllib.request import urlopen
import subprocess
import webbrowser
from time import sleep

def isNotBusiness2(data,i,oKeys):
    temp = ''

    # this field is not in all files but is needed in prefor
    try:
      if data['Ownership'].iloc[i] == "OWNER OCCUPIED":
          return True
    except:
      pass

    for k in oKeys:
      try:
          temp = temp + str( data[k].iloc[i] )
      except:
          pass

    temp = temp.lower()
    #print(temp)

    ids = ['lp','llc','llp','inc','corp','ltd','trust']

    for s in ids:
        if temp.find(s) != -1:
            #print('found ', s, ' in ',temp)
            return False

    return True

def isNotBusiness(data,i):
    temp = ''

    #if data['Ownership'].iloc[i] == "OWNER OCCUPIED":
    #    return True

    try:
        temp = temp + str( data['Mortgagor First Name'].iloc[i] )
    except:
        pass

    try:
        temp = temp + " " + str( data['Mortgagor Last Name'].iloc[i]  )
    except:
        pass

    try:
        temp = temp + " " + str( data['Owner First Name'].iloc[i]  )
    except:
        pass

    try:
        temp = temp + " " + str( data['Owner Last Name'].iloc[i]  )
    except:
        pass

    temp = temp.lower()
    #print(temp)

    ids = ['lp','llc','llp','inc','corp','ltd','trust']

    for s in ids:
        if temp.find(s) != -1:
            #print('found ', s, ' in ',temp)
            return False

    return True

def getNames(data, i):
  names = {}
  ofn = []
  oln = []

  try:
    mfn = data['Mortgagor First Name'].iloc[i].split("&")
    mln = data['Mortgagor Last Name'].iloc[i].split("&")

    if len(mln) == 1 and len(mfn) > 1:
        mln.append(mln[0])

    j = 0
    while j < len(mfn):
        k = 'mfn' +str(j)
        names[k] = mfn[j].strip().lower()
        j += 1

    j = 0
    while j < len(mln):
        k = 'mln' +str(j)
        names[k] = mln[j].strip().lower()
        j += 1

  except:
    #print('\tNo mortgage name ') #, data['Mortgagor First Name'].iloc[i], data['Mortgagor Last Name'].iloc[i] )
    pass


  if data['Owner First Name'].iloc[i] != "N/A":
    try:
        ofn = data['Owner First Name'].iloc[i].split("&")
        oln = data['Owner Last Name'].iloc[i].split("&")
        #print("ofn:", ofn, "oln:",oln)

        # handle case of two first names but only one last name
        if len(oln) == 1 and len(ofn) > 1:
            mln.append(oln[0])

        # copy first and last to names dictionary
        j = 0
        while j < len(ofn):
            k = 'ofn' +str(j)
            names[k] = ofn[j].strip().lower()
            j += 1

        j = 0
        while j < len(oln):
            k = 'oln' +str(j)
            names[k] = oln[j].strip().lower()
            j += 1

    except:
        pass


  if data['Relative Full Name'].iloc[i] != "N/A":
    try:
        rn = str(data['Relative Full Name'].iloc[i]).split(" ")
        rfn = rn[0].lower()
        suffixs = ['sr','jr','ii','iii']

        for s in suffixs:
            if str(rn[len(rn)-1]).lower() == s:
                #print("\tsuffix ", s," found in ",rn, rn[len(rn)-2],rn[len(rn)-1])
                rln = rn[len(rn)-2] + " " + rn[len(rn)-1]
                rln = rln.lower()
                break
            else:
                rln = str(rn[len(rn)-1]).lower()
    except:
        pass



  try:
    names['rfn'] = rfn
  except:
    pass

  try:
    names['rln'] = rln
  except:
    pass



  #print(names)

  return names

def splitFullName(n):
  if n != "N/A":
    suffixs = ['sr','jr','ii','iii']
    splitNames = {}

    try:
      x = n.split(" ")
      splitNames['fn'] = x[0].lower()

      for s in suffixs:
        if str(x[len(x)-1]).lower() == s:
            #print("\tsuffix ", s," found in ",x, x[len(x)-2],x[len(x)-1])
            ln = x[len(x)-2] + " " + x[len(x)-1]
            splitNames['ln'] = ln.lower()
            break
        else:
            splitNames['ln'] = str(x[len(x)-1]).lower()
    except:
      pass


  return splitNames

def pickPhone(data,index,num,cols):
  # fill an array with data identified
  best = []
  for c in cols:
    # using try to discard missing data
    try:
      best.append('1' + str(int(data[c].iloc[index])))
    except:
      pass

  # result must be num in length - pad the list with empty strings
  while len(best) <= num:
    best.append('')
  return best

def pickEmail(data,index,num,cols):
  # fill an array with data identified
  best = []
  for c in cols:
    try:
      s = str(data[c].iloc[index]).strip()
      if s != "N/A": 
        best.append( s )
    except:
      pass

  # result must be num in length - pad the list with empty strings
  while len(best) <= num:
    best.append(np.NaN)

  return best

def makeMetaData(data,index):
  s = (data[data.columns].iloc[index].astype(str))
  #print(type(s),s)
  return(s)

def difFiles(earlyFile,lateFile):
    print(earlyFile,lateFile)

    df1 = pd.read_excel(earlyFile)
    df2 = pd.read_excel(lateFile)

    print(df2)
    #comparison_values = df1.values == df2.values
    #print (comparison_values)
    
    #difference = df1[df1!=df2]
    #print( difference)

def fbPreForList(data,fileName):
  phoneColumns = ['CellPhones1', 'CellPhones2', 'CellPhones3','CellPhones4','CellPhones5',
                  'PhoneNumbers1','PhoneNumbers2','PhoneNumbers3','PhoneNumbers4','PhoneNumbers5'  ]
  emailColumns = ['EmailAddresses1','EmailAddresses2','EmailAddresses3','EmailAddresses4','EmailAddresses5' ]
  relPhoneColumns = ['Relative Phone1','Relative Phone2','Relative Phone3']
  relEmailColumns = ['Relative EmailAddresses1','Relative EmailAddresses2','Relative EmailAddresses3'  ]
  ownerKeys = ['Mortgagor First Name','Mortgagor Last Name','Owner First Name','Owner Last Name']


  row = {}
  row["fn"] = ''
  row["ln"] = ''
  row['ct'] = ''
  row['st'] = ''
  row['zip'] = ''
  row['phone1'] = ''
  row['phone2'] = ''
  row['phone3'] = ''
  row['email1'] = ''
  row['email2'] = ''
  row['type'] = ''
  row['meta'] = ''

  workbook = xlsxwriter.Workbook(fileName)
  worksheet = workbook.add_worksheet()
  print("Starting on ", fileName, "length of list", len(data))

  for j, (k, v) in enumerate(row.items(), start=0):
    worksheet.write(0, j, k)

  i = 0
  r = 1
  for i in range(0, len(data)):
    if isNotBusiness2(data,i,ownerKeys):
      mNames = getNames(data, i)
      #print("Source Row", i, mNames)
      if len(mNames) == 0:
        print('\tNo Names in row ', i)
      else:
        phones = pickPhone(data,i,3,phoneColumns)
        emails = pickEmail(data,i,3,emailColumns)
        #print(emails)
        #meta = makeMetaData(data,i)
        meta = "Not Yet"

        rphones = pickPhone(data,i,3,relPhoneColumns)
        remails =  pickEmail(data,i,3,relEmailColumns)

        try:
            row = {}
            row["fn"] = mNames['mfn0']
            row["ln"] = mNames['mln0']
            row['ct'] = data['Property City'].iloc[i]
            row['st'] = 'Texas'
            row['zip'] = data['Property Zip Code'].iloc[i]
            row['phone1'] = phones[0]
            row['phone2'] = phones[1]
            row['phone3'] = phones[2]
            row['email1'] = emails[0]
            row['email2'] = emails[1]
            row['type'] = 'mortgage first'
            row['meta'] = meta

            for j, (k, v) in enumerate(row.items(), start=0):
                if v == 'nan':
                    v = ''
                worksheet.write(r, j, v)

            #print('Added mortgage 1 to excel')
            r += 1
        except:
            pass

        try:
            row = {}
            row["fn"] = mNames['mfn1']
            row["ln"] = mNames['mln1']
            row['ct'] = data['Property City'].iloc[i]
            row['st'] = 'Texas'
            row['zip'] = data['Property Zip Code'].iloc[i]
            row['phone1'] = phones[0]
            row['phone2'] = phones[1]
            row['phone3'] = phones[2]
            row['email1'] = emails[0]
            row['email2'] = emails[1]
            row['type'] = 'mortgage second'
            row['meta'] = meta

            for j, (k, v) in enumerate(row.items(), start=0):
                if v == 'nan':
                    v = ''
                worksheet.write(r, j, v)

            #print('Added mortgage 2 to excel')
            r += 1
        except:
            pass

        try:
            row = {}
            row["fn"] = mNames['ofn0']
            row["ln"] = mNames['oln0']
            row['ct'] = data['Owner City'].iloc[i]
            row['st'] = data['Owner State'].iloc[i]
            row['zip'] = data['Owner Zip Code'].iloc[i]
            row['phone1'] = phones[0]
            row['phone2'] = phones[1]
            row['phone3'] = phones[2]
            row['email1'] = emails[0]
            row['email2'] = emails[1]
            row['type'] = 'owner first'
            row['meta'] = meta

            for j, (k, v) in enumerate(row.items(), start=0):
                if v == 'nan':
                    v = ''
                worksheet.write(r, j, v)

            #print('Added owner to excel')
            r += 1
        except:
            pass

        try:
            row = {}
            row["fn"] = mNames['ofn1']
            row["ln"] = mNames['oln1']
            row['ct'] = data['Owner City'].iloc[i]
            row['st'] = data['Owner State'].iloc[i]
            row['zip'] = data['Owner Zip Code'].iloc[i]
            row['phone1'] = phones[0]
            row['phone2'] = phones[1]
            row['phone3'] = phones[2]
            row['email1'] = emails[0]
            row['email2'] = emails[1]
            row['type'] = 'owner second'
            row['meta'] = meta

            for j, (k, v) in enumerate(row.items(), start=0):
                if v == 'nan':
                    v = ''
                worksheet.write(r, j, v)

            #print('Added owner to excel')
            r += 1
        except:
            pass

        try:
            row = {}
            row["fn"] = mNames['rfn']
            row["ln"] = mNames['rln']
            row['ct'] = data['Relative City'].iloc[i]
            row['st'] = data['Relative State'].iloc[i]
            row['zip'] = data['Relative Zip'].iloc[i]
            row['phone1'] = rphones[0]
            row['phone2'] = rphones[1]
            row['phone3'] = rphones[2]
            row['email1'] = remails[0]
            row['email2'] = remails[1]
            row['type'] = 'relative'
            row['meta'] = meta

            for j, (k, v) in enumerate(row.items(), start=0):
                if v == 'nan':
                    v = ''
                worksheet.write(r, j, v)

            #print('Added relative to excel')
            r += 1
        except:
            pass


  workbook.close()

def fbProbateList(data,fileName):
    phoneColumns = ['CellPhones1', 'CellPhones2', 'CellPhones3','CellPhones4','CellPhones5',
                    'PhoneNumbers1','PhoneNumbers2','PhoneNumbers3','PhoneNumbers4','PhoneNumbers5'  ]
    emailColumns = ['EmailAddresses1','EmailAddresses2','EmailAddresses3','EmailAddresses4','EmailAddresses5' ]
    relPhoneColumns = ['Relative Phone1','Relative Phone2','Relative Phone3']
    relEmailColumns = ['Relative EmailAddresses1','Relative EmailAddresses2','Relative EmailAddresses3'  ]
    #ownerKeys = ['Grantee First Name','Grantee Last Name'] #,'Owner First Name','Owner Last Name']

    row = {}
    row["fn"] = ''
    row["ln"] = ''
    row['ct'] = ''
    row['st'] = ''
    row['zip'] = ''
    row['phone1'] = ''
    row['phone2'] = ''
    row['phone3'] = ''
    row['email1'] = ''
    row['email2'] = ''
    row['type'] = ''
    row['meta'] = ''

    workbook = xlsxwriter.Workbook(fileName)
    worksheet = workbook.add_worksheet()
    print("Starting on ", fileName, "length of list", len(data))

    for j, (k, v) in enumerate(row.items(), start=0):
        worksheet.write(0, j, k)

    i = 0
    r = 1

    for i in range(0, len(data)):
        #print("Source Row", i)


        phones = pickPhone(data,i,3,phoneColumns)
        emails = pickEmail(data,i,3,emailColumns)
        #print(emails)
        #meta = makeMetaData(data,i)
        meta = "Not Yet"

        rphones = pickPhone(data,i,3,relPhoneColumns)
        remails =  pickEmail(data,i,3,relEmailColumns)

        try:
            row = {}
            row["fn"]  = data['Grantee First Name'].iloc[i]
            row["ln"]  = data['Grantee Last Name'].iloc[i]
            row['ct']  = data['Mailing City'].iloc[i]
            row['st']  = data['Mailing State'].iloc[i]
            row['zip'] = data['Mailing Zip Code'].iloc[i]
            row['phone1'] = phones[0]
            row['phone2'] = phones[1]
            row['phone3'] = phones[2]
            row['email1'] = emails[0]
            row['email2'] = emails[1]
            row['type'] = 'grantee'
            row['meta'] = meta

            for j, (k, v) in enumerate(row.items(), start=0):
                worksheet.write(r, j, v)
            r += 1
        except:
            pass


        try:
            x = splitFullName(data['Relative Full Name'].iloc[i])
            row = {}
            row["fn"]  = x['fn']
            row["ln"]  = x['ln']
            row['ct']  = data['Relative City'].iloc[i]
            row['st']  = data['Relative State'].iloc[i]
            row['zip'] = data['Relative Zip'].iloc[i]
            row['phone1'] = rphones[0]
            row['phone2'] = rphones[1]
            row['phone3'] = rphones[2]
            row['email1'] = remails[0]
            row['email2'] = remails[1]
            row['type'] = 'relative'
            row['meta'] = meta

            for j, (k, v) in enumerate(row.items(), start=0):
                worksheet.write(r, j, v)


            r += 1
        except:
            pass


    workbook.close()

def fbHeirshipList(data,fileName):
    phoneColumns = ['CellPhones1', 'CellPhones2', 'CellPhones3','CellPhones4','CellPhones5',
                    'PhoneNumbers1','PhoneNumbers2','PhoneNumbers3','PhoneNumbers4','PhoneNumbers5'  ]
    emailColumns = ['EmailAddresses1','EmailAddresses2','EmailAddresses3','EmailAddresses4','EmailAddresses5' ]
    relPhoneColumns = ['Relative Phone1','Relative Phone2','Relative Phone3']
    relEmailColumns = ['Relative EmailAddresses1','Relative EmailAddresses2','Relative EmailAddresses3'  ]
    #ownerKeys = ['Grantee First Name','Grantee Last Name'] #,'Owner First Name','Owner Last Name']

    row = {}
    row["fn"] = ''
    row["ln"] = ''
    row['ct'] = ''
    row['st'] = ''
    row['zip'] = ''
    row['phone1'] = ''
    row['phone2'] = ''
    row['phone3'] = ''
    row['email1'] = ''
    row['email2'] = ''
    row['type'] = ''
    row['meta'] = ''

    workbook = xlsxwriter.Workbook(fileName)
    worksheet = workbook.add_worksheet()
    print("Starting on ", fileName, "length of list", len(data))

    for j, (k, v) in enumerate(row.items(), start=0):
        worksheet.write(0, j, k)

    i = 0
    r = 1

    for i in range(0, len(data)):
        #print("Source Row", i)


        phones = pickPhone(data,i,3,phoneColumns)
        emails = pickEmail(data,i,3,emailColumns)
        #print(emails)
        #meta = makeMetaData(data,i)
        meta = "Not Yet"

        rphones = pickPhone(data,i,3,relPhoneColumns)
        remails =  pickEmail(data,i,3,relEmailColumns)

        try:
            row = {}
            row["fn"]  = data['First Name'].iloc[i]
            row["ln"]  = data['Last Name'].iloc[i]
            row['ct']  = data['Mailing City'].iloc[i]
            row['st']  = data['Mailing State'].iloc[i]
            row['zip'] = data['Mailing Zip Code'].iloc[i]
            row['phone1'] = phones[0]
            row['phone2'] = phones[1]
            row['phone3'] = phones[2]
            row['email1'] = emails[0]
            row['email2'] = emails[1]
            row['type'] = 'hier'
            row['meta'] = meta

            for j, (k, v) in enumerate(row.items(), start=0):
                worksheet.write(r, j, v)
            r += 1
        except:
            pass


        try:
            x = splitFullName(data['Relative Full Name'].iloc[i])
            row = {}
            row["fn"]  = x['fn']
            row["ln"]  = x['ln']
            row['ct']  = data['Relative City'].iloc[i]
            row['st']  = data['Relative State'].iloc[i]
            row['zip'] = data['Relative Zip'].iloc[i]
            row['phone1'] = rphones[0]
            row['phone2'] = rphones[1]
            row['phone3'] = rphones[2]
            row['email1'] = remails[0]
            row['email2'] = remails[1]
            row['type'] = 'relative'
            row['meta'] = meta

            for j, (k, v) in enumerate(row.items(), start=0):
                worksheet.write(r, j, v)


            r += 1
        except:
            pass


    workbook.close()
    

def fbDivorceList(data,fileName):
  phoneColumns = ['CellPhones1', 'CellPhones2', 'CellPhones3','CellPhones4','CellPhones5',
                  'PhoneNumbers1','PhoneNumbers2','PhoneNumbers3','PhoneNumbers4','PhoneNumbers5'  ]
  emailColumns = ['EmailAddresses1','EmailAddresses2','EmailAddresses3','EmailAddresses4','EmailAddresses5' ]
  relPhoneColumns = ['Relative Phone1','Relative Phone2','Relative Phone3']
  relEmailColumns = ['Relative EmailAddresses1','Relative EmailAddresses2','Relative EmailAddresses3'  ]
  ownerKeys = ['Mortgagor First Name','Mortgagor Last Name','Owner First Name','Owner Last Name']


  row = {}
  row["fn"] = ''
  row["ln"] = ''
  row['ct'] = ''
  row['st'] = ''
  row['zip'] = ''
  row['phone1'] = ''
  row['phone2'] = ''
  row['phone3'] = ''
  row['email1'] = ''
  row['email2'] = ''
  row['type'] = ''
  row['meta'] = ''

  workbook = xlsxwriter.Workbook(fileName)
  worksheet = workbook.add_worksheet()
  print("Starting on ", fileName, "length of list", len(data))

  for j, (k, v) in enumerate(row.items(), start=0):
    worksheet.write(0, j, k)

  i = 0
  r = 1
  for i in range(0, len(data)):
    if isNotBusiness2(data,i,ownerKeys):
      mNames = getNames(data, i)
      #print("Source Row", i, mNames)
      if len(mNames) == 0:
        print('\tNo Names in row ', i)
      else:
        phones = pickPhone(data,i,3,phoneColumns)
        emails = pickEmail(data,i,3,emailColumns)
        #print(emails)
        #meta = makeMetaData(data,i)
        meta = "Not Yet"

        rphones = pickPhone(data,i,3,relPhoneColumns)
        remails =  pickEmail(data,i,3,relEmailColumns)

        try:
            row = {}
            row["fn"] = mNames['mfn0']
            row["ln"] = mNames['mln0']
            row['ct'] = data['Property City'].iloc[i]
            row['st'] = 'Texas'
            row['zip'] = data['Property Zip Code'].iloc[i]
            row['phone1'] = phones[0]
            row['phone2'] = phones[1]
            row['phone3'] = phones[2]
            row['email1'] = emails[0]
            row['email2'] = emails[1]
            row['type'] = 'mortgage first'
            row['meta'] = meta

            for j, (k, v) in enumerate(row.items(), start=0):
                if v == 'nan':
                    v = ''
                worksheet.write(r, j, v)

            #print('Added mortgage 1 to excel')
            r += 1
        except:
            pass

        try:
            row = {}
            row["fn"] = mNames['mfn1']
            row["ln"] = mNames['mln1']
            row['ct'] = data['Property City'].iloc[i]
            row['st'] = 'Texas'
            row['zip'] = data['Property Zip Code'].iloc[i]
            row['phone1'] = phones[0]
            row['phone2'] = phones[1]
            row['phone3'] = phones[2]
            row['email1'] = emails[0]
            row['email2'] = emails[1]
            row['type'] = 'mortgage second'
            row['meta'] = meta

            for j, (k, v) in enumerate(row.items(), start=0):
                if v == 'nan':
                    v = ''
                worksheet.write(r, j, v)

            #print('Added mortgage 2 to excel')
            r += 1
        except:
            pass

        try:
            row = {}
            row["fn"] = mNames['ofn0']
            row["ln"] = mNames['oln0']
            row['ct'] = data['Owner City'].iloc[i]
            row['st'] = data['Owner State'].iloc[i]
            row['zip'] = data['Owner Zip Code'].iloc[i]
            row['phone1'] = phones[0]
            row['phone2'] = phones[1]
            row['phone3'] = phones[2]
            row['email1'] = emails[0]
            row['email2'] = emails[1]
            row['type'] = 'owner first'
            row['meta'] = meta

            for j, (k, v) in enumerate(row.items(), start=0):
                if v == 'nan':
                    v = ''
                worksheet.write(r, j, v)

            #print('Added owner to excel')
            r += 1
        except:
            pass

        try:
            row = {}
            row["fn"] = mNames['ofn1']
            row["ln"] = mNames['oln1']
            row['ct'] = data['Owner City'].iloc[i]
            row['st'] = data['Owner State'].iloc[i]
            row['zip'] = data['Owner Zip Code'].iloc[i]
            row['phone1'] = phones[0]
            row['phone2'] = phones[1]
            row['phone3'] = phones[2]
            row['email1'] = emails[0]
            row['email2'] = emails[1]
            row['type'] = 'owner second'
            row['meta'] = meta

            for j, (k, v) in enumerate(row.items(), start=0):
                if v == 'nan':
                    v = ''
                worksheet.write(r, j, v)

            #print('Added owner to excel')
            r += 1
        except:
            pass

        try:
            row = {}
            row["fn"] = mNames['rfn']
            row["ln"] = mNames['rln']
            row['ct'] = data['Relative City'].iloc[i]
            row['st'] = data['Relative State'].iloc[i]
            row['zip'] = data['Relative Zip'].iloc[i]
            row['phone1'] = rphones[0]
            row['phone2'] = rphones[1]
            row['phone3'] = rphones[2]
            row['email1'] = remails[0]
            row['email2'] = remails[1]
            row['type'] = 'relative'
            row['meta'] = meta

            for j, (k, v) in enumerate(row.items(), start=0):
                if v == 'nan':
                    v = ''
                worksheet.write(r, j, v)

            #print('Added relative to excel')
            r += 1
        except:
            pass


  workbook.close()




#def xyz(csvcols,):
#  #           column    rule
#  csvcols = {['fn'    , 'prefor',
#             ['ln'    , '',
#             ['ct'    , '',
#             ['st'    , '',
#             ['zip'   , '',
#             ['phone1', '',
#             ['phone2', '',
#             ['phone3', '',
#             ['email1', '',
#             ['email2', '',
#             ['type'  , '',
#             ['meta'  , '',
#             }  
#  
#
is_64bits = sys.maxsize > 2**32
print(is_64bits)

chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'

downloads = 'C:/Users/gmcge/Downloads'


#True
#ID
#ListMonth
#ListGroup
#FileName
#Active
#URL

diffDate = "2020_11_17"
thisday = datetime.date.today().strftime("%Y_%m_%d")
output_folder = "lead_files\\"
filePath = output_folder + thisday + "\\"

if not os.path.exists(output_folder):
    try:
        os.mkdir(output_folder)
    except:
        print("Unable to create {} directory".format(output_folder))

if not os.path.exists(filePath):
    try:
        os.mkdir(filePath)
    except:
        print("Unable to create {} directory".format(filePath))




#print(thisday)

cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                      r'DBQ=E:\ReIQ_fetcher.accdb;')


#cnxn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\gmcge\Documents\GitHub\reiq_lists\ReIQ_fetcher.accdb;')


cursor = cnxn.cursor()
sql = "SELECT * FROM tbl_ListSources"
data = pd.read_sql(sql,con=cnxn)

#if not os.path.exists(directory):
#    os.makedirs(directory)


# clean up old files in download folder    
for i in range(0, len(data)):
    if data['Active'].iloc[i] == True:
        dl_file = os.path.join(downloads,data['FileName'].iloc[i])
        try:
            os.remove(dl_file)
        except OSError:
            pass
       
# use chrome to get the files
for i in range(0, len(data)):
    if data['Active'].iloc[i] == True:
        webbrowser.get(chrome_path).open(data['URL'].iloc[i])


# wait for all files to download        
for i in range(0, len(data)):
    if data['Active'].iloc[i] == True:
        dl_file = os.path.join(downloads,data['FileName'].iloc[i])
        limit = 50
        while not os.path.exists(dl_file) and limit > 0:
            sleep(0.05)
            limit -= 1

        #print(limit)
        if limit == 0:
            print("File {} was not found.  Tried {} times".format(dl_file),limit)


# open each sheet and process
for i in range(0, len(data)):
    if data['Active'].iloc[i] == True:
        dl_file = os.path.join(downloads,data['FileName'].iloc[i])
        xl = pd.ExcelFile(dl_file)
        print( "File {} has sheets {}".format(dl_file, xl.sheet_names))
        for sheet in xl.sheet_names:
            if data['ListGroup'].iloc[i] == 'Probates':
                fileName = os.path.join(filePath, 
                                        #thisday + "_" +
                                        str(data['ListGroup'].iloc[i]) + "_" +
                                        str(data['ListMonth'].iloc[i]) + "_" +
                                        sheet + ".xlsx" )
                print(fileName)
                if sheet == "Probate":
                    try:
                        xl_data = pd.read_excel(dl_file, sheet_name=sheet)
                        fbProbateList(xl_data,fileName)
                    except:
                        myErrorList.append("something went wrong processing {} {}".format(fileName,sql) )

                if sheet == "Heirship":
                    try:
                        xl_data = pd.read_excel(dl_file, sheet_name=sheet)
                        fbHeirshipList(xl_data,fileName)
                    except:
                        myErrorList.append("something went wrong processing {} {}".format(fileName,sql) )
            if data['ListGroup'].iloc[i] == 'PreFor':
                fileName = os.path.join(filePath, 
                                        #thisday + "_" +
                                        str(data['ListGroup'].iloc[i]) + "_" +
                                        str(data['ListMonth'].iloc[i]) + "_" +
                                        sheet + ".xlsx" )
                print(fileName)
                try:
                    xl_data = pd.read_excel(dl_file, sheet_name=sheet)
                    fbPreForList(xl_data,fileName)
                except:
                    myErrorList.append("something went wrong processing {} {}".format(fileName,sql) )

            if data['ListGroup'].iloc[i] == 'Divorce':
                fileName = os.path.join(filePath, 
                                        #thisday + "_" +
                                        str(data['ListGroup'].iloc[i]) + "_" +
                                        str(data['ListMonth'].iloc[i]) + "_" +
                                        sheet + ".xlsx" )
                print(fileName)
                try:
                    xl_data = pd.read_excel(dl_file, sheet_name=sheet)
                    fbDivorceList(xl_data,fileName)
                except:
                    myErrorList.append("something went wrong processing {} {}".format(fileName,sql) )

                
subfolders = [f.name for f in os.scandir(filePath + '../') if f.is_dir()]
subfolders.sort()
prev = "lead_files\\" +subfolders[len(subfolders)-2] 

subprocess.Popen([r"C:\Program Files\Beyond Compare 4\BCompare.exe", prev, filePath])

        