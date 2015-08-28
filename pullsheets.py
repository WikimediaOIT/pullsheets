#!/usr/bin/python

# Google oauth secrice account credentials file
# http://gspread.readthedocs.org/en/latest/oauth2.html
# be sure to add service account email to document with View permissions
import gspread
import json
from oauth2client.client import SignedJwtAssertionCredentials

import unicodedata

import cPickle as pickle
import time
picklefile = 'pullsheets.pkl'

import pprint
pp = pprint.PrettyPrinter(indent=4)

# Pull configuration from local file
from pullsheets_config import *

# Google Oauth2 Stuff
json_key = json.load(open(json_file))
credentials = SignedJwtAssertionCredentials(
    json_key['client_email'],
    json_key['private_key'],
    ['https://spreadsheets.google.com/feeds']
    )
gc = gspread.authorize(credentials)

# Try to read last pickle file
try:
    lastdata =  pickle.load( open( picklefile, "rb" ) )
except:
    lastdata = []

# Storage for current data
bigdata = []

# Start looking at spreadsheets
for spreadsheet in sheetlist:
    print 'Speadsheet:', spreadsheet
    # Look at sub worksheets
    for subsheet in sheetlist[spreadsheet]:
        print '\tWorksheet', subsheet
        worksheet = gc.open(spreadsheet).worksheet(subsheet)

        # Walk each row of the sheet
        for row in worksheet.get_all_values():
            # Start a dictionary object for this row
            output=dict()
            output['Spreadsheet'] = spreadsheet
            output['Worksheet'] = subsheet

            # Decode fields and store values
            for field,location in sheetlist[spreadsheet][subsheet].items():
                output[field] = row[location]

            # Bounds checking, ignore this row
            if output['FirstName'] == 'First Name': continue
            if len(output['FirstName']) == 0: continue

            # Unicode decode (umlauts!)
            try: output['LastName'] = unicodedata.normalize(
                    'NFKD',
                    output['LastName']).encode('ascii','ignore')
            except: pass

            # Infer email address
            try:
                email_guess = "%s%s@wikimedia.org" % (
                    output['FirstName'][0].lower(), output['LastName'].lower())
            except:
                print 'PARSE ERROR', row
            output['Email (guess)'] = email_guess

            bigdata.append(output)


# Compare lastdata with bigdata
changefound=False
for record in bigdata:
    if record not in lastdata:
        print '-'*80
        print 'Change Found:'
        print json.dumps(record, sort_keys=True, indent=4)
        changefound=True

if changefound:
    pickle.dump( bigdata, open( picklefile, "wb" ) )
    picklefile_ts = "%s-%s" % (picklefile, time.strftime("%Y%m%d-%H%M"))
    pickle.dump( bigdata, open( picklefile_ts, "wb" ) )
