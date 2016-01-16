# -*- coding: utf-8 -*-
import requests
import xlrd
import sys
reload(sys)
sys.setdefaultencoding('utf8')

def open_file(url, filename):
    r = requests.get(url)
    if r.status_code == 200:
        with open(filename, 'wb') as f:
            f.write(r.content)
    xlsfile = xlrd.open_workbook(filename)
    return xlsfile

# The title page has titles in varying columns.
def get_offices(xlsfile,column=1):
    offices = []
    sheet = xlsfile.sheets()[0]
    last = sheet.nrows-1
    if last == 1:
        rows = [1]
    else:
        rows = range(1,last)
    for i in rows:
        offices.append(sheet.row_values(i)[column])
    return offices

def detect_headers(sheet):
    for i in range(3,12):
        if sheet.row_values(i)[2].strip() == 'Total Votes Cast':
            if 'REP' in sheet.row_values(i) or 'DEM' in sheet.row_values(i):
                parties = [x for x in sheet.row_values(i)[3:] if x != None]
                candidates = [x for x in sheet.row_values(i+1)[3:] if x!= None]
                start_row = i+2
            else:
                parties = sheet.row_values(i-1)[3:]
                candidates = sheet.row_values(i)[3:]
                start_row = i+1
            return [zip(candidates, parties), start_row]
        else:
            continue

def parse_sheet(sheet, office):
    output = []
    combo, start_row = detect_headers(sheet)
    if 'DISTRICT' in office.upper():
        split = office.split('-')
        if (len(split) == 1):
          # This '–' is a different character than this '–'
          office = office.replace('–','-')
          split = office.split('-')
        # Office string comes in formats:
        #  * STATE SENATE - DISTRICT 1 - REPUBLICAN
        #  * STATE SENATE   DISTRICT 1 - REPUBLICAN
        if (len(split) == 2):
            try:
                office, district = office.split(' - ')
            except:
                office, district = office.split(u' – ')
        # Assumes STATE SENATE - DISTRICT 1 - REPUBLICAN format
        else:
          try:
              office, party, district = office.split(' - ')
          except:
              office, party, district = office.split(u' – ')
        district = district.replace('DISTRICT ','')
    else:
        district = None
    for i in range(start_row, sheet.nrows):
        results = sheet.row_values(i)
        if "Totals" in results[1]:
            continue
        if results[0].strip() != '':
            county = results[0].strip()
        else:
          county = county
        ward = results[1].strip()
        total_votes = int(results[2]) if results[2] else results[2]
        candidate_votes = results[3:]
        for candidate, party in combo:
            index = [x[0] for x in combo].index(candidate)
            if candidate == None:
                continue
            elif candidate.strip() == '':
                continue
            else:
                output.append([county, ward, office, district, total_votes, party, candidate, candidate_votes[index]])
    return output

def process_all(url, filename):
    results = []
    xlsfile = open_file(url, filename)
    offices = get_offices(xlsfile)
    for office in offices:
        index = [x for x in offices].index(office)
        sheet = xlsfile.sheets()[index+1]
        print "parsing %s" % office
        results.append(parse_sheet(sheet, office))

    return [r for result in results for r in result]
