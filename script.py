SHEET = 'Log'
PATH = 'D:/Desktop/T'
HEADERS = ['Date', 'Time', 'Department', 'Category', 'Year', 'File', 'Remarks']


import traceback
import gspread
from os import path
from glob import glob
from datetime import datetime
from time import sleep



def getdate(f):
    date = max(path.getctime(f), path.getmtime(f))
    return datetime.strftime(datetime.fromtimestamp(date), "%Y-%m-%d %H:%M:%S").split()

def diff_lists(list1, list2):

    for l1 in list1:
        if l1 != '' and l1 not in list2:            
            return True
    
    for l2 in list2:
        if l2 != '' and l2 not in list1:
            return True

    return False


def dict_in_list_search(dictlist, keys, values):
    for d in dictlist:
        flag = False
        for i in range(len(keys)):
            if d[keys[i]] != values[i]:
                flag = True
                break
        if flag:
            continue
        return d
    return {'Remarks':''}

wks = gspread.service_account().open(SHEET).get_worksheet(0)


while True:
    try:
        gsheet_records = wks.get_all_records()
        
        gsheet_rows = [list(row.values()) for row in gsheet_records]
        paths = glob(path.join(PATH, "**/*"), recursive=True)

        files = [f.replace('\\\\', '/').replace('\\', '/') for f in paths if path.isfile(f)]
        
        rows = [sortedrow[1:] for sortedrow in sorted([ [max(path.getctime(f), path.getmtime(f))] + getdate(f) + f.split('/')[-4:] + [dict_in_list_search(gsheet_records, HEADERS[2:6], f.split('/')[-4:])['Remarks']]  for f in files], reverse=True)]
    
        if diff_lists(rows, gsheet_rows):
            
            wks.batch_update([{'range': '', 'values':[HEADERS]+rows}])
            
            print('Sheet updated', str(datetime.now()))

    except KeyError:
        wks.batch_update([{'range': 'A1:G1', 'values': [HEADERS]}])

    except Exception as err:
        traceback.print_exc()
        sleep(10)
    
    sleep(2)
