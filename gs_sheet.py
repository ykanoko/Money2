import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

credential = {
                "type": "service_account",
                "project_id": os.environ['SHEET_PROJECT_ID'],
                "private_key_id": os.environ['SHEET_PRIVATE_KEY_ID'],
                "private_key": os.environ['SHEET_PRIVATE_KEY'],
                "client_email": os.environ['SHEET_CLIENT_EMAIL'],
                "client_id": os.environ['SHEET_CLIENT_ID'],
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                "client_x509_cert_url":  os.environ['SHEET_CLIENT_X509_CERT_URL']
             }

credentials = ServiceAccountCredentials.from_json_keyfile_dict(credential, scope)
gc = gspread.authorize(credentials)

SPREADSHEET_KEY = '1AjXVHcDBE32vbCVxwTCcqzHj0olxE6UlapdigoBELGs'
wb = gc.open_by_key(SPREADSHEET_KEY)
ws = wb.sheet1
PERSON1_NAME='和也'
PERSON2_NAME='花乃香'
PERSON1_COLUMN=6
PERSON2_COLUMN=7
NUMBER_COLUMN=2
NUMBER_START_ROW=4
DATE_COLUMN=3
MONEY_COLUMN=4

#「支出」の関数#
def pay_gs_sheet(p):
    #番号・日付date・2人で使った金額mの入力#
    import time
    from time import strftime
    date = strftime("%Y/%m/%d", time.localtime())
    #PERSON1_NAME='和也'
    #PERSON2_NAME='花乃香'
    #PERSON1_COLUMN=6
    #PERSON2_COLUMN=7
    i=NUMBER_START_ROW+1
    while not ws.cell(i, NUMBER_COLUMN).value == None:
        print(ws.cell(i, NUMBER_COLUMN).value)
    else:
        ws.update_cell(i,NUMBER_COLUMN,i-NUMBER_START_ROW)
        ws.update_cell(i,DATE_COLUMN,date)
        ws.update_cell(i,MONEY_COLUMN,p)

    #残金の計算#
    money_kazuya = int(ws.cell(i-1,PERSON1_COLUMN).value)-p/2
    money_kanoko = int(ws.cell(i-1,PERSON2_COLUMN).value)-p/2
    ws.update_cell(i,PERSON1_COLUMN, money_kazuya)
    ws.update_cell(i,PERSON2_COLUMN, money_kanoko)
    return 'No. ' + str(i-NUMBER_START_ROW) +'\n' + str(PERSON1_NAME) + 'の残金：' + str(money_kazuya) + '円\n' + str(PERSON2_NAME) + 'の残金：' + str(money_kanoko) + '円'

#「キャンセル」の関数#
def cancel_gs_sheet():
    i=NUMBER_START_ROW
    while not ws.cell(i, NUMBER_COLUMN).value == None:
        i += 1
    else:
        ws.delete_row(i-1)
    return 'No. ' + str(i-(NUMBER_START_ROW+1)) +'を削除しました'