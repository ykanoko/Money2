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

#支出の関数#
def pay_gs_sheet(p):
    #番号・日付date・2人で使った金額mの入力#
    import time
    from time import strftime
    date = strftime("%Y/%m/%d", time.localtime())
    i = 5
    while not ws.cell(i, 2).value == None:
        print(ws.cell(i, 2).value)
        #コードが動いてるか確認用(完成時に消す)
        i += 1
    else:
        ws.update_cell(i,2,i-4)
        ws.update_cell(i,3,date)
        ws.update_cell(i,4,p)

    #残金の計算#
    money_kazuya = int(ws.cell(i-1,6).value)-p/2
    money_kanoko = int(ws.cell(i-1,7).value)-p/2
    ws.update_cell(i,6, money_kazuya)
    ws.update_cell(i,7, money_kanoko)
    return '番号：' + str(i-4) +'\n' + '和也の残金：' + str(money_kazuya) + '円\n' + '花乃香の残金：' + str(money_kanoko) + '円'

money_kazuya=20
money_kanoko=10
print('番号：' + str(i-4) +'\n' + '和也の残金：' + str(money_kazuya) + '円\n' + '花乃香の残金：' + str(money_kanoko) + '円')