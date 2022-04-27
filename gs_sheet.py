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
ws = wb.get_worksheet(0)

PERSON1_NAME='和也'
PERSON2_NAME='花乃香'
NUMBER_COLUMN=2
DATE_COLUMN=3
MONEY_COLUMN=4
PAY_COLUMN=5
PERSON1_COLUMN=6
PERSON2_COLUMN=7
NUMBER_START_ROW=5

#「支出」の関数#
def pay_gs_sheet(p,q):
    #番号・日付date・2人で使った金額mの入力#
    import time
    from time import strftime
    date = strftime("%Y/%m/%d", time.localtime())
    #↑関数の外に出せない？毎回使う
    i=NUMBER_START_ROW
    while not ws.cell(i, NUMBER_COLUMN).value == None:
        i += 1
    else:
        ws.update_cell(i,NUMBER_COLUMN,i-(NUMBER_START_ROW-1))
        ws.update_cell(i,DATE_COLUMN,date)
        ws.update_cell(i,MONEY_COLUMN,p)
        ws.update_cell(i,PAY_COLUMN,q)

    #残金の計算#
    pay_money_kazuya = int(ws.cell(i-1,PERSON1_COLUMN).value)-p/2
    pay_money_kanoko = int(ws.cell(i-1,PERSON2_COLUMN).value)-p/2
    ws.update_cell(i,PERSON1_COLUMN, pay_money_kazuya)
    ws.update_cell(i,PERSON2_COLUMN, pay_money_kanoko)
    return 'No. ' + str(i-(NUMBER_START_ROW-1)) +'\n' + str(PERSON1_NAME) + 'の残金：' + str(pay_money_kazuya) + '円\n' + str(PERSON2_NAME) + 'の残金：' + str(pay_money_kanoko) + '円'

#「収入(PERSON1)」の関数#
def gain_gs_sheet(g,h):
    #番号・日付date・収入gの入力#
    import time
    from time import strftime
    date = strftime("%Y/%m/%d", time.localtime())

    i=NUMBER_START_ROW
    while not ws.cell(i, NUMBER_COLUMN).value == None:
        i += 1
    else:
        ws.update_cell(i,NUMBER_COLUMN,i-(NUMBER_START_ROW-1))
        ws.update_cell(i,DATE_COLUMN,date)
        ws.update_cell(i,MONEY_COLUMN,g+h)

    #残金の計算#
    gain_money_person1 = int(ws.cell(i-1,PERSON1_COLUMN).value)+g
    gain_money_person2 = int(ws.cell(i-1,PERSON2_COLUMN).value)+h
    ws.update_cell(i,PERSON1_COLUMN, gain_money_person1)
    ws.update_cell(i,PERSON2_COLUMN, gain_money_person2)
    return 'No. ' + str(i-(NUMBER_START_ROW-1)) +'\n' + str(PERSON1_NAME) + 'の残金：' + str(gain_money_person1) + '円\n'+ str(PERSON2_NAME) + 'の残金：' + str(gain_money_person2) + '円'

#「キャンセル」の関数#
def cancel_gs_sheet():
    i=NUMBER_START_ROW
    while not ws.cell(i, NUMBER_COLUMN).value == None:
        i += 1
    else:
        cancel_money=str(ws.cell(i-1,MONEY_COLUMN).value)
        ws.delete_row(i-1)
    return 'No. ' + str(i-NUMBER_START_ROW) +'を削除しました\n'+'No.'+ str(i-NUMBER_START_ROW)+'\n'+'金額：' + cancel_money

#清算の関数#
def monthly_gs_sheet():
    import time
    from time import strftime
    mounth_date = strftime("%Y/%m", time.localtime())
    MOUNTH_NUMBER_COLUMN=8
    MOUNTH_DATE_COLUMN=9
    MOUNTH_MONEY_COLUMN=10

    ws.update_cell(4,MOUNTH_NUMBER_COLUMN,0)
    ws.update_cell(4,MOUNTH_DATE_COLUMN,mounth_date)

    i=NUMBER_START_ROW
    mounth_money=0
    while not ws.cell(i, NUMBER_COLUMN).value == None:
        if mounth_date not in ws.cell(i, DATE_COLUMN).value:
            i += 1
        else:
            mounth_money += int(ws.cell(i, MONEY_COLUMN).value)
            i += 1
    ws.update_cell(4,MOUNTH_MONEY_COLUMN,mounth_money)
    return str(mounth_date) +'の合計支出：' + str(mounth_money) + '円'

      #2つめのシートに月の支出を入力
    #支出の半額から払った分を引いたりして、払うべき人と金額を算出