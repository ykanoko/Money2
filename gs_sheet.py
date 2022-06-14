import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
from time import strftime
date = strftime("%Y/%m/%d", time.localtime())

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
TYPE_COLUMN=4
MONEY_COLUMN=5
PAY_COLUMN=6
PERSON1_COLUMN=7
PERSON2_COLUMN=8
NUMBER_START_ROW=5

#「合計支出」の関数#
def pay_sum_gs_sheet(t,m,p):
    i=NUMBER_START_ROW
    while not ws.cell(i, NUMBER_COLUMN).value == None:
        i += 1
    else:
        ws.update_cell(i,NUMBER_COLUMN,i-(NUMBER_START_ROW-1))
        ws.update_cell(i,DATE_COLUMN,date)
        ws.update_cell(i,TYPE_COLUMN,t)
        ws.update_cell(i,MONEY_COLUMN,m)
        ws.update_cell(i,PAY_COLUMN,p)

    #残金の計算#
    pay_sum_money_person1 = int(ws.cell(i-1,PERSON1_COLUMN).value)-m/2
    pay_sum_money_person2 = int(ws.cell(i-1,PERSON2_COLUMN).value)-m/2
    ws.update_cell(i,PERSON1_COLUMN, pay_sum_money_person1)
    ws.update_cell(i,PERSON2_COLUMN, pay_sum_money_person2)
    return 'No. ' + str(i-(NUMBER_START_ROW-1)) +'\n' + str(PERSON1_NAME) + 'の残金：' + str(pay_sum_money_person1) + '円\n' + str(PERSON2_NAME) + 'の残金：' + str(pay_sum_money_person2) + '円'

#「（個人）支出」の関数#
def pay_gs_sheet(t,m,n,p):
    i=NUMBER_START_ROW
    while not ws.cell(i, NUMBER_COLUMN).value == None:
        i += 1
    else:
        ws.update_cell(i,NUMBER_COLUMN,i-(NUMBER_START_ROW-1))
        ws.update_cell(i,DATE_COLUMN,date)
        ws.update_cell(i,TYPE_COLUMN,t)
        ws.update_cell(i,MONEY_COLUMN,m+n)
        ws.update_cell(i,PAY_COLUMN,p)

    #残金の計算#
    pay_money_person1 = int(ws.cell(i-1,PERSON1_COLUMN).value)-m
    pay_money_person2 = int(ws.cell(i-1,PERSON2_COLUMN).value)-n
    ws.update_cell(i,PERSON1_COLUMN, pay_money_person1)
    ws.update_cell(i,PERSON2_COLUMN, pay_money_person2)
    return 'No. ' + str(i-(NUMBER_START_ROW-1)) +'\n' + str(PERSON1_NAME) + 'の残金：' + str(pay_money_person1) + '円\n' + str(PERSON2_NAME) + 'の残金：' + str(pay_money_person2) + '円'

#「収入」の関数#
def gain_gs_sheet(t,m,n,p):
    i=NUMBER_START_ROW
    while not ws.cell(i, NUMBER_COLUMN).value == None:
        i += 1
    else:
        ws.update_cell(i,NUMBER_COLUMN,i-(NUMBER_START_ROW-1))
        ws.update_cell(i,DATE_COLUMN,date)
        ws.update_cell(i,TYPE_COLUMN,t)
        ws.update_cell(i,MONEY_COLUMN,m+n)
        ws.update_cell(i,PAY_COLUMN,p)

    #残金の計算#
    gain_money_person1 = int(ws.cell(i-1,PERSON1_COLUMN).value)+m
    gain_money_person2 = int(ws.cell(i-1,PERSON2_COLUMN).value)+n
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

#「清算」の関数#
def monthly_gs_sheet():
    MONTH_NUMBER_COLUMN=10
    MONTH_DATE_COLUMN=11
    MONTH_MONEY2_COLUMN=12
    MONTH_MONEY_COLUMN=13
    MONTH_PAY_NAME_COLUMN=14
    MONTH_PAY_MONEY_COLUMN=15

    ws.update_cell(4,MONTH_NUMBER_COLUMN,0)
    ws.update_cell(4,MONTH_DATE_COLUMN,month_date)

    i=NUMBER_START_ROW
    month_money2=0
    paid_money_person1=0
    paid_money_person2=0
    while not ws.cell(i, NUMBER_COLUMN).value == None:
        if month_date in ws.cell(i, DATE_COLUMN).value and ws.cell(i, PAY_COLUMN).value != None:
            month_money2 += int(ws.cell(i, MONEY_COLUMN).value)
            print(month_money2)
            if ws.cell(i, PAY_COLUMN).value == PERSON1_NAME:
                paid_money_person1 += int(ws.cell(i, MONEY_COLUMN).value)
            if ws.cell(i, PAY_COLUMN).value == PERSON2_NAME:
                paid_money_person2 += int(ws.cell(i, MONEY_COLUMN).value)
            i += 1
        else:
            i += 1

    month_money=month_money2/2
    month_money_person1 = paid_money_person1 - month_money
    month_money_person2 = paid_money_person2 - month_money
    if month_money_person1 < 0:
        month_pay_name=PERSON1_NAME
        month_pay_money = 0 - month_money_person1
    else:
        month_pay_name=PERSON2_NAME
        month_pay_money = 0 - month_money_person2
    ws.update_cell(4,MONTH_PAY_NAME_COLUMN,month_pay_name)
    ws.update_cell(4,MONTH_PAY_MONEY_COLUMN,month_pay_money)
    ws.update_cell(4,MONTH_MONEY2_COLUMN,month_money2)
    ws.update_cell(4,MONTH_MONEY_COLUMN,month_money)

    return str(month_date)  +'\n' +'・合計支出：' + str(month_money2) + '円\n' + '・支出：' + str(month_money) + '円\n' + '・払うべき人：' + str(month_pay_name) + '\n' + '・金額：' + str(month_pay_money) + '円'

      #2つめのシートに月の支出を入力