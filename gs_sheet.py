import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
from time import strftime

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
ws = wb.worksheet('精算用')

date = strftime("%Y/%m/%d", time.localtime())
month_date = strftime("%Y/%m", time.localtime())

PERSON1_NAME='和也'
PERSON2_NAME='花乃香'
NUMBER_COLUMN=1
DATE_COLUMN=2
TYPE_COLUMN=3
MONEY_COLUMN=4
PAY_COLUMN=5
PERSON1_COLUMN=6
PERSON2_COLUMN=7
CURRENT_NUMBER_COLUMN=8
CURRENT_NUMBER_ROW=2
NUMBER_START_ROW=2

#収支#
def money_gs_sheet(t,m,n,p):
    try:
        if int(ws.cell(CURRENT_NUMBER_ROW,CURRENT_NUMBER_COLUMN).value) > 5:
            i = int(ws.cell(CURRENT_NUMBER_ROW,CURRENT_NUMBER_COLUMN).value)
        else:
            i=NUMBER_START_ROW
        while not ws.cell(i, NUMBER_COLUMN).value == None:
            i += 1
        else:
            ws.update_cell(i,NUMBER_COLUMN,i-NUMBER_START_ROW)
            ws.update_cell(i,DATE_COLUMN,date)
            ws.update_cell(i,TYPE_COLUMN,t)
            ws.update_cell(i,MONEY_COLUMN,m+n)
            ws.update_cell(i,PAY_COLUMN,p)
            ws.update_cell(CURRENT_NUMBER_ROW,CURRENT_NUMBER_COLUMN,i-NUMBER_START_ROW)
        
        if t=='合計支出':
            money_person1 = float(ws.cell(i-1,PERSON1_COLUMN).value)-(m+n)/2
            money_person2 = float(ws.cell(i-1,PERSON2_COLUMN).value)-(m+n)/2
        elif t=='支出':
            money_person1 = float(ws.cell(i-1,PERSON1_COLUMN).value)-m
            money_person2 = float(ws.cell(i-1,PERSON2_COLUMN).value)-n
        elif t=='収入':
            money_person1 = float(ws.cell(i-1,PERSON1_COLUMN).value)+m
            money_person2 = float(ws.cell(i-1,PERSON2_COLUMN).value)+n

        ws.update_cell(i,PERSON1_COLUMN, str(money_person1))
        ws.update_cell(i,PERSON2_COLUMN, str(money_person2))
        return 'No. ' + str(i-NUMBER_START_ROW) +'\n' + str(PERSON1_NAME) + 'の残金：' + str(money_person1) + '円\n' + str(PERSON2_NAME) + 'の残金：' + str(money_person2) + '円'
    except Exception as e:
        raise
        
#「キャンセル」の関数#
def cancel_gs_sheet():
    try:
        if int(ws.cell(CURRENT_NUMBER_ROW,CURRENT_NUMBER_COLUMN).value) > 5:
            i = int(ws.cell(CURRENT_NUMBER_ROW,CURRENT_NUMBER_COLUMN).value)
        else:
            i=NUMBER_START_ROW
        while not ws.cell(i, NUMBER_COLUMN).value == None:
            i += 1
        else:
            cancel_money=str(ws.cell(i-1,MONEY_COLUMN).value)
            ws.delete_row(i-1)
            ws.update_cell(CURRENT_NUMBER_ROW,CURRENT_NUMBER_COLUMN,i-(NUMBER_START_ROW+2))
        return 'No. ' + str(i-(NUMBER_START_ROW+1)) +'を削除しました\n'+'No.'+ str(i-(NUMBER_START_ROW+1))+'\n'+'金額：' + cancel_money
    except Exception as e:
        raise

#「清算」の関数#
def monthly_gs_sheet():
    MONTH_NUMBER_COLUMN=9
    MONTH_DATE_COLUMN=10
    MONTH_MONEY2_COLUMN=11
    MONTH_MONEY_COLUMN=12
    MONTH_PAY_NAME_COLUMN=13
    MONTH_PAY_MONEY_COLUMN=14

    try:
        j=NUMBER_START_ROW
        month_money2=0
        paid_money_person1=0
        paid_money_person2=0
        while not ws.cell(j, NUMBER_COLUMN).value == None:
            if '2022/08' in ws.cell(j, DATE_COLUMN).value and ws.cell(j, TYPE_COLUMN).value == '合計支出':
                month_money2 += float(ws.cell(j, MONEY_COLUMN).value)
                if ws.cell(j, PAY_COLUMN).value == PERSON1_NAME:
                    paid_money_person1 += float(ws.cell(j, MONEY_COLUMN).value)
                if ws.cell(j, PAY_COLUMN).value == PERSON2_NAME:
                    paid_money_person2 += float(ws.cell(j, MONEY_COLUMN).value)
                j += 1
            else:
                j += 1

        month_money=month_money2/2
        month_money_person1 = paid_money_person1 - month_money
        month_money_person2 = paid_money_person2 - month_money
        if month_money_person1 < 0:
            month_pay_name=PERSON1_NAME
            month_pay_money = 0 - month_money_person1
        else:
            month_pay_name=PERSON2_NAME
            month_pay_money = 0 - month_money_person2
        
        i=NUMBER_START_ROW
        while not ws.cell(i, MONTH_NUMBER_COLUMN).value == None:
            i += 1
        else:
            ws.update_cell(i,MONTH_NUMBER_COLUMN,i-(NUMBER_START_ROW-1))
            ws.update_cell(i,MONTH_DATE_COLUMN,month_date)
            ws.update_cell(i,MONTH_MONEY2_COLUMN,month_money2)
            ws.update_cell(i,MONTH_MONEY_COLUMN,str(month_money))
            ws.update_cell(i,MONTH_PAY_NAME_COLUMN,month_pay_name)
            ws.update_cell(i,MONTH_PAY_MONEY_COLUMN,month_pay_money)

        return str(month_date)  +'\n' +'・合計支出：' + str(month_money2) + '円\n' + '・支出：' + str(month_money) + '円\n' + '・払うべき人：' + str(month_pay_name) + '\n' + '・金額：' + str(month_pay_money) + '円'
    except Exception as e:
        raise

#「清算2.1」の関数#
def smonthly_gs_sheet():
    SMONTH_NUMBER_COLUMN=16
    SMONTH_LAST_NUMBER_COLUMN=17
    SMONTH_MONEY2_COLUMN=18
    SMONTH_MONEY_COLUMN=19
    SMONTH_PAID_PERSON1_COLUMN=20
    SMONTH_PAID_PERSON2_COLUMN=21
    SMONTH_PAY_NAME_COLUMN=22
    SMONTH_PAY_MONEY_COLUMN=23

    print('!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!')
    try:
        j = 68
        smonth_money2=0
        spaid_money_person1=0
        spaid_money_person2=0
        for i in range(j, j+3):
            if ws.cell(i, TYPE_COLUMN).value == '合計支出':
                smonth_money2 += float(ws.cell(i, MONEY_COLUMN).value)
                if ws.cell(i, PAY_COLUMN).value == PERSON1_NAME:
                    spaid_money_person1 += float(ws.cell(i, MONEY_COLUMN).value)
                if ws.cell(j, PAY_COLUMN).value == PERSON2_NAME:
                    spaid_money_person2 += float(ws.cell(i, MONEY_COLUMN).value)

        smonth_money = smonth_money2 / 2
        smonth_money_person1 = spaid_money_person1 - smonth_money
        smonth_money_person2 = spaid_money_person2 - smonth_money
        if smonth_money_person1 < 0:
            smonth_pay_name = PERSON1_NAME
            smonth_pay_money = 0 - smonth_money_person1
        else:
            smonth_pay_name = PERSON2_NAME
            smonth_pay_money = 0 - smonth_money_person2
        
        k = NUMBER_START_ROW
        while not ws.cell(k, SMONTH_NUMBER_COLUMN).value == None:
            i += 1
        else:
            ws.update_cell(k, SMONTH_NUMBER_COLUMN, k-(NUMBER_START_ROW-1))
            ws.update_cell(k, SMONTH_LAST_NUMBER_COLUMN, j+3)
            ws.update_cell(k, SMONTH_MONEY2_COLUMN, smonth_money2)
            ws.update_cell(k, SMONTH_MONEY_COLUMN, str(smonth_money))
            ws.update_cell(k, SMONTH_PAID_PERSON1_COLUMN, spaid_money_person1)
            ws.update_cell(k, SMONTH_PAID_PERSON2_COLUMN, spaid_money_person2)
            ws.update_cell(k, SMONTH_PAY_NAME_COLUMN, smonth_pay_name)
            ws.update_cell(k, SMONTH_PAY_MONEY_COLUMN, smonth_pay_money)

        return ('LNo. ' + str(j+3) +'\n'+
                '・合計支出：' + str(smonth_money2) + '円\n' +
                '・支出：' + str(smonth_money) + '円\n' +
                '・和也：' + str(spaid_money_person1) + '円\n' +
                '・花乃香：' + str(spaid_money_person2) + '円\n' +
                '・払うべき人：' + str(smonth_pay_name) + '\n' +
                '・金額：' + str(smonth_pay_money) + '円')
    except Exception as e:
        raise


def calculate_gs_sheet():
    MONTH_NUMBER_COLUMN=9
    MONTH_DATE_COLUMN=10
    MONTH_MONEY2_COLUMN=11
    MONTH_MONEY_COLUMN=12
    MONTH_PAY_NAME_COLUMN=13
    MONTH_PAY_MONEY_COLUMN=14

    try:
        j=NUMBER_START_ROW
        month_money2=0
        paid_money_person1=0
        paid_money_person2=0
        while not ws.cell(j, NUMBER_COLUMN).value == None:
            if '2022/08' in ws.cell(j, DATE_COLUMN).value and ws.cell(j, TYPE_COLUMN).value == '合計支出':
                month_money2 += float(ws.cell(j, MONEY_COLUMN).value)
                if ws.cell(j, PAY_COLUMN).value == PERSON1_NAME:
                    paid_money_person1 += float(ws.cell(j, MONEY_COLUMN).value)
                if ws.cell(j, PAY_COLUMN).value == PERSON2_NAME:
                    paid_money_person2 += float(ws.cell(j, MONEY_COLUMN).value)
                j += 1
            else:
                j += 1

        month_money=month_money2/2
        month_money_person1 = paid_money_person1 - month_money
        month_money_person2 = paid_money_person2 - month_money
        if month_money_person1 < 0:
            month_pay_name=PERSON1_NAME
            month_pay_money = 0 - month_money_person1
        else:
            month_pay_name=PERSON2_NAME
            month_pay_money = 0 - month_money_person2
        
        i=NUMBER_START_ROW
        while not ws.cell(i, MONTH_NUMBER_COLUMN).value == None:
            i += 1
        else:
            ws.update_cell(i,MONTH_NUMBER_COLUMN,i-(NUMBER_START_ROW-1))
            ws.update_cell(i,MONTH_DATE_COLUMN,month_date)
            ws.update_cell(i,MONTH_MONEY2_COLUMN,month_money2)
            ws.update_cell(i,MONTH_MONEY_COLUMN,str(month_money))
            ws.update_cell(i,MONTH_PAY_NAME_COLUMN,month_pay_name)
            ws.update_cell(i,MONTH_PAY_MONEY_COLUMN,month_pay_money)

        return str(month_date)  +'\n' +'・合計支出：' + str(month_money2) + '円\n' + '・支出：' + str(month_money) + '円\n' + '・払うべき人：' + str(month_pay_name) + '\n' + '・金額：' + str(month_pay_money) + '円'
    except Exception as e:
        raise