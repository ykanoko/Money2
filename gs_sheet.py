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
ws = wb.get_worksheet(0)

date = strftime("%Y/%m/%d", time.localtime())
month_date = strftime("%Y/%m", time.localtime())

PERSON1_NAME='和也'
PERSON2_NAME='花乃香'
CURRENT_NUMBER_COLUMN=1
CURRENT_NUMBER_ROW=2
NUMBER_COLUMN=2
DATE_COLUMN=3
TYPE_COLUMN=4
MONEY_COLUMN=5
PAY_COLUMN=6
PERSON1_COLUMN=7
PERSON2_COLUMN=8
NUMBER_START_ROW=2
#支出（月）
M_MONEY2_COLUMN=9
#精算用
S_PERSON1_COLUMN=10
S_PERSON2_COLUMN=11
S_NAME_COLUMN=12
S_MONEY_PAY_COLUMN=13
#M_NUMBER_COLUMN=10
S_MONEY2_COLUMN=14
S_MONEY_COLUMN=15
S_PAID_PERSON1_COLUMN=16
S_PAID_PERSON2_COLUMN=17
S_PAY_NAME_COLUMN=18
S_PAY_MONEY_COLUMN=19

###NO.0は全項目0で初期状態？

#収支、精算#
##精算：スプシの場所の移動
##精算：合計支出の数字が大きくなりすぎるのを防ぎたい、合計支出＞○○円で合計支出、払ったお金から一定値引く？
##精算のデータがいつも30個分になるように、記入と同時に消す、前のデータがある状態にはしておく必要あり（1個上が空欄でなければ）
##精算、キャンセルにも対応できるように、おそらく行消すだけでok
###1か月分の支出：個人支出も併せて、それぞれの1カ月支出が出せたら良いね
###1か月分の支出：別のスプシに月々の支出を上の行の日付と違う場合に記入
def money_gs_sheet(t,m,n,p):
    try:
        if int(ws.cell(CURRENT_NUMBER_ROW, CURRENT_NUMBER_COLUMN).value) > 5:
            i = int(ws.cell(CURRENT_NUMBER_ROW, CURRENT_NUMBER_COLUMN).value)
        else:
            i = NUMBER_START_ROW

        while not ws.cell(i, NUMBER_COLUMN).value == None:
            i += 1
        
        ws.update_cell(i, NUMBER_COLUMN, i-NUMBER_START_ROW)  
        ws.update_cell(i, DATE_COLUMN, date)
        ws.update_cell(i, TYPE_COLUMN, t)
        ws.update_cell(i, MONEY_COLUMN, m+n)
        ws.update_cell(i, PAY_COLUMN, p)
        ws.update_cell(CURRENT_NUMBER_ROW, CURRENT_NUMBER_COLUMN, i-NUMBER_START_ROW)

        if month_date in ws.cell(i-1, DATE_COLUMN).value:
            m_money2 = int(ws.cell(i-1, M_MONEY2_COLUMN).value)
            m_money = m_money2 / 2
        else:
            m_money2 = 0
            m_money = 0
        s_person1 = float(ws.cell(i-1, S_PERSON1_COLUMN).value)
        s_person2 = float(ws.cell(i-1, S_PERSON2_COLUMN).value)
        s_name = str(ws.cell(i-1, S_NAME_COLUMN).value)
        s_money_pay = float(ws.cell(i-1, S_MONEY_PAY_COLUMN).value)
        s_money2 = int(ws.cell(i-1, S_MONEY2_COLUMN).value)
        s_money = s_money2 / 2
        s_paid_person1 = int(ws.cell(i-1, S_PAID_PERSON1_COLUMN).value)
        s_paid_person2 = int(ws.cell(i-1, S_PAID_PERSON2_COLUMN).value)
        s_pay_name = str(ws.cell(i-1, S_PAY_NAME_COLUMN).value)
        s_pay_money = float(ws.cell(i-1, S_PAY_MONEY_COLUMN).value)

        if t == '収入':
            money_person1 = float(ws.cell(i-1,PERSON1_COLUMN).value) + m
            money_person2 = float(ws.cell(i-1,PERSON2_COLUMN).value) + n
        
        elif t == '支出':
            money_person1 = float(ws.cell(i-1,PERSON1_COLUMN).value) - m
            money_person2 = float(ws.cell(i-1,PERSON2_COLUMN).value) - n

        elif t == '合計支出':
            money_person1 = float(ws.cell(i-1,PERSON1_COLUMN).value) - (m+n)/2
            money_person2 = float(ws.cell(i-1,PERSON2_COLUMN).value) - (m+n)/2

            #月支出
            m_money2 += m+n
            m_money = m_money2 / 2

            #精算
            s_person1 += m/2 - n/2
            s_person2 -= m/2 - n/2
            if s_person1 <0:
                s_name = PERSON1_NAME
                s_money_pay = 0 - s_person1
            else:
                s_name = PERSON2_NAME
                s_money_pay = 0 - s_person2

            #精算old
            s_money2 += m+n
            s_paid_person1 += m
            s_paid_person2 += n

            s_money = s_money2 / 2
            s_money_person1 = s_paid_person1 - s_money
            s_money_person2 = s_paid_person2 - s_money
            if s_money_person1 < 0:
                s_pay_name = PERSON1_NAME
                s_pay_money = 0 - s_money_person1
            else:
                s_pay_name = PERSON2_NAME
                s_pay_money = 0 - s_money_person2
            
        ws.update_cell(i, PERSON1_COLUMN, str(money_person1))
        ws.update_cell(i, PERSON2_COLUMN, str(money_person2))
        ws.update_cell(i, M_MONEY2_COLUMN, m_money2)
        ws.update_cell(i, S_PERSON1_COLUMN, str(s_person1))
        ws.update_cell(i, S_PERSON2_COLUMN, str(s_person2))
        ws.update_cell(i, S_NAME_COLUMN, str(s_name))
        ws.update_cell(i, S_MONEY_PAY_COLUMN, str(s_money_pay))        
        #ws.update_cell(j, S_NUMBER_COLUMN, j-(NUMBER_START_ROW-1))
        ws.update_cell(i, S_MONEY2_COLUMN, s_money2)
        ws.update_cell(i, S_MONEY_COLUMN, str(s_money))
        ws.update_cell(i, S_PAID_PERSON1_COLUMN, s_paid_person1)
        ws.update_cell(i, S_PAID_PERSON2_COLUMN, s_paid_person2)
        ws.update_cell(i, S_PAY_NAME_COLUMN, s_pay_name)
        ws.update_cell(i, S_PAY_MONEY_COLUMN, str(s_pay_money))

        return ('[残金]\n'
                'No. ' + str(i-NUMBER_START_ROW) +'\n' + 
                '・' + str(PERSON1_NAME) + 'の残金：' + str(money_person1) + '円\n' + 
                '・' + str(PERSON2_NAME) + 'の残金：' + str(money_person2) + '円\n'
                '[月支出]\n'
                '・合計：' + str(m_money2) + '円\n' +
                '・1人当たり：' + str(m_money) + '円\n' +
                '[精算]\n'
                '・和也：' + str(s_person1) + '円\n' +
                '・花乃香：' + str(s_person2) + '円\n' +
                #'No. ' + str(j-(NUMBER_START_ROW-1)) +'\n'+
                # '・合計支出：' + str(s_money2) + '円\n' +
                # '・和也：' + str(s_paid_person1) + '円\n' +
                # '・花乃香：' + str(s_paid_person2) + '円\n' +
                '・払うべき人：' + str(s_pay_name) + '\n' +
                '・金額：' + str(s_pay_money) + '円')
    except Exception as e:
        raise
        
#「キャンセル」の関数#
##精算（改訂版）に対応する必要あり
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
def smonthly_gs_sheet(n):
    SMONTH_NUMBER_COLUMN=16
    SMONTH_LAST_NUMBER_COLUMN=17
    SMONTH_MONEY2_COLUMN=18
    SMONTH_MONEY_COLUMN=19
    SMONTH_PAID_PERSON1_COLUMN=20
    SMONTH_PAID_PERSON2_COLUMN=21
    SMONTH_PAY_NAME_COLUMN=22
    SMONTH_PAY_MONEY_COLUMN=23

    try:
        j = n+2
        N_REPEAT = 5
        LAST_INDEX = j + N_REPEAT - 3

        #smonth_money2=0
        #spaid_money_person1=0
        #spaid_money_person2=0
        k = NUMBER_START_ROW
        while not ws.cell(k, SMONTH_NUMBER_COLUMN).value == None:
            k += 1
        SMONTH_ROW = k

        smonth_money2 = int(ws.cell(SMONTH_ROW-1, SMONTH_MONEY2_COLUMN).value)
        spaid_money_person1 = int(ws.cell(SMONTH_ROW-1, SMONTH_PAID_PERSON1_COLUMN).value)
        spaid_money_person2 = int(ws.cell(SMONTH_ROW-1, SMONTH_PAID_PERSON2_COLUMN).value)
        for i in range(j, j+N_REPEAT):
            if ws.cell(i, TYPE_COLUMN).value == '合計支出':
                smonth_money2 += float(ws.cell(i, MONEY_COLUMN).value)
                if ws.cell(i, PAY_COLUMN).value == PERSON1_NAME:
                    spaid_money_person1 += float(ws.cell(i, MONEY_COLUMN).value)
                if ws.cell(i, PAY_COLUMN).value == PERSON2_NAME:
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
  
        ws.update_cell(SMONTH_ROW, SMONTH_NUMBER_COLUMN, SMONTH_ROW-(NUMBER_START_ROW-1))
        ws.update_cell(SMONTH_ROW, SMONTH_LAST_NUMBER_COLUMN, LAST_INDEX)
        ws.update_cell(SMONTH_ROW, SMONTH_MONEY2_COLUMN, smonth_money2)
        ws.update_cell(SMONTH_ROW, SMONTH_MONEY_COLUMN, str(smonth_money))
        ws.update_cell(SMONTH_ROW, SMONTH_PAID_PERSON1_COLUMN, spaid_money_person1)
        ws.update_cell(SMONTH_ROW, SMONTH_PAID_PERSON2_COLUMN, spaid_money_person2)
        ws.update_cell(SMONTH_ROW, SMONTH_PAY_NAME_COLUMN, smonth_pay_name)
        ws.update_cell(SMONTH_ROW, SMONTH_PAY_MONEY_COLUMN, smonth_pay_money)

        return ('LNo. ' + str(LAST_INDEX) +'\n'+
                '・合計支出：' + str(smonth_money2) + '円\n' +
                '・支出：' + str(smonth_money) + '円\n' +
                '・和也：' + str(spaid_money_person1) + '円\n' +
                '・花乃香：' + str(spaid_money_person2) + '円\n' +
                '・払うべき人：' + str(smonth_pay_name) + '\n' +
                '・金額：' + str(smonth_pay_money) + '円')
    except Exception as e:
        raise