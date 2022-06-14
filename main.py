from flask import Flask, request, abort

from linebot import (
    LineBotApi, WebhookHandler
)
from linebot.exceptions import (
    InvalidSignatureError
)
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage,
)
import os

from gs_sheet import PERSON1_NAME, PERSON2_NAME, pay_sum_gs_sheet, pay_gs_sheet, gain_gs_sheet, cancel_gs_sheet, monthly_gs_sheet

app = Flask(__name__)
#環境変数取得
YOUR_CHANNEL_ACCESS_TOKEN = os.environ["YOUR_CHANNEL_ACCESS_TOKEN"]
YOUR_CHANNEL_SECRET = os.environ["YOUR_CHANNEL_SECRET"]

line_bot_api = LineBotApi(YOUR_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(YOUR_CHANNEL_SECRET)

@app.route("/callback", methods=['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']

    # get request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)

    # handle webhook body
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return 'OK'


@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    if event.message.text=='ヘルプ':
        help = "＜入力方法一覧＞\n" + "・「合計支出（払った人）（金額）」\n"+"2人で使ったお金を記録します。\n" + "・「支出（名前）（金額）」\n"+"個人で使ったお金を記録します。\n" + "・「収入（名前）（金額）」\n"+"個人の収入を記録します。\n"+ "・「キャンセル」\n最後の項目を削除します。" 
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=help))
    
    if event.message.text[:4]=='合計支出':
        if PERSON1_NAME in event.message.text[4:]:
            pay_sum_person1 = pay_sum_gs_sheet(int(event.message.text[4+len(PERSON1_NAME):]),PERSON1_NAME)
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text=pay_sum_person1))
        else:
            pay_sum_person2 = pay_sum_gs_sheet(int(event.message.text[4+len(PERSON2_NAME):]),PERSON2_NAME)
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text=pay_sum_person2))

    if event.message.text[:2]=='支出':
        if PERSON1_NAME in event.message.text[2:]:
            pay_person1 = pay_gs_sheet(int(event.message.text[2+len(PERSON1_NAME):]),int(0))
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text=pay_person1))
        else:
            pay_person2 = pay_gs_sheet(int(0),int(event.message.text[2+len(PERSON2_NAME):]))
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text=pay_person2))
    
    #収入を入力
    if event.message.text[:2]=='収入':
        if PERSON1_NAME in event.message.text[2:]:
            gain_person1 = gain_gs_sheet(int(event.message.text[2+len(PERSON1_NAME):]),int(0))
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text=gain_person1))
        else:
            gain_person2 = gain_gs_sheet(int(0),int(event.message.text[2+len(PERSON2_NAME):]))
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text=gain_person2))

    if event.message.text == 'キャンセル':
        cancel = cancel_gs_sheet()
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=cancel))

    if event.message.text == '精算':
        seisan=monthly_gs_sheet()
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=seisan))

            
    #if event.message.text=='スプシテスト':
        #message_text = test_gs_sheet()
        #line_bot_api.reply_message(
            #event.reply_token,
            #TextSendMessage(text=message_text))
    #else:
        #line_bot_api.reply_message(
            #event.reply_token,
            #TextSendMessage(text=event.message.text))


if __name__ == "__main__":
#    app.run()
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)     