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

from gs_sheet import cancel_gs_sheet, pay_gs_sheet, gain_person1_gs_sheet, gain_person2_gs_sheet

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
        help = "＜入力方法一覧＞\n" + "・「支出〇〇」\n"+"2人で使ったお金を記録します。\n" + "・「キャンセル」\n最後の項目を削除します。\n" + "・「収入〇〇」\n"+"2人で使ったお金を記録します。"
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=help))
    
    if event.message.text[:2]=='支出':
        #[:2]←先頭の2文字(=支出)を指定
        zannkinn = pay_gs_sheet(int(event.message.text[2:]))
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=zannkinn))
    
    if event.message.text[:4]=='収入和也':
        money_person1 = gain_person1_gs_sheet(int(event.message.text[4:]))
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=money_person1))
    
    if event.message.text[:5]=='収入花乃香':
        money_person2 = gain_person2_gs_sheet(int(event.message.text[5:]))
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=money_person2))

    if event.message.text == 'キャンセル':
        cancel = cancel_gs_sheet()
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=cancel))
            
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