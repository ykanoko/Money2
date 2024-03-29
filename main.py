from flask import Flask, request, abort

from linebot import (LineBotApi, WebhookHandler)
from linebot.exceptions import (InvalidSignatureError)
from linebot.models import (
    MessageEvent,
    TextMessage,
    TextSendMessage,
)
import os
import json

from gs_sheet import PERSON1_NAME, PERSON2_NAME, money_gs_sheet, cancel_gs_sheet, smonthly_gs_sheet

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


@app.route("/hello", methods=['GET'])
def hello():
    return "OK"


@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    if event.message.text == 'ヘルプ' or event.message.text == 'へるぷ':
        help = "＜入力方法一覧＞\n" + "・「合計支出（払った人）（金額）」\n" + "2人で使ったお金を記録します。\n" + "・「支出（名前）（金額）」\n" + "個人で使ったお金を記録します。\n" + "・「収入（名前）（金額）」\n" + "個人の収入を記録します。\n" + "・「キャンセル」\n最後の項目を削除します。"
        try:
            # event.source(object)を返信
            group_id = event.source.group_id
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=group_id))
            # line_bot_api.reply_message(event.reply_token, TextSendMessage(text=json.dumps(event)))
        except:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=help))

    #収支#
    if PERSON1_NAME in event.message.text or PERSON2_NAME in event.message.text:
        if '合計支出' in event.message.text:
            t = '合計支出'
        elif '支出' in event.message.text:
            t = '支出'
        elif '収入' in event.message.text:
            t = '収入'

        if PERSON1_NAME in event.message.text:
            p = PERSON1_NAME
            m = int(event.message.text[len(t) + len(p):])
            n = 0
        else:
            p = PERSON2_NAME
            m = 0
            n = int(event.message.text[len(t) + len(p):])

        try:
            money = money_gs_sheet(t, m, n, p)
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=money))
        except Exception as error:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=str(error)))

    #キャンセル#
    if event.message.text == 'キャンセル':
        try:
            cancel = cancel_gs_sheet()
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=cancel))
        except Exception as error:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=str(error)))

    # #精算#
    # #if event.message.text == '精算':
    # if '精算' in event.message.text:
    #     n = int(event.message.text[len('精算'):])
    #     try:
    #         seisan = smonthly_gs_sheet(n)
    #         line_bot_api.reply_message(
    #             event.reply_token,
    #             TextSendMessage(text=seisan))
    #     except Exception as error:
    #         line_bot_api.reply_message(
    #             event.reply_token,
    #             TextSendMessage(text=str(error)))


if __name__ == "__main__":
    #    app.run()
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
