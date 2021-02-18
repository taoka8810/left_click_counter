#
#click_counter.pyで記録したExcelファイルを読み込んで、左クリックした回数をLINEで通知するプログラム 
#

import json, os
import openpyxl as px
from datetime import date
from linebot import LineBotApi
from linebot.models import TextSendMessage

today=date.today()
file_json=open("info.json", "r")
info=json.load(file_json)
wb=px.load_workbook(str(today)+".xlsx")
ws=wb.worksheets[0]
x=ws.max_row
n=ws.cell(row=x, column=1).value
text="本日左クリックをした回数："+str(n)+"回"

CHANNEL_ACCESS_TOKEN=info["CHANNEL_ACCESS_TOKEN"]
line_bot_api=LineBotApi(CHANNEL_ACCESS_TOKEN)

def main():
    USER_ID=info["USER_ID"]
    messages=TextSendMessage(text)
    line_bot_api.push_message(USER_ID, messages)

main()
os.remove(str(today)+".xlsx")


