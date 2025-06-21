import os
import xlwings as xw
import win32com.client
import tkinter as tk
import tkinter.messagebox as messagebox
from string import Template
from dotenv import load_dotenv

# 環境変数.envの読み込み
load_dotenv()

# スクリプト自身のディレクトリを基準にする
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

################################################################################
# 定義と環境変数取得
class CustomError(Exception):
    pass

# 操作ファイルpath
bookPath = os.environ['BOOK_PATH']
mailTemplatePath = os.path.abspath(os.path.join(BASE_DIR, "..", "src", "mail", "mail_template.txt"))

# 操作シートを定義
sheetName = "sheet1"

# 曜日の定義
d_week = {'Sun': '日', 'Mon': '月', 'Tue': '火', 'Wed': '水', 'Thu': '木', 'Fri': '金', 'Sat': '土'}

################################################################################

# メールのtemplateテキストが存在するか確認する
if not os.path.exists(mailTemplatePath):
    raise FileNotFoundError(f"ファイルが見つかりません: {mailTemplatePath}")

# ファイルを開く
workBook = xw.Book(bookPath)

# 対象シートを取得
workSheet = workBook.sheets[sheetName]

# セルの値を取得
day = workSheet.range("B1").value
name = workSheet.range("B2").value
contents = workSheet.range("B3").value
messages = workSheet.range("B4").value
remarks = workSheet.range("B5").value
status = workSheet.range("B6").value
mail_address = workSheet.range("B7").value

# statusチェック
if not status == "OK":
    tk.Tk().withdraw()
    messagebox.showerror('日報ファイルエラー', '日報ファイルの状態がOKではありません。OKにしてからやり直してください。')
    raise CustomError("状態の値がOKではありません")

# 日付形式 -> 年月日（曜日）
key = day.strftime('%a')
w = d_week[key]
today = day.strftime('%Y年%m月%d日') + f'({w})'

context = {
    "today": today,
    "name": name,
    "contents": contents,
    "messages": messages,
    "remarks": remarks,
}

with open(mailTemplatePath, "r", encoding="utf-8") as f:
    templateSentence = f.read()

template = Template(templateSentence)
mail_body = template.substitute(context) # templateの該当箇所にcontext内容を埋め込む

# ファイルの保存
workBook.save()

# ファイルを閉じる
workBook.close()

outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)

# 宛先
mail.To = mail_address
# 本文形式
mail.BodyFormat = 2
# 本文
mail.Body = mail_body

mail.Display()