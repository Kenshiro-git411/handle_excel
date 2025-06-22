import os
import xlwings as xw
import win32com.client
import pythoncom
import psutil
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

# メールイベントフラグ
mail_event = True

# メール送信フラグ
exit_flag = False

# 操作ファイルpath
bookPath = os.environ['BOOK_PATH']
mailTemplatePath = os.path.abspath(os.path.join(BASE_DIR, "..", "src", "mail", "mail_template.txt"))

# 操作シートを定義
sheetName = "sheet1"

# 曜日の定義
d_week = {'Sun': '日', 'Mon': '月', 'Tue': '火', 'Wed': '水', 'Thu': '木', 'Fri': '金', 'Sat': '土'}

################################################################################

# メールの確認イベント
class MailEvents:
    def OnSend(self, cancel):
        global exit_flag
        print("送信イベントが発生しました")
        root = tk.Tk()
        root.withdraw()
        result = messagebox.askyesno("確認", "本当にこのメールを送信しますか")
        
        if not result:
            cancel.Value = True # 送信キャンセル
            messagebox.showinfo("キャンセル", "送信を中止しました")
        else:
            messagebox.showinfo("送信", "メールが送信されました")
            exit_flag = True
        
        root.destroy()

# outlookが起動中確認処理
def is_outlook_running():
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and 'OUTLOOK.EXE' in proc.info['name'].upper():
            return True
    return False

def mail_final_check():
    pass

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

# イベントをフック
mail_with_events = win32com.client.WithEvents(mail, MailEvents)

# 宛先
mail.To = mail_address
# 本文形式
mail.BodyFormat = 2
# メールタイトル
mail.Subject = "業務日報の連絡"
# 本文
mail.Body = mail_body

# outlookの新規メール画面を表示
mail.Display()

# イベント監視ループ
print("送信イベント監視中")
while mail_event:
    pythoncom.PumpWaitingMessages()

    if exit_flag:
        print("監視を終了します")
        break

    if not is_outlook_running():
        print("Outlookが終了されたため、監視を終了します")
        break