import os
import smtplib
import time
import threading
from email.mime.text import MIMEText
from dotenv import load_dotenv

# .envファイルから環境変数をロード
dotenv_path = r"C:\Users\ksuzuki4\Desktop\.env"
load_dotenv(dotenv_path)
print(dotenv_path)
gmail_user = os.getenv("GMAIL_USER")
gmail_password = os.getenv("GMAIL_PASSWORD")
print(gmail_user)
if not gmail_user or not gmail_password:
    raise ValueError("GMAIL_USERまたはGMAIL_PASSWORDが設定されていません")

sender = gmail_user
recipient = "fer65s@gmail.com"
subject = "テストメール"
body = "テスト"

# MIMETextを使用してメール本文を作成
msg = MIMEText(body, "plain", "utf-8")
msg["Subject"] = subject
msg["From"] = sender
msg["To"] = recipient

server = None  # server変数を初期化

# カウントダウン終了用のフラグ
stop_countdown = False

def countdown(seconds):
    for i in range(seconds, 0, -1):
        if stop_countdown:
            break
        print(f"接続中: 残り {i} 秒...")
        time.sleep(1)

# カウントダウンスレッドを開始
countdown_thread = threading.Thread(target=countdown, args=(10,))
countdown_thread.start()

try:
    # タイムアウトを5秒に設定してSMTPサーバに接続（ポート587でSTARTTLSを使用）
    server = smtplib.SMTP("smtp.gmail.com", 587, timeout=10)
    server.ehlo()
    server.starttls()
    server.ehlo()
    # 接続完了後、カウントダウンを止めるためのフラグを設定
    stop_countdown = True
    countdown_thread.join()
    
    response = server.login(gmail_user, gmail_password)
    print("アカウント認証に成功しました:", response)
    server.sendmail(sender, [recipient], msg.as_string())
    print("メール送信に成功しました")
except Exception as e:
    print("メール送信中にエラーが発生しました:", e)
finally:
    if server:
        server.quit()