import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.utils import formatdate
import yfinance as yf
import matplotlib.pyplot as plt
import japanize_matplotlib

# --- 設定 ---
GMAIL_USER = os.environ.get("GMAIL_USER")
GMAIL_PASSWORD = os.environ.get("GMAIL_PASSWORD")
TO_EMAIL = GMAIL_USER

def test_run():
    print("=== テスト実行開始 ===")

    # 1. ダミーデータの取得とグラフ作成
    print("グラフを作成中...")
    try:
        # トヨタ(7203.T)の直近1ヶ月のデータ
        df = yf.download("7203.T", period="1mo", progress=False)
        
        plt.figure(figsize=(10, 5))
        plt.plot(df.index, df['Close'], label="Toyota")
        plt.title("テスト送信: トヨタ自動車 (直近1ヶ月)")
        plt.legend()
        plt.grid(True)
        
        # 画像保存
        chart_filename = "test_chart.png"
        plt.savefig(chart_filename)
        plt.close()
        print(f"画像保存完了: {chart_filename}")
    except Exception as e:
        print(f"グラフ作成エラー: {e}")
        return

    # 2. メール送信
    if not GMAIL_USER or not GMAIL_PASSWORD:
        print("エラー: GitHub Secrets (GMAIL_USER, GMAIL_PASSWORD) が設定されていません。")
        return

    print("メール送信準備中...")
    msg = MIMEMultipart()
    msg['Subject'] = "【テスト】GitHub Actionsからの画像付きメール"
    msg['From'] = GMAIL_USER
    msg['To'] = TO_EMAIL
    msg['Date'] = formatdate()

    # 本文
    body = "これはテスト送信です。\nチャート画像が添付されていれば成功です！\n確認したらこのファイルは削除してOKです。"
    msg.attach(MIMEText(body, 'plain'))

    # 画像添付
    if os.path.exists(chart_filename):
        with open(chart_filename, 'rb') as f:
            img_data = f.read()
            image = MIMEImage(img_data, name=os.path.basename(chart_filename))
            msg.attach(image)

    # 送信
    try:
        smtp = smtplib.SMTP('smtp.gmail.com', 587)
        smtp.starttls()
        smtp.login(GMAIL_USER, GMAIL_PASSWORD)
        smtp.sendmail(GMAIL_USER, TO_EMAIL, msg.as_string())
        smtp.close()
        print("✅ テストメール送信成功！受信トレイを確認してください。")
    except Exception as e:
        print(f"❌ メール送信失敗: {e}")

if __name__ == "__main__":
    test_run()
