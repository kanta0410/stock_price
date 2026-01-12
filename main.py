import yfinance as yf
import pandas as pd
import time
import sys
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from tqdm import tqdm

# --- 設定 ---
TEST_MODE = False 

# --- メール送信関数 ---
def send_email(df_results, csv_filename):
    # GitHub Secretsから情報を取得
    gmail_user = os.environ.get("MAIL_USERNAME")
    gmail_password = os.environ.get("MAIL_PASSWORD")
    to_email = os.environ.get("MAIL_TO")

    if not gmail_user or not gmail_password or not to_email:
        print("★メール設定が見つからないため、メール送信をスキップします。")
        return

    print("メール送信の準備中...")

    # メールの作成
    msg = MIMEMultipart()
    msg['Subject'] = f"【株価分析】本日の有望銘柄 Top {len(df_results)}"
    msg['From'] = gmail_user
    msg['To'] = to_email

    # 本文（HTML形式で見やすくする）
    # 必要な列だけ抽出
    display_cols = ["Ticker", "Name", "Score", "Price", "PER", "PBR", "ROE", "Insider", "Buyback", "Trend", "GC", "RSI"]
    html_table = df_results[display_cols].to_html(index=False, border=1)
    
    body = f"""
    <html>
      <body>
        <h2>本日のバフェット流スクリーニング結果</h2>
        <p>スクリーニングが完了しました。上位の銘柄をお知らせします。</p>
        {html_table}
        <p>※全データは添付のCSVをご確認ください。</p>
      </body>
    </html>
    """
    msg.attach(MIMEText(body, 'html'))

    # CSVファイルの添付
    try:
        with open(csv_filename, "rb") as f:
            part = MIMEApplication(f.read(), Name=csv_filename)
            part['Content-Disposition'] = f'attachment; filename="{csv_filename}"'
            msg.attach(part)
    except Exception as e:
        print(f"添付ファイルのエラー: {e}")

    # Gmailサーバー経由で送信
    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(gmail_user, gmail_password)
        server.send_message(msg)
        server.quit()
        print(f"メールを送信しました！ 宛先: {to_email}")
    except Exception as e:
        print(f"メール送信に失敗しました: {e}")

# --- 1. 全銘柄リストを取得する関数 ---
def get_all_jpx_tickers():
    print("JPX公式サイトから銘柄一覧を取得中...")
    url = "https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xls"
    try:
        df_list = pd.read_excel(url)
        tickers = df_list["コード"].astype(str) + ".T"
        print(f"取得完了: {len(tickers)} 銘柄が見つかりました。")
        return tickers.tolist()
    except Exception:
        print("リスト取得失敗。予備リストを使用。")
        return ["7203.T", "6758.T", "8035.T", "9984.T", "6861.T"]

# --- 2. 一次スクリーニング ---
def check_buffett_criteria(ticker_symbol):
    try:
        stock = yf.Ticker(ticker_symbol)
        info = stock.info
        if 'totalRevenue' not in info or 'grossProfits' not in info: return None
        
        revenue = info.get('totalRevenue')
        gross_profit = info.get('grossProfits')
        if not revenue or not gross_profit: return None

        gross_margin = gross_profit / revenue
        roe = info.get('returnOnEquity', 0)

        # 粗利益率40%以上 かつ ROE 15%以上
        if gross_margin >= 0.40 and roe >= 0.15:
            return {"Ticker": ticker_symbol}
        return None
    except Exception:
        return None

# --- 3. テクニカル指標 ---
def calculate_technicals(hist):
    if len(hist) < 200:
        return {"RSI": None, "GC": False, "Trend": "-"}

    close = hist['Close']
    delta = close.diff()
    gain = (delta.where(delta > 0, 0)).rolling(window=14).mean()
    loss = (-delta.where(delta < 0, 0)).rolling(window=14).mean()
    rs = gain / loss
    rsi = 100 - (100 / (1 + rs))
    
    sma50 = close.rolling(window=50).mean().iloc[-1]
    sma75 = close.rolling(window=75).mean().iloc[-1]
    sma200 = close.rolling(window=200).mean().iloc[-1]
    current_price = close.iloc[-1]

    is_gc = sma50 > sma200
    if current_price > sma75: trend = "↑上昇"
    elif current_price < sma75: trend = "↓下降"
    else: trend = "→横ばい"

    return {"RSI": rsi.iloc[-1], "GC": is_gc, "Trend": trend}

# --- 4. 詳細分析 ---
def get_deep_buffett_analysis(candidate_data):
    ticker_symbol = candidate_data["Ticker"]
    try:
        stock = yf.Ticker(ticker_symbol)
        info = stock.info
        income_stmt = stock.financials
        balance_sheet = stock.balance_sheet
        cashflow = stock.cashflow

        if income_stmt.empty or balance_sheet.empty: return None

        # データ取得
        revenue = info.get('totalRevenue', 0)
        gross_profit = info.get('grossProfits', 0)
        operating_income = income_stmt.loc['Operating Income'].iloc[0] if 'Operating Income' in income_stmt.index else 0
        sga = gross_profit - operating_income
        net_income = info.get('netIncomeToCommon', 0)
        long_term_debt = info.get('longTermDebt', 0)
        insider_pct = info.get('heldPercentInsiders', 0)

        # 自社株買い判定
        buyback_flag = "なし"
        cf_items = ['Repurchase Of Capital Stock', 'Common Stock Repurchased', 'Purchase Of Capital Stock']
        for item in cf_items:
            if item in cashflow.index:
                if cashflow.loc[item].iloc[0] < 0:
                    buyback_flag = "あり"
                break

        # テクニカル
        hist = stock.history(period="1y")
        tech = calculate_technicals(hist)
        
        # スコアリング
        score = 0
        analysis_log = []

        if gross_profit > 0 and (sga / gross_profit) <= 0.30: score += 2; analysis_log.append("SGA◎")
        if net_income > 0 and (long_term_debt / net_income) < 3.0: score += 1; analysis_log.append("借金少")
        
        roe = info.get('returnOnEquity', 0)
        if roe > 0.20: score += 1; analysis_log.append("ROE★")
        if buyback_flag == "あり": score += 1; analysis_log.append("自社株買")
        if insider_pct > 0.10: score += 1; analysis_log.append("役員保有")

        return {
            "Ticker": ticker_symbol,
            "Name": info.get('shortName', ticker_symbol),
            "Score": score,
            "Price": info.get('currentPrice'),
            "PER": info.get('trailingPE'),
            "PBR": info.get('priceToBook'),
            "ROE": roe,
            "Insider": f"{insider_pct:.1%}" if insider_pct else "-",
            "Buyback": buyback_flag,
            "Trend": tech["Trend"],
            "GC": "発生中" if tech["GC"] else "-",
            "RSI": f"{tech['RSI']:.1f}" if tech['RSI'] else "-",
            "Analysis": " ".join(analysis_log)
        }
    except Exception:
        return None

# --- メイン実行 ---
if __name__ == "__main__":
    print("=== バフェット流スクリーニング (メール送信機能付き) ===")
    
    all_tickers = get_all_jpx_tickers()
    if TEST_MODE: all_tickers = all_tickers[:50]

    # Phase 1
    print(f"\nPhase 1: 足切りスクリーニング ({len(all_tickers)}銘柄)...")
    candidates = []
    for ticker in tqdm(all_tickers, ncols=80):
        res = check_buffett_criteria(ticker)
        if res: candidates.append(res)
        time.sleep(0.05)

    if not candidates: sys.exit(0)

    # Phase 2
    print(f"\nPhase 2: 詳細分析 & テクニカル計算...")
    final_results = []
    for cand in tqdm(candidates, ncols=80):
        det = get_deep_buffett_analysis(cand)
        if det: final_results.append(det)
        time.sleep(0.5)

    # 結果処理
    if final_results:
        df = pd.DataFrame(final_results)
        df = df.sort_values(by=["Score", "ROE"], ascending=[False, False])
        
        # 数値調整
        df_display = df.copy()
        df_display["ROE"] = df_display["ROE"].apply(lambda x: f"{x:.1%}")
        
        print("\n【Top 15 銘柄】")
        print(df_display.head(15).to_markdown(index=False))
        
        # CSV保存
        csv_file = "buffett_daily_result.csv"
        df_display.to_csv(csv_file, index=False)
        print(f"\nCSV保存完了: {csv_file}")

        # ★メール送信実行
        send_email(df_display.head(15), csv_file)
    else:
        print("候補なし")
