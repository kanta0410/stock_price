import os
import smtplib
import math
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.utils import formatdate
import yfinance as yf
import pandas as pd
import time
from tqdm import tqdm
import matplotlib.pyplot as plt
import japanize_matplotlib

# --- 設定: 環境変数から取得 ---
GMAIL_USER = os.environ.get("GMAIL_USER")
GMAIL_PASSWORD = os.environ.get("GMAIL_PASSWORD")
TO_EMAIL = GMAIL_USER 

# ---------------------------------------------------------
# 関数1: 全銘柄リスト取得
# ---------------------------------------------------------
def get_all_jpx_tickers():
    print("JPX公式サイトから銘柄一覧を取得中...")
    url = "https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xls"
    try:
        df_list = pd.read_excel(url)
        tickers = df_list["コード"].astype(str) + ".T"
        print(f"取得完了: {len(tickers)} 銘柄が見つかりました。")
        return tickers.tolist()
    except Exception as e:
        print(f"リスト取得エラー: {e}")
        return ["7203.T", "6758.T", "8035.T", "9984.T", "6861.T"]

# ---------------------------------------------------------
# 関数2: 一次スクリーニング (粗利率 & ROE)
# ---------------------------------------------------------
def check_basic_criteria(ticker_symbol):
    try:
        stock = yf.Ticker(ticker_symbol)
        info = stock.info
        
        if 'totalRevenue' not in info or 'grossProfits' not in info: return None
        revenue = info.get('totalRevenue')
        gross_profit = info.get('grossProfits')
        
        if not revenue or revenue == 0: return None
        if not gross_profit: return None

        gross_margin = gross_profit / revenue
        roe = info.get('returnOnEquity', 0)

        # 判定: 粗利益率40%以上 かつ ROE15%以上
        if gross_margin >= 0.40 and roe >= 0.15:
            return {
                "Ticker": ticker_symbol,
                "Name": info.get('shortName', ticker_symbol),
                "Price": info.get('currentPrice'),
                "GrossMargin": gross_margin,
                "ROE": roe
            }
        return None
    except:
        return None

# ---------------------------------------------------------
# 関数3: 詳細分析 (バフェット・スコア算出)
# ---------------------------------------------------------
def get_deep_analysis(ticker_data):
    ticker = ticker_data["Ticker"]
    try:
        stock = yf.Ticker(ticker)
        info = stock.info
        income = stock.financials
        balance = stock.balance_sheet
        cashflow = stock.cashflow

        if income.empty or balance.empty or cashflow.empty: return None

        gross_profit = info.get('grossProfits', 0)
        op_income = income.loc['Operating Income'].iloc[0] if 'Operating Income' in income.index else 0
        sga = gross_profit - op_income
        
        interest_exp = abs(income.loc['Interest Expense'].iloc[0]) if 'Interest Expense' in income.index else 0
        net_income = info.get('netIncomeToCommon', 0)
        long_term_debt = info.get('longTermDebt', 0)
        
        retained_now = balance.loc['Retained Earnings'].iloc[0] if 'Retained Earnings' in balance.index else 0
        retained_prev = balance.loc['Retained Earnings'].iloc[1] if 'Retained Earnings' in balance.index and len(balance.columns) > 1 else 0
        
        capex = abs(cashflow.loc['Capital Expenditure'].iloc[0]) if 'Capital Expenditure' in cashflow.index else 0

        score = 0
        log = []

        # 1. SGA比率
        sga_ratio = sga / gross_profit if gross_profit > 0 else 1.0
        if sga_ratio <= 0.30: score += 2; log.append("◎SGA低")
        elif sga_ratio <= 0.50: score += 1

        # 2. 利払負担
        if op_income > 0 and (interest_exp / op_income) < 0.15: score += 1

        # 3. 借金
        debt_years = long_term_debt / net_income if net_income > 0 else 99
        if debt_years < 3.0: score += 1

        # 4. 設備投資
        capex_ratio = capex / net_income if net_income > 0 else 99
        if capex_ratio < 0.25: score += 2; log.append("◎CapEx少")
        elif capex_ratio < 0.50: score += 1

        # 5. 内部留保増加
        if retained_now > retained_prev: score += 1

        # 6. ROE > 20%
        if ticker_data["ROE"] > 0.20: score += 1; log.append("★ROE高")

        return {
            "Ticker": ticker,
            "Name": ticker_data["Name"],
            "Buffett_Score": score,
            "Price": ticker_data["Price"],
            "Analysis": " ".join(log)
        }
    except:
        return None

# ---------------------------------------------------------
# 関数4: 最終分析 (テクニカル & オーナーシップ)
# ---------------------------------------------------------
def get_ultimate_data(base_data):
    ticker = base_data["Ticker"]
    try:
        stock = yf.Ticker(ticker)
        info = stock.info
        
        insider_pct = info.get('heldPercentInsiders', 0) * 100
        is_buyback = "-"
        try:
            bs = stock.balance_sheet
            if 'Ordinary Shares Number' in bs.index and len(bs.columns) > 1:
                shares_now = bs.loc['Ordinary Shares Number'].iloc[0]
                shares_prev = bs.loc['Ordinary Shares Number'].iloc[1]
                if shares_now < shares_prev * 0.99: is_buyback = "★実施"
        except: pass

        hist = stock.history(period="6mo")
        if len(hist) < 75: return None
        
        close = hist['Close']
        ma5 = close.rolling(5).mean().iloc[-1]
        ma25 = close.rolling(25).mean().iloc[-1]
        ma75 = close.rolling(75).mean().iloc[-1]
        
        trend = "レンジ/下降"
        if close.iloc[-1] > ma25 and ma25 > ma75: trend = "上昇"
        if ma5 > ma25 and ma25 > ma75: trend = "★パーフェクト"

        delta = close.diff()
        gain = delta.where(delta > 0, 0).rolling(14).mean().iloc[-1]
        loss = -delta.where(delta < 0, 0).rolling(14).mean().iloc[-1]
        rs = gain / loss if loss != 0 else 0
        rsi = 100 - (100 / (1 + rs))

        return {
            "社名": base_data["Name"],
            "コード": ticker,
            "現在値": base_data["Price"],
            "スコア": base_data["Buffett_Score"],
            "判定メモ": base_data["Analysis"],
            "インサイダー": f"{insider_pct:.1f}%",
            "自社株買い": is_buyback,
            "トレンド": trend,
            "RSI": f"{rsi:.0f}"
        }
    except:
        return None

# ---------------------------------------------------------
# 関数5: チャート画像生成 (新規追加)
# ---------------------------------------------------------
def generate_charts(results_list, filename="chart_summary.png"):
    print("チャート画像を生成中...")
    if not results_list: return None

    codes = [d["コード"] for d in results_list]
    num_plots = len(codes)
    
    # レイアウト計算 (3列固定)
    cols = 3
    rows = math.ceil(num_plots / cols)
    
    # グラフサイズ設定
    fig, axes = plt.subplots(rows, cols, figsize=(20, 5 * rows))
    # 1行の場合や1つだけの場合のaxesの型を統一
    if num_plots == 1: axes = [axes]
    else: axes = axes.flatten()

    for i, code in enumerate(codes):
        try:
            # データ取得 (1年分)
            df = yf.download(code, period="1y", progress=False)
            ax = axes[i]

            if len(df) == 0:
                ax.text(0.5, 0.5, "No Data", ha='center')
                continue

            # 移動平均線
            df['MA5'] = df['Close'].rolling(window=5).mean()
            df['MA25'] = df['Close'].rolling(window=25).mean()
            df['MA75'] = df['Close'].rolling(window=75).mean()

            # プロット
            ax.plot(df.index, df['Close'], label='株価', color='#333333', linewidth=1.5, alpha=0.7)
            ax.plot(df.index, df['MA5'], label='5日', color='#ff7f0e', linewidth=1.5)
            ax.plot(df.index, df['MA25'], label='25日', color='#1f77b4', linewidth=1.5)
            ax.plot(df.index, df['MA75'], label='75日', color='#2ca02c', linewidth=1.5, linestyle='--')

            # タイトルと装飾
            stock_name = next((item["社名"] for item in results_list if item["コード"] == code), code)
            ax.set_title(f"{stock_name} ({code})", fontsize=14, fontweight='bold')
            ax.grid(True, alpha=0.3)
            ax.legend(loc='upper left', fontsize='small')
        
        except Exception as e:
            print(f"Plot Error {code}: {e}")

    # 余った枠を非表示
    if num_plots < len(axes):
        for j in range(num_plots, len(axes)):
            fig.delaxes(axes[j])

    plt.tight_layout()
    plt.savefig(filename) # 画像として保存
    plt.close() # メモリ開放
    print(f"チャート画像を保存しました: {filename}")
    return filename

# ---------------------------------------------------------
# メール送信関数 (画像添付対応版)
# ---------------------------------------------------------
def send_email_with_image(subject, body, image_path=None):
    if not GMAIL_USER or not GMAIL_PASSWORD:
        print("メール設定なし: 送信スキップ")
        return

    # マルチパート形式（テキスト + 画像）にする
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = GMAIL_USER
    msg['To'] = TO_EMAIL
    msg['Date'] = formatdate()

    # 本文を添付
    msg.attach(MIMEText(body, 'plain'))

    # 画像を添付
    if image_path and os.path.exists(image_path):
        with open(image_path, 'rb') as f:
            img_data = f.read()
            image = MIMEImage(img_data, name=os.path.basename(image_path))
            msg.attach(image)
        print("画像をメールに添付しました。")

    try:
        smtp = smtplib.SMTP('smtp.gmail.com', 587)
        smtp.starttls()
        smtp.login(GMAIL_USER, GMAIL_PASSWORD)
        smtp.sendmail(GMAIL_USER, TO_EMAIL, msg.as_string())
        smtp.close()
        print("メール送信完了！")
    except Exception as e:
        print(f"メール送信エラー: {e}")

# ---------------------------------------------------------
# メイン処理
# ---------------------------------------------------------
if __name__ == "__main__":
    print("=== 全銘柄スクリーニング開始 ===")
    
    # 1. 全銘柄リスト取得
    all_tickers = get_all_jpx_tickers()
    
    # ★テスト用 (最初は数を絞って試すならコメントアウトを外す)
    # all_tickers = all_tickers[:50]

    # 2. 一次スクリーニング
    print(f"\nStep 1: 財務基準 (粗利40%, ROE15%) で絞り込み中...")
    first_pass = []
    for t in tqdm(all_tickers):
        res = check_basic_criteria(t)
        if res: first_pass.append(res)
        time.sleep(0.05)
    
    print(f"→ 一次通過: {len(first_pass)} 銘柄")

    if not first_pass:
        send_email_with_image("【株分析】該当なし", "本日の基準を満たす銘柄はありませんでした。")
        exit()

    # 3. 詳細スコアリング
    print(f"\nStep 2: バフェット・スコア算出中...")
    second_pass = []
    for data in tqdm(first_pass):
        res = get_deep_analysis(data)
        if res: second_pass.append(res)
        time.sleep(0.1)

    # 上位15社に絞る (チャートが見づらくなるため)
    df_scores = pd.DataFrame(second_pass)
    df_scores = df_scores.sort_values("Buffett_Score", ascending=False).head(15)
    top_candidates = df_scores.to_dict('records')

    # 4. 最終分析
    print(f"\nStep 3: 上位{len(top_candidates)}銘柄の最終チェック...")
    final_results = []
    for data in tqdm(top_candidates):
        res = get_ultimate_data(data)
        if res: final_results.append(res)
    
    # 5. チャート生成とメール送信
    if final_results:
        # チャート作成
        chart_file = generate_charts(final_results)

        df_final = pd.DataFrame(final_results)
        table_str = df_final.to_markdown(index=False)
        
        mail_body = (
            f"おはようございます。本日のスクリーニング結果です。\n\n"
            f"【バフェット流 有望銘柄ピックアップ】\n"
            f"{table_str}\n\n"
            f"▼ 添付ファイルに日足チャート画像をつけました。\n"
            f"移動平均線: 5日(橙), 25日(青), 75日(緑)\n\n"
            f"※GitHub Actionsから自動送信"
        )
        
        print("\n=== 最終結果 ===")
        print(table_str)
        send_email_with_image(f"【厳選】本日の有望株レポート ({len(final_results)}銘柄)", mail_body, chart_file)
    else:
        print("詳細分析の結果、残った銘柄はありませんでした。")
