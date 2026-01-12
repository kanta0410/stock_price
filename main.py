import yfinance as yf
import pandas as pd
import time
import sys
import numpy as np
from tqdm import tqdm

# --- 設定 ---
# テストモード: Trueにすると最初の50銘柄だけでテスト実行します（本番はFalseにしてください）
TEST_MODE = False 

# --- 1. 全銘柄リストを取得する関数 ---
def get_all_jpx_tickers():
    print("JPX公式サイトから銘柄一覧を取得中...")
    url = "https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xls"
    
    try:
        df_list = pd.read_excel(url)
        tickers = df_list["コード"].astype(str) + ".T"
        print(f"取得完了: {len(tickers)} 銘柄が見つかりました。")
        return tickers.tolist()
    except Exception as e:
        print("リストの自動取得に失敗しました。予備リストを使用します。")
        return ["7203.T", "6758.T", "8035.T", "9984.T", "6861.T", "4063.T", "8058.T", "8001.T"]

# --- 2. 一次スクリーニング関数 (足切り) ---
def check_buffett_criteria(ticker_symbol):
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

        # 足切り条件: 粗利益率40%以上 かつ ROE 15%以上
        if gross_margin >= 0.40 and roe >= 0.15:
            return {"Ticker": ticker_symbol}
        return None
    except Exception:
        return None

# --- 3. テクニカル指標計算関数 ---
def calculate_technicals(hist):
    if len(hist) < 200:
        return {"RSI": None, "GC": False, "Trend": "データ不足"}

    # 終値
    close = hist['Close']

    # RSI (14日)
    delta = close.diff()
    gain = (delta.where(delta > 0, 0)).rolling(window=14).mean()
    loss = (-delta.where(delta < 0, 0)).rolling(window=14).mean()
    rs = gain / loss
    rsi = 100 - (100 / (1 + rs))
    current_rsi = rsi.iloc[-1]

    # 移動平均線
    sma25 = close.rolling(window=25).mean().iloc[-1]
    sma50 = close.rolling(window=50).mean().iloc[-1]
    sma75 = close.rolling(window=75).mean().iloc[-1]
    sma200 = close.rolling(window=200).mean().iloc[-1]
    current_price = close.iloc[-1]

    # ゴールデンクロス (50日線が200日線より上にある状態)
    is_gc = sma50 > sma200

    # トレンド判定 (現在値が75日移動平均より上なら上昇トレンドとみなす簡易判定)
    if current_price > sma75:
        trend = "↑上昇"
    elif current_price < sma75:
        trend = "↓下降"
    else:
        trend = "→横ばい"

    return {"RSI": current_rsi, "GC": is_gc, "Trend": trend, "SMA50": sma50, "SMA200": sma200}

# --- 4. 詳細分析関数 (スコアリング + テクニカル) ---
def get_deep_buffett_analysis(candidate_data):
    ticker_symbol = candidate_data["Ticker"]
    try:
        stock = yf.Ticker(ticker_symbol)
        info = stock.info
        income_stmt = stock.financials
        balance_sheet = stock.balance_sheet
        cashflow = stock.cashflow

        # 必要なデータが空ならスキップ
        if income_stmt.empty or balance_sheet.empty: return None

        # --- A. ファンダメンタルズ取得 ---
        revenue = info.get('totalRevenue', 0)
        gross_profit = info.get('grossProfits', 0)
        operating_income = income_stmt.loc['Operating Income'].iloc[0] if 'Operating Income' in income_stmt.index else 0
        sga = gross_profit - operating_income
        net_income = info.get('netIncomeToCommon', 0)
        long_term_debt = info.get('longTermDebt', 0)
        
        # 自社株買い (キャッシュフロー計算書の「株式の取得」を確認)
        # ※マイナス値で記録されることが多い
        buyback_value = 0
        buyback_flag = "なし"
        # 項目名のゆらぎに対応
        cf_items = ['Repurchase Of Capital Stock', 'Common Stock Repurchased', 'Purchase Of Capital Stock']
        for item in cf_items:
            if item in cashflow.index:
                # 直近1年(or 4四半期)の合計を確認
                buyback_value = cashflow.loc[item].iloc[0] # 最新期
                if buyback_value < 0: # お金が出ていっている＝自社株買いしている
                    buyback_flag = "あり"
                break

        # インサイダー保有率
        insider_pct = info.get('heldPercentInsiders', 0)

        # --- B. テクニカル分析 ---
        # 過去1年分のデータを取得
        hist = stock.history(period="1y")
        tech_data = calculate_technicals(hist)
        
        # --- C. スコアリング ---
        score = 0
        analysis_log = []

        # 1. SGA比率 (30%以下)
        sga_ratio = 1.0
        if gross_profit > 0:
            sga_ratio = sga / gross_profit
            if sga_ratio <= 0.30: score += 2; analysis_log.append("SGA◎")
            elif sga_ratio <= 0.50: score += 1

        # 2. 借金 (純利益×3年以内)
        debt_years = 0
        if net_income and net_income > 0:
            debt_years = long_term_debt / net_income
            if debt_years < 3.0: score += 1; analysis_log.append("借金少")
        
        # 3. ROE (20%超え)
        roe = info.get('returnOnEquity', 0)
        if roe > 0.20: score += 1; analysis_log.append("ROE★")

        # 4. 自社株買い加点
        if buyback_flag == "あり":
            score += 1; analysis_log.append("自社株買")

        # 5. インサイダー保有率 (高いとオーナー企業として期待)
        if insider_pct > 0.10: # 10%以上
            score += 1; analysis_log.append("役員保有")

        # --- D. 結果まとめ ---
        return {
            "Ticker": ticker_symbol,
            "Name": info.get('shortName', ticker_symbol),
            "Score": score,
            "Price": info.get('currentPrice'),
            "PER": info.get('trailingPE'),
            "PBR": info.get('priceToBook'),    # 追加
            "ROE": roe,                        # 数値のまま保持(ソート用)
            "Insider": f"{insider_pct:.1%}" if insider_pct else "-", # 追加
            "Buyback": buyback_flag,           # 追加
            "Trend": tech_data["Trend"],       # 追加
            "GC": "発生中" if tech_data["GC"] else "-", # 追加
            "RSI": f"{tech_data['RSI']:.1f}" if tech_data['RSI'] else "-", # 追加
            "Analysis": " ".join(analysis_log)
        }

    except Exception:
        return None

# --- メイン実行処理 ---
if __name__ == "__main__":
    print("=== バフェット流スクリーニング (テクニカル分析付き) ===")
    
    # 1. リスト取得
    all_tickers = get_all_jpx_tickers()
    if TEST_MODE:
        print("★テストモード: 50銘柄のみ")
        all_tickers = all_tickers[:50]

    # 2. 一次スクリーニング
    print(f"\nPhase 1: 粗利・ROEによる足切り開始 ({len(all_tickers)}銘柄)...")
    candidates = []
    for ticker in tqdm(all_tickers, ncols=80):
        result = check_buffett_criteria(ticker)
        if result: candidates.append(result)
        time.sleep(0.05)

    print(f"\nPhase 1 完了: {len(candidates)} 銘柄通過")

    # 3. 詳細分析
    if not candidates: sys.exit(0)

    print(f"\nPhase 2: 詳細分析 (PBR/インサイダー/テクニカル)...")
    final_results = []
    for candidate in tqdm(candidates, ncols=80):
        detail = get_deep_buffett_analysis(candidate)
        if detail: final_results.append(detail)
        time.sleep(0.5)

    # 4. 結果出力
    if final_results:
        df = pd.DataFrame(final_results)
        # スコア高い順 -> ROE高い順
        df = df.sort_values(by=["Score", "ROE"], ascending=[False, False])
        
        # ROEを表示用に%変換
        df["ROE"] = df["ROE"].apply(lambda x: f"{x:.1%}")

        print("\n" + "="*80)
        print("【本日の有望銘柄リスト (Top 15)】")
        print("="*80)
        
        # 表示する項目を選択
        cols = ["Ticker", "Name", "Score", "Price", "PER", "PBR", "ROE", "Insider", "Buyback", "Trend", "GC", "RSI"]
        print(df[cols].head(15).to_markdown(index=False))
        
        df.to_csv("buffett_daily_result.csv", index=False)
        print("\n全データを 'buffett_daily_result.csv' に保存しました。")
    else:
        print("候補なし")
