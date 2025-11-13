#!/usr/bin/env python3

import pandas as pd
import yfinance as yf
from datetime import timedelta

# ==========================================================
# CONFIG
# ==========================================================
STOCKANALYSIS_URL = "https://stockanalysis.com/list/biggest-companies/"
TOP_N = 1000                      # target number of tickers

YFINANCE_PERIOD = "1y"
YFINANCE_BATCH_SIZE = 200         # tickers per yfinance request

HORIZONS_DAYS = {
    "1w": 7,
    "1m": 30,
    "3m": 90,
    "6m": 180,
}

MASTER_CSV = "top1000_returns_master.csv"
EXCEL_REPORT = "top1000_return_report.xlsx"


# ==========================================================
# SYMBOL NORMALIZATION FOR YAHOO
# ==========================================================
def to_yahoo_symbol(sym: str) -> str:
    """
    Convert StockAnalysis / 'dot' style tickers to Yahoo style.
    Examples: BRK.B -> BRK-B, PBR.A -> PBR-A, HEI.A -> HEI-A
    """
    sym = sym.strip().upper()

    overrides = {
        "BRK.B": "BRK-B",
        "BRK.A": "BRK-A",
        "BF.B": "BF-B",
        "HEI.A": "HEI-A",
        "PBR.A": "PBR-A",
    }
    if sym in overrides:
        return overrides[sym]

    # Generic rule: replace '.' with '-'
    return sym.replace(".", "-")


# ==========================================================
# 1) SCRAPE ~TOP 1,000 TICKERS FROM STOCKANALYSIS (MULTI-PAGE)
# ==========================================================
def fetch_top_tickers_from_stockanalysis(top_n: int = TOP_N, max_pages: int = 5):
    """
    Scrape StockAnalysis 'Biggest Companies' list.
    Walks pages ?page=1,2,... until we have at least top_n tickers
    or we run out of pages.
    """
    all_syms = []
    page = 1

    while len(all_syms) < top_n and page <= max_pages:
        url = STOCKANALYSIS_URL if page == 1 else f"{STOCKANALYSIS_URL}?page={page}"
        print(f"[STEP] Fetching table from StockAnalysis page {page}: {url}")

        try:
            tables = pd.read_html(url)
        except Exception as e:
            print(f"[WARN] Failed to read HTML table on page {page}: {e}")
            break

        if not tables:
            print(f"[WARN] No tables found on page {page}.")
            break

        df = tables[0]

        # Find symbol column
        symbol_col = None
        for c in df.columns:
            name = str(c).lower()
            if "symbol" in name or "ticker" in name:
                symbol_col = c
                break
        if symbol_col is None:
            if len(df.columns) >= 2:
                symbol_col = df.columns[1]
                print(f"[WARN] Symbol column not labelled clearly. Using '{symbol_col}' on page {page}.")
            else:
                raise ValueError(f"Could not identify symbol column in scraped table on page {page}.")

        syms = (
            df[symbol_col]
            .dropna()
            .astype(str)
            .str.strip()
            .str.upper()
            .tolist()
        )

        all_syms.extend(syms)
        print(f"[INFO] Collected {len(all_syms)} raw symbols after page {page}.")

        page += 1

    # De-duplicate while preserving order
    seen = set()
    clean = []
    for s in all_syms:
        if s and s not in seen:
            seen.add(s)
            clean.append(s)

    tickers = clean[:top_n]
    print(f"[INFO] Got {len(tickers)} unique tickers from StockAnalysis (target {top_n}).")
    print(f"[INFO] First 10 tickers: {tickers[:10]}")
    return tickers


# ==========================================================
# 2) DOWNLOAD PRICES FROM YAHOO (USING NORMALIZED SYMBOLS)
# ==========================================================
def download_price_history_batched(tickers, period=YFINANCE_PERIOD, batch_size=YFINANCE_BATCH_SIZE):
    """
    tickers: list of original symbols (e.g. BRK.B)
    We normalize to Yahoo symbols internally (e.g. BRK-B)
    but keep columns / index as the original tickers.
    """
    print(f"[STEP] Downloading daily prices for {len(tickers)} tickers from Yahoo Finance...")
    all_closes = {}

    for i in range(0, len(tickers), batch_size):
        batch_orig = tickers[i:i + batch_size]
        batch_yahoo = [to_yahoo_symbol(t) for t in batch_orig]

        print(f"[INFO]  Batch {i // batch_size + 1}: {len(batch_orig)} tickers "
              f"({', '.join(batch_orig[:5])}...)")

        data = yf.download(
            tickers=batch_yahoo,
            period=period,
            interval="1d",
            auto_adjust=True,
            group_by="ticker",
            progress=False,
            threads=True,
        )

        if isinstance(data.columns, pd.MultiIndex):
            # multiple tickers
            for orig, ysym in zip(batch_orig, batch_yahoo):
                try:
                    if (ysym, "Close") in data.columns:
                        series = data[(ysym, "Close")].dropna()
                    elif (ysym, "Adj Close") in data.columns:
                        series = data[(ysym, "Adj Close")].dropna()
                    else:
                        continue
                    if not series.empty:
                        all_closes[orig] = series
                except KeyError:
                    continue
        else:
            # single ticker
            orig = batch_orig[0]
            if "Close" in data.columns:
                series = data["Close"].dropna()
            elif "Adj Close" in data.columns:
                series = data["Adj Close"].dropna()
            else:
                series = None

            if series is not None and not series.empty:
                all_closes[orig] = series

    if not all_closes:
        raise RuntimeError("No price data downloaded. Check tickers or internet connection.")

    prices = pd.DataFrame(all_closes).sort_index()
    print(f"[INFO] Price history shape: {prices.shape} (rows = days, cols = tickers)")
    return prices


# ==========================================================
# 3) COMPUTE HORIZON RETURNS
# ==========================================================
def compute_horizon_returns(prices, horizons_days):
    """
    Returns DataFrame:
      index  = ticker (original symbols)
      cols   = '1w', '1m', '3m', '6m'
      values = % return (e.g. 12.34 means +12.34%)
    """
    print("[STEP] Computing horizon returns...")
    prices = prices.sort_index()
    last_date = prices.index.max()
    last_prices = prices.loc[last_date]

    result = {}

    for label, days in horizons_days.items():
        anchor_date = last_date - timedelta(days=days)
        eligible_dates = prices.index[prices.index <= anchor_date]

        if len(eligible_dates) == 0:
            print(f"[WARN] Not enough history for {label} ({days} days). Filling NaNs.")
            result[label] = pd.Series(index=prices.columns, dtype=float)
            continue

        anchor_prices = prices.loc[eligible_dates.max()]
        horizon_ret = (last_prices / anchor_prices - 1.0) * 100.0
        result[label] = horizon_ret

    returns_df = pd.DataFrame(result)
    returns_df.index.name = "ticker"
    return returns_df


# ==========================================================
# 4) RANKED VIEWS + SAVING (NICER DISPLAY + README)
# ==========================================================
def build_ranked_views(returns_df):
    ranked_views = {}
    for col in returns_df.columns:
        ranked_views[col] = returns_df.sort_values(col, ascending=False)
    return ranked_views


def save_report(returns_df, ranked_views, master_csv=MASTER_CSV, excel_path=EXCEL_REPORT):
    print("[STEP] Saving master CSV and Excel report...")

    # ---------- CSV (keep values as % numbers, not decimal) ----------
    returns_df_csv = returns_df.rename(
        columns={
            "1w": "1w_return_pct",
            "1m": "1m_return_pct",
            "3m": "3m_return_pct",
            "6m": "6m_return_pct",
        }
    )
    returns_df_csv.to_csv(master_csv, float_format="%.4f")
    print(f"[OUT] Master CSV saved to {master_csv}")

    # ---------- Excel ----------
    # For Excel, convert to decimal (0.1234) and format as %
    def display_df(df):
        df_disp = df.copy() / 100.0
        df_disp.columns = ["1W %", "1M %", "3M %", "6M %"]
        return df_disp

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        # 1) Master sheet
        display_df(returns_df).to_excel(writer, sheet_name="Master")

        # 2) Horizon sheets
        for horizon, df_view in ranked_views.items():
            sheet_name = f"Top_{horizon}"
            display_df(df_view).to_excel(writer, sheet_name=sheet_name)

        # 3) ReadMe sheet explaining each tab
        readme_data = {
            "Sheet": [
                "Master",
                "Top_1w",
                "Top_1m",
                "Top_3m",
                "Top_6m",
            ],
            "Description": [
                "All tickers with 1-week, 1-month, 3-month, and 6-month total returns, expressed as price-change percentages. One row per ticker.",
                "Same table as Master, sorted by 1-week return (1W %) from highest to lowest. Top recent weekly performers at the top.",
                "Same table as Master, sorted by 1-month return (1M %) from highest to lowest.",
                "Same table as Master, sorted by 3-month return (3M %) from highest to lowest.",
                "Same table as Master, sorted by 6-month return (6M %) from highest to lowest.",
            ],
        }
        pd.DataFrame(readme_data).to_excel(writer, sheet_name="ReadMe", index=False)

        # Apply Excel % formatting to numeric cells on all sheets except ReadMe
        wb = writer.book
        for ws in wb.worksheets:
            if ws.title == "ReadMe":
                continue
            # Data starts at row 2 (row 1 = header), col 2 (col 1 = index "ticker")
            for row in ws.iter_rows(min_row=2, min_col=2):
                for cell in row:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "0.00%"

    print(f"[OUT] Excel report saved to {excel_path}")


def print_top_snippets(returns_df, top_n=20):
    print("\n================ TOP MOVERS ================")
    for col in returns_df.columns:
        print(f"\n--- Top {top_n} by {col} return (percent points) ---")
        subset = returns_df.sort_values(col, ascending=False)
        print(subset.head(top_n).round(2))


# ==========================================================
# MAIN
# ==========================================================
def main():
    print("[START] Script started.")

    # 1) Scrape up to 1,000 tickers (multi-page)
    tickers = fetch_top_tickers_from_stockanalysis(TOP_N)
    print(f"[INFO] Using {len(tickers)} tickers.")

    # 2) Download price history
    prices = download_price_history_batched(tickers, period=YFINANCE_PERIOD)

    # 3) Compute returns
    returns_df = compute_horizon_returns(prices, HORIZONS_DAYS)
    returns_df = returns_df.dropna(how="all")

    # 4) Ranked views
    ranked_views = build_ranked_views(returns_df)

    # 5) Console preview
    print_top_snippets(returns_df, top_n=20)

    # 6) Save report
    save_report(returns_df, ranked_views, MASTER_CSV, EXCEL_REPORT)

    print("[DONE] All finished.")


# ------------------------------------------------------------------
# 7) OPTIONAL: Copy outputs to Drive (Colab), Windows Google Drive,
#    AND your local project folder.
# ------------------------------------------------------------------
try:
    import os, shutil

    # Local output files created by this script
    files_to_copy = []
    if os.path.exists(MASTER_CSV):
        files_to_copy.append(MASTER_CSV)
    if os.path.exists(EXCEL_REPORT):
        files_to_copy.append(EXCEL_REPORT)

    # --- A) Copy into your LOCAL project folder (Windows PC) ---
    local_project_path = r"C:\Users\Tommy\top_stock_returns"
    if os.path.isdir(local_project_path):
        for fname in files_to_copy:
            dst = os.path.join(local_project_path, os.path.basename(fname))
            try:
                shutil.copy2(fname, dst)
                print(f"[COPY] {fname} -> {dst}")
            except Exception as e:
                print(f"[WARN] Could not copy {fname} to local project folder: {e}")

    # --- B) Copy into Windows Google Drive folder ---
    windows_drive_path = r"G:\My Drive\Top Stocks Output"
    if os.path.isdir(windows_drive_path):
        for fname in files_to_copy:
            dst = os.path.join(windows_drive_path, os.path.basename(fname))
            try:
                shutil.copy2(fname, dst)
                print(f"[COPY] {fname} -> {dst}")
            except Exception as e:
                print(f"[WARN] Could not copy {fname} to Windows Drive: {e}")

    # --- C) Copy into Colab Google Drive folder ---
    colab_drive_path = "/content/drive/MyDrive/Top Stocks Output"
    if os.path.isdir(colab_drive_path):
        for fname in files_to_copy:
            dst = os.path.join(colab_drive_path, os.path.basename(fname))
            try:
                shutil.copy2(fname, dst)
                print(f"[COPY] {fname} -> {dst}")
            except Exception as e:
                print(f"[WARN] Could not copy {fname} to Colab Drive: {e}")

except Exception as outer_err:
    print(f"[WARN] Output copy step failed: {outer_err}")


if __name__ == "__main__":
    import traceback
    try:
        main()
    except Exception as e:
        print("\n[ERROR] An exception occurred:")
        print(e)
        traceback.print_exc()
