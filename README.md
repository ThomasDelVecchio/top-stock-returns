# Top Stocks Trend Analysis

This project analyzes ~1,000 large-cap stocks using multi-horizon momentum,
volatility-adjusted strength, moving-average trend health, and simple momentum/pullback signals.

The output is a clean Excel report designed to be readable, sortable, and fast to interpret.

---

## 🚀 Features

- Scrapes ~1,000 tickers from StockAnalysis.com  
- Downloads 1-year historical prices via Yahoo Finance  
- Computes 1-week, 1-month, 3-month, and 6-month returns  
- Calculates percentile trend scores  
- Computes volatility-adjusted strength  
- Detects accelerating momentum  
- Detects pullbacks in strong uptrends  
- Consolidates everything into a clean Excel workbook

---

## 📊 Excel Output Structure

### **1. Master**
- The full stock universe  
- Columns: 1W, 1M, 3M, 6M returns  
- You can sort directly in Excel to get "Top 1W", "Top 1M", etc.  
- Most users spend 90%+ of their time on this tab

### **2. Top_Trend_10**
Top 10 stocks with the strongest and most consistent trend across all horizons.  
Includes:

- Trend Score  
- Risk-Adjusted Strength  
- Risk-Adjusted Rank  
- Above 50-day / Above 200-day  
- Momentum accelerating?  
- Recent pullback?

### **3. Top_Sector_Leaders**
Top 3 leaders in each sector, ranked by trend strength.

### **4. Top_Final_10**
The final recommended list using the weighted scoring system.  
Represents the highest-conviction ideas based on all available metrics.

---

## 📦 Running the Script

```bash
python top_return_ranks.py
```

Outputs:

- top1000_returns_master.csv  
- top1000_return_report.xlsx  

These files are also auto-copied to your Windows folder and Google Drive (if present).

---

## 📁 Repository Structure

```
README.md
top_return_ranks.py
Sample_Report.xlsx
Sample_Master.csv
Trend_Analysis_CheatSheet.docx
Math_Appendix.md
```

---

# 📘 Mathematical Appendix
A plain-text, GitHub-safe explanation of every column and formula used in the script.

---

## 1. Horizon Returns (1W, 1M, 3M, 6M)

Horizon return is computed as:

```
return_H = ((price_end / price_start) - 1) * 100
```

Where:
- price_end = most recent adjusted close  
- price_start = adjusted close at least H days prior  
- H ∈ {7, 30, 90, 180}

---

## 2. Percentile Trend Scores

Each horizon (1W, 1M, 3M, 6M) is ranked relative to all stocks.

Percentile score:

```
percentile_score = 1 - ((rank - 1) / (n - 1))
```

Where rank = 1 is best.

---

## 3. Composite Trend Score

Average of all four percentile scores:

```
composite_score = (score_1w + score_1m + score_3m + score_6m) / 4
```

Range: 0 to 1.

---

## 4. Daily Volatility (6-month)

Daily return:

```
daily_return_t = (P_t - P_(t-1)) / P_(t-1)
```

Volatility:

```
vol_6m = standard_deviation(daily_return_t)
```

---

## 5. Risk-Adjusted Strength

Convert 6M return to fraction:

```
R_6m = six_month_return / 100
```

Then normalize by volatility:

```
risk_adj_score = R_6m / vol_6m
```

If vol_6m = 0 → NA.

---

## 6. Risk-Adjusted Percentile Rank

```
risk_adj_pct = 1 - ((rank - 1) / (n - 1))
```

---

## 7. Moving Average Flags

```
above_50d  = price_last > MA_50
above_200d = price_last > MA_200
```

---

## 8. Pullback Flag

Triggered when:

```
strong_trend = (3m_return >= 25) and (6m_return >= 40)
recent_dip   = (1w_return <= 0) or (1m_return <= 0)

pullback_flag = strong_trend and recent_dip
```

---

## 9. New Momentum (Acceleration) Flag

Momentum must be positive across all horizons:

```
all_positive = (1w > 0) and (1m > 0) and (3m > 0) and (6m > 0)
```

Acceleration conditions:

```
1m >= (3m / 3)
3m >= (6m / 3)
```

Final:

```
new_momentum_flag = all_positive and acceleration_rules
```

---

## 10. Final Score (Master Ranking)

Weighted combination:

```
final_score =
    0.4 * composite_score
  + 0.3 * risk_adj_pct
  + 0.2 * new_momentum_flag
  + 0.1 * pullback_flag
```

Flags are numeric (1 = Yes, 0 = No).

---

## 11. Sector Assignment

Sector pulled from Yahoo Finance (`Ticker.info["sector"]`).  
Fallback: "Unknown".

---

# End of Appendix
