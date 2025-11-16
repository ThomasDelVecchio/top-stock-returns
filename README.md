# Stock Trend Analysis Script (Simplified & Consolidated)

This repository contains a full pipeline that:

1.  Scrapes \~1,000 tickers from StockAnalysis\
2.  Downloads 1-year price history from Yahoo Finance\
3.  Computes 1W, 1M, 3M, 6M returns\
4.  Calculates trend scores, risk-adjusted metrics, moving-average
    health flags\
5.  Ranks the strongest stocks\
6.  Outputs a clean Excel report with **four tabs**:
    -   **Master** -- full universe (sortable)
    -   **Top_Trend_10** -- best consistent performers
    -   **Top_Sector_Leaders** -- sector leadership
    -   **Top_Final_10** -- final recommendations

------------------------------------------------------------------------

## 📊 Sheet Descriptions

### **Master**

Full universe of tickers with raw returns.\
You can sort 1W, 1M, 3M, or 6M to see leaders.

### **Top_Trend_10**

Shows the strongest performers across all time horizons.\
Includes: - Trend Score\
- Risk-Adjusted Strength & Rank\
- Trend Health Flags (50d / 200d)\
- Speeding Up?\
- Recent Pullback?

### **Top_Sector_Leaders**

Top 3 stocks per sector, ranked by composite trend strength.

### **Top_Final_10 (Final Recommendations)**

Weighted scoring system: - **40%** Trend Score\
- **30%** Risk-Adjusted Rank\
- **20%** Speeding Up\
- **10%** Recent Pullback

This list represents the highest-rated opportunities.

------------------------------------------------------------------------

## 🚀 How to Run

``` bash
python top1000_trend_analysis.py
```

Outputs: - `top1000_returns_master.csv` - `top1000_return_report.xlsx`

------------------------------------------------------------------------

## 📁 Files Included

-   `top1000_trend_analysis.py` -- main script\
-   `Trend_Analysis_CheatSheet_UPDATED.docx` -- simplified explanation\
-   `top1000_return_report.xlsx` -- generated after running

------------------------------------------------------------------------

## 📝 Notes

This version removes redundant tabs (Top_1w, Top_1m, Top_3m, Top_6m).\
Sorting is now done directly inside the **Master** sheet.
