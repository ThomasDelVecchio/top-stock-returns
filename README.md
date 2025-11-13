# 📈 Top 1,000 Stock Return Analyzer
### Automatically ranks the largest 1,000 U.S. stocks by 1-week, 1-month, 3-month, and 6-month returns — and identifies the Top 10 “Trend Candidates.”

---

## 🚀 What This Project Does

This script:
- Scrapes the **top ~1,000 U.S. companies by market cap** from StockAnalysis.com  
- Pulls up to **1 year of historical prices** from Yahoo Finance  
- Calculates **1W / 1M / 3M / 6M** total returns for every ticker  
- Builds:
  - **Master table** of all returns  
  - **Ranked tabs** (Top 1W, Top 1M, Top 3M, Top 6M)  
  - **Trend Candidates tab** → the 10 stocks showing the strongest multi-horizon momentum  
- Saves output as:
  - `top1000_returns_master.csv`  
  - `top1000_return_report.xlsx`

---

## 📊 Trend Candidate Methodology (Top_Trend_10)

To identify early “trend” opportunities, the script computes a **composite score** for each ticker:

1. For each return horizon (**1W, 1M, 3M, 6M**), rank all tickers from highest return → lowest.  
2. Convert each rank into a **percentile score** between 0 and 1:  
   - `1.00 = top performer`  
   - `0.00 = bottom performer`  
3. Average all four percentile scores → **composite_score**.  
4. Filter out any tickers that do NOT have at least **2 positive-return horizons**.  
5. Take the **Top 10 composite scores** → these become the **Top_Trend_10** sheet.

This method highlights names that are strong **across multiple timeframes**, not just short-term spikes.

---

## 🛠 How to Run

### **1. Install dependencies**
```
pip install pandas yfinance lxml openpyxl
```

### **2. Run the script**
```
python top_return_ranks.py
```

This will create:
- `top1000_returns_master.csv`
- `top1000_return_report.xlsx`

---

## 📁 Output Files Explained

### **Master**
All 1W/1M/3M/6M returns for all ~1000 tickers.

### **Top_1w / Top_1m / Top_3m / Top_6m**
Each sheet is sorted by return for that horizon (highest → lowest).

### **Top_Trend_10**
Top 10 stocks with the strongest multi-horizon momentum based on composite scoring.

### **ReadMe (inside Excel)**
Contains detailed descriptions of each tab + methodology.

---

## 💻 Project Folder Structure

```
top-stock-returns/
│
├── top_return_ranks.py
├── README.md
├── .gitignore
│
├── top1000_returns_master.csv        (auto-generated)
└── top1000_return_report.xlsx        (auto-generated)
```

---

## 🧠 Why This Exists

This tool helps you:
- Spot **market leaders** early  
- Identify **momentum shifts** before they show up on mainstream screens  
- Track consistency across different timeframes  
- Build intuition around **trend persistence** and **relative strength**  

---

## 📬 Questions / Improvements?

Open an issue or request enhancements on GitHub!  
