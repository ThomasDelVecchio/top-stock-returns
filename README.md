
# Top 1000 Stock Trend Analyzer

A simple, beginner-friendly tool that scans the **top ~1000 U.S. stocks** and identifies:

### ✅ 1. Short‑term and long‑term performance  
1‑week, 1‑month, 3‑month, and 6‑month returns.

### ✅ 2. Trend strength  
A clean “Trend Score” based on how consistently strong the stock has been across all time frames.

### ✅ 3. Pullback opportunities  
Strong stocks that temporarily dipped (good for bargain‑entry ideas).

### ✅ 4. Momentum opportunities  
Stocks that are accelerating upward recently.

### ✅ 5. Final simplified Top 10 list  
The easiest‑to‑understand “Top 10 Recommendations” based on simplified logic.

This project is designed to be readable by non‑experts.  
No finance experience needed.

---

## 🔧 How It Works (Simple Version)

The script does four main things:

### **1. Gets the top ~1000 stocks**
Scraped from StockAnalysis.com (public website).

### **2. Pulls 1 year of price data**
Using Yahoo Finance (via `yfinance`).

### **3. Calculates returns**
- 1‑week  
- 1‑month  
- 3‑month  
- 6‑month

### **4. Builds scores**
Simplified metrics that anyone can understand:

| Metric | Meaning |
|--------|---------|
| **Trend Score** | How consistently strong the stock has performed across all horizons |
| **Momentum Flag** | Is the stock speeding up recently? |
| **Pullback Flag** | Has a strong stock recently dipped? |
| **Above 200‑day MA** | Is long‑term trend healthy? |
| **Final Score** | Clean combined ranking used for Top Final 10 |

---

## 📊 Output Files

The script creates an Excel file with these simplified tabs:

### **Master**
All raw % returns for every stock.

### **Top_1w / Top_1m / Top_3m / Top_6m**
Ranked lists of the best performers over each time frame.

### **Top_Trend_10 (Simple Version)**
- Trend Score  
- Basic metrics  
- Easy interpretation

### **Top_Final_10 (Simplified Recommendations)**
Your “final” top 10 ideas sorted by:
- Trend strength  
- Momentum  
- Healthy long‑term trend  
- Favorable dips  

### **Legend**
A simple English explanation of every column.

---

## ▶️ How to Run (Windows)

1. Install Python 3.10+
2. Install dependencies:
   ```
   pip install yfinance pandas openpyxl
   ```
3. Run script:
   ```
   python top_return_ranks.py
   ```
4. Find your outputs in:
   - Your project folder  
   - (Optional) Windows Drive  
   - (Optional) Google Drive  

---

## ▶️ How to Run in Google Colab

1. Open a new notebook  
2. Mount Google Drive:
   ```python
   from google.colab import drive
   drive.mount('/content/drive')
   ```
3. Install deps:
   ```python
   !pip install yfinance openpyxl --quiet
   ```
4. Download script from GitHub:
   ```python
   !wget -O top_return_ranks.py "YOUR_GITHUB_RAW_URL"
   ```
5. Run:
   ```python
   !python top_return_ranks.py
   ```
6. Files automatically save and download to your device.

---

## Why This Exists

To help regular people:

- Spot trends earlier  
- Understand performance without complexity  
- Get a clean Top 10 list that makes sense  
- Skip confusing quant terminology  

Made to be simple, readable, and actionable.

---

## License
MIT License.
