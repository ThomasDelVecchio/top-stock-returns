# Top 1000 Stock Trend Analyzer

## Overview
Analyzes the top ~1000 US-listed companies and computes:
- 1W, 1M, 3M, 6M returns
- Multi-horizon trend scores (percentile-based)
- Risk-adjusted rankings
- Pullback & momentum flags
- Sector leaders
- Final top 10 recommended opportunities

## Key Outputs
### 1. Master
Raw returns for all tickers across all horizons.

### 2. Top Horizon Sheets
Ranked lists:
- Top_1w
- Top_1m
- Top_3m
- Top_6m

### 3. Top_Trend_10
Top multi-horizon trend leaders based on:
- Percentile scores for each horizon
- Composite trend score
- Momentum & pullback signals
- Sector, volatility, and moving-average filters

### 4. Top_Final_10
Final ranked opportunities using:
- 40% composite trend strength
- 30% risk-adjusted score
- 20% new momentum flag
- 10% pullback flag

### 5. Top_Sector_Leaders
Top names per sector based on composite strength.

## How Trend Scores Work
Each horizon (1W, 1M, 3M, 6M) is ranked → converted to a percentile score (0–1).
Composite trend score = average of all four percentile scores.

## Running the Script (Windows)
```
python top_return_ranks.py
```

## Running in Google Colab
1. Mount Drive  
2. Install dependencies  
3. Download script from GitHub  
4. Run script  
5. Outputs sync automatically to Drive and download to device

## Goal
Surface **early trend opportunities** by combining momentum, consistency, volatility, and sector strength.
