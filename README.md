# Fragrance Price Intelligence — Web App

## What it does
Upload any supplier price list → get instant benchmarking against 5 wholesaler databases (MTZ, Nandansons, PCA, GE, PTC) → download PriceAnalysis + PackagePrice Excel reports.

## How to run locally (easiest)

1. Install Python 3.9+ from python.org
2. Open Terminal / Command Prompt
3. Run these commands:

```
pip install streamlit pandas openpyxl xlrd
streamlit run app.py
```

4. Browser opens at http://localhost:8501
5. Upload your wholesaler files in the sidebar (one-time per session)
6. Upload any offer file and click "Run Price Analysis"

## How to deploy to Streamlit Cloud (free, always online)

1. Create a free account at share.streamlit.io
2. Upload app.py and requirements.txt to a GitHub repo
3. Connect the repo on Streamlit Cloud
4. Your app is live at a permanent URL — no more chat sessions needed!

## Deal scoring
- 🟢 −40%+ → SHARP — excellent deal
- 🔵 −35 to −40% → GOOD — solid margin
- 🟠 −30 to −35% → MINIMUM — viable, proceed with caution
- ⚠️ Below −30% → counter or negotiate
- 🔴 Above market → hard pass

## EUR conversion
Check the "Convert EUR → USD" box and enter the exchange rate before running.
