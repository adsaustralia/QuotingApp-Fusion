# Quoting App — Row-based Input (Streamlit)

Use this app when customer templates store values in **fixed rows**, and each **column** represents a store/item.

Example rows:
- Size row = 5
- Material row = 6
- Sides row = 10
- Qty row = 57

You pick a **column range** (e.g. AC:IG).  
The app creates one line item per column where Qty > 0.

## DS loading
If sides is **DS**, line total is multiplied by **(1 + DS loading %)**.

## Export
Export preserves original formatting by editing the original workbook and writing prices into:
- **Price row** (defaults to Qty row + 1, i.e. "next to qty").

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```


## Restore saved settings
When you save a sheet to the bundle, the app stores the row/column settings and will restore them when you return to that sheet (if Auto-restore is enabled).
