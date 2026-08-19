# VIBES × Prelegent Inventory Command Center

Premium Streamlit inventory portal for VIBES consignment stock at Prelegent.

## Business logic
- Mall current stock = `Salable Stock`
- Warehouse current stock = `WI Stock`
- Current physical stock = Mall + Warehouse
- `WI Out Stock` = movement/history, not added again to physical stock
- `Return Stock` = separate return pool
- Variants with Return Stock > 0 are blocked from Safe Send recommendation
- Warehouse stock is considered before recommending new VIBES shipment

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud
1. Create a GitHub repository.
2. Upload the contents of this folder to the repository root.
3. In Streamlit Community Cloud, create a new app from that repository.
4. Branch: `main`
5. Main file path: `app.py`
6. Python: `3.12` is a safe/default choice.
7. Click Deploy.

The default workbook is stored at:
`data/Prelegent_Inventory_Analysis.xlsx`

A newer `.xlsx` can also be uploaded from the portal sidebar without changing the source code.
