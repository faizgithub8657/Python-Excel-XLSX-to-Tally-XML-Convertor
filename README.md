# ABHL Excel → Tally Prime XML Converter

A web tool by **ABHL & Associates, Chartered Accountants** that converts
an Excel voucher sheet into a Tally Prime-importable XML file.

🔗 **Live app:** https://abhl-tally.streamlit.app

## How to use

1. Open the live app URL.
2. Upload your Excel file (`.xlsx` or `.xlsm`).
3. Confirm the sheet name (default `Sheet1`).
4. Click **Convert**, then download the XML.
5. In Tally Prime → **Gateway of Tally → Import → Vouchers → XML File** → browse to the downloaded XML.

## Required Excel columns

The converter auto-detects the header row anywhere on the sheet (A1, C3, anywhere).
Headers needed (case-insensitive):

| Column | Required? | Notes |
|---|---|---|
| `DATE` | Yes | Date cell, or text in DD-MM-YYYY / DD/MM/YYYY / YYYY-MM-DD |
| `Vch_num` | Optional | Voucher number / reference |
| `Narration` | Optional | Free text |
| `VCHTYPE` | Yes | PAYMENT, RECEIPT, JOURNAL, CONTRA, or custom |
| `Dr_Ledger_1` … `Dr_Ledger_4` | At least one | Ledger name (must exist in Tally) |
| `Dr_amt_1` … `Dr_amt_4` | Pair with each Dr ledger | Positive amount |
| `Cr_Ledger_1` … `Cr_Ledger_4` | At least one | Ledger name |
| `Cr_amt_1` … `Cr_amt_4` | Pair with each Cr ledger | Positive amount |

Each row must balance: total Dr amounts = total Cr amounts.

## Tally sign convention used

- `ISDEEMEDPOSITIVE=YES` → debit leg, `AMOUNT` is stored negative
- `ISDEEMEDPOSITIVE=NO` → credit leg, `AMOUNT` is stored positive

For `RECEIPT` vouchers, the converter automatically swaps:
the "Dr_Ledger" Excel column (party received from) becomes the Cr leg
in XML, and the "Cr_Ledger" Excel column (bank) becomes the Dr leg.

## Running locally (optional)

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

Or CLI mode without the UI:

```bash
python converter.py input.xlsx output.xml Sheet1
```

## Files in this repo

| File | Purpose |
|---|---|
| `streamlit_app.py` | The web UI |
| `converter.py` | Core conversion logic (importable; also runnable from CLI) |
| `requirements.txt` | Python dependencies for Streamlit Cloud |
| `README.md` | This file |

---

© ABHL & Associates, Chartered Accountants
