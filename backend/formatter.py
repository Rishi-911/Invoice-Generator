import pandas as pd

def format_value(key,value):
    if pd.isna(value):
        return ""
    if isinstance(value,pd.Timestamp):
        return value.strftime("%d-%m-%y")

    currency_columns = {
        "Amount",
        "Unit Price",
        "Subtotal",
        "Total",
        "IGST",
        "CGST",
        "SGST"
    }

    if key in currency_columns:
        return f"{float(value):,.2f}"
    return str(value)