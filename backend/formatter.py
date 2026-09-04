import pandas as pd
from datetime import datetime, date

def format_indian_number(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    
    try:
        num = float(value)
    except (ValueError, TypeError):
        return str(value)

    formatted_str = f"{abs(num):.2f}"
    int_part, dec_part = formatted_str.split('.')

    if len(int_part) <= 3:
        grouped_int = int_part
    else:
        last_three = int_part[-3:]
        remaining = int_part[:-3]

        groups = []
        while len(remaining) > 2:
            groups.append(remaining[-2:])
            remaining = remaining[:-2]
        if remaining:
            groups.append(remaining)

        groups.reverse()
        grouped_int = ",".join(groups) + "," + last_three

    sign = "-" if num < 0 and float(formatted_str) != 0 else ""
    return f"{sign}{grouped_int}.{dec_part}"


def format_value(key, value):
    if pd.isna(value) or value is None:
        return ""
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.strftime("%d/%m/%Y")

    currency_columns = {
        "amount",
        "unit price",
        "subtotal",
        "total",
        "igst",
        "cgst",
        "sgst"
    }

    if str(key).strip().lower() in currency_columns:
        return format_indian_number(value)
    return str(value)
