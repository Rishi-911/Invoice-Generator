from pathlib import Path
import pandas as pd

def read_file(data_path):
    df = pd.read_excel(data_path)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    return df

