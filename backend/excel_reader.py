from pathlib import Path
import pandas as pd


def read_file(data_path):
    df = pd.read_excel(data_path)
    df.columns = df.columns.str.strip()
    return df
