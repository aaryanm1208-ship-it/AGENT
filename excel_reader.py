import pandas as pd

def load_excel(file_name):
    try:
        df = pd.read_excel(file_name, engine="openpyxl")
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        print("Error:", e)
        return None
