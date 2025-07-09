import pandas as pd

def read_data(file_path):
    ext = file_path.split('.')[-1].lower()
    if ext in ['xlsx', 'xls']:
        return pd.read_excel(file_path)
    elif ext == 'csv':
        return pd.read_csv(file_path)
    elif ext == 'ods':
        return pd.read_excel(file_path, engine='odf')
    else:
        raise ValueError(f"Unsupported file format: {ext}")

def apply_conditions(df1, df2, conditions):
    result = df1.copy()
    for field, condition_type in conditions:
        if field not in df1.columns or field not in df2.columns:
            raise ValueError(f"Field '{field}' not found in both datasets")

        if condition_type == "Совпадают":
            result = result[result[field].isin(df2[field])]
        elif condition_type == "Не совпадают":
            result = result[~result[field].isin(df2[field])]
        else:
            raise ValueError(f"Unknown condition type: {condition_type}")
    return result
