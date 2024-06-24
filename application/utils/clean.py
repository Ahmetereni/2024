import pandas as pd
import re
def clean_excel(data):
    for column in data.columns:
        data[column] = data[column].apply(lambda x: clean_cell_value(x) if pd.notnull(x) else x)
    return data

def clean_cell_value(cell_value):
    if isinstance(cell_value, str):
        if re.search(r'\d', cell_value) and re.search(r'\D', cell_value):
            numbers = re.findall(r'\d+', cell_value)
            return ' '.join(numbers)
    return cell_value

def group_and_average(data:pd.DataFrame)-> pd.DataFrame:
    # Convert columns to numeric, errors='coerce' will convert non-numeric to NaN
    numeric_data = data.apply(pd.to_numeric, errors='coerce')
    # Select only numeric columns for grouping and averaging
    numeric_data = numeric_data.dropna(axis=1, how='all')
    grouped_data = numeric_data.groupby(data.iloc[:, 0]).mean()
    return grouped_data
