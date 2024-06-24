import openpyxl
import re
import pandas as pd


def group_and_average(data):
    """
    Groups the data by the first column and calculates the average for each group.
    
    Args:
    data (pd.DataFrame): The input data to be grouped and averaged.
    
    Returns:
    pd.DataFrame: The grouped and averaged data.
    """
    grouped_data = data.groupby(data.columns[0]).mean()
    return grouped_data

if __name__ == '__main__':
    input_file_path = "cleaned.xlsx"
    output_file_path = "grouped_file.xlsx"
    
    # Step 1: Clean the Excel file
    #clean_excel(input_file_path, output_file_path)
    
    # Step 2: Read the cleaned Excel file into a pandas DataFrame
    data = pd.read_excel(input_file_path)
    
    
    # Step 3: Group the data by the first column and calculate the average for each group
    demogroup = group_and_average(data)
    
    # Step 4: Save the grouped and averaged data to a new Excel file
    demogroup.to_excel("grouped_averages.xlsx")
