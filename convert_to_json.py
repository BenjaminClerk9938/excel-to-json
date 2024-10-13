import pandas as pd
import json

def convert_excel_to_json(excel_file_path, json_file_path):
    # Read the Excel file
    df = pd.read_excel(excel_file_path)
    columns_to_include = ['issuer_name', 'issuer_description']  # Replace with your column names
    df_selected = df[columns_to_include]
    # Convert DataFrame to JSON
    json_data = df_selected.to_json(orient='records')

    # Save JSON data to a file
    with open(json_file_path, 'w') as json_file:
        json_file.write(json_data)
