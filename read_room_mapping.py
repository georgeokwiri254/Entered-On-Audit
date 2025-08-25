"""
Read the room mapping Excel file to understand correct room codes
"""

import pandas as pd
import os

def read_room_mapping():
    """Read the room mapping Excel file"""
    
    file_path = r"C:\Users\reservations\Documents\EXCEL\Entered On Audit\Entered On room Map.xlsx"
    
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return
    
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        print("ROOM MAPPING FROM EXCEL FILE")
        print("="*80)
        print(f"File: {os.path.basename(file_path)}")
        print(f"Columns: {list(df.columns)}")
        print(f"Total rows: {len(df)}")
        print("\nRoom Mapping Data:")
        print("-" * 80)
        
        # Display the data
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', None)
        pd.set_option('display.max_colwidth', 50)
        
        print(df)
        
        print("\n" + "="*80)
        print("ROOM MAPPING ANALYSIS COMPLETED")
        print("="*80)
        
        return df
        
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

if __name__ == "__main__":
    mapping_df = read_room_mapping()