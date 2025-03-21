import os
import pandas as pd

def consolidate_excel_files(input_directory, output_file='master_stocklist.xlsx'):
    """
    Search a directory for Excel files and consolidate EAN and price columns into a master spreadsheet.
    
    Parameters:
    - input_directory (str): Path to the directory containing Excel files
    - output_file (str, optional): Name of the output master spreadsheet file
    
    Returns:
    - pandas.DataFrame: Consolidated DataFrame with EAN and price columns
    """
    # List to store individual DataFrames
    all_data = []

    def check_columns_and_add(dataframe):
        try:
            # Add filename column to track source
            dataframe['source_file'] = filename
            
            # Select only EAN and price columns (and source file)
            dataframe = dataframe[['EAN', 'QTY', 'EUR', 'USD', 'GBP', 'source_file', 'DESCRIPTION']]
            for column in dataframe:
                # Append to list of DataFrames
                try:
                    print(column)
                    all_data.append(column)
                except Exception:
                    print(f'No column {column}')

        except Exception as e:
            print(f'error: {str(e)}')

    
    # Walk through the directory
    for filename in os.listdir(input_directory):
        # Check if file is an Excel file
        if filename.endswith(('.xlsx', '.xls', '.xlsm', '.xlsb', '.odf', '.ods', '.odt')):
            try:
                # Full path to the file
                file_path = os.path.join(input_directory, filename)
                
                # Read the Excel file
                df = pd.read_excel(file_path)
                
                check_columns_and_add(df)
            
            except Exception as e:
                print(f"Error processing {filename}: {str(e)}")
    
    # Consolidate all DataFrames
    if all_data:
        master_df = pd.concat(all_data, ignore_index=True)
        
        # Save to Excel
        master_df.to_excel(output_file, index=False)
        
        print(f"Master spreadsheet saved as {output_file}")
        
        return master_df
    else:
        print("No valid Excel files found with required columns.")
        return None

# Example usage
if __name__ == "__main__":

    input_directory = '/Users/mygollygosh/Desktop/codestuff/excel_consolidation/stock_lists'
    
    # Run the consolidation
    result = consolidate_excel_files(input_directory)
    
    # Optional: Display the first few rows
    if result is not None:
        print("\nMaster Spreadsheet Preview:")
        print(result.head())