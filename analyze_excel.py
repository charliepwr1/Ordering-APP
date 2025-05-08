import pandas as pd

# Path to the Excel file
file_path = '/mnt/c/Users/charl/Projects/Cannabis-order-app/CannabisRetailersManualOrderForm.xlsm'

try:
    # Get the sheet names
    excel_file = pd.ExcelFile(file_path)
    print(f"Sheet names in the file: {excel_file.sheet_names}")
    
    # Try to read the Catalogue sheet (if it exists)
    if 'Catalogue' in excel_file.sheet_names:
        # Read with openpyxl engine to handle .xlsm files
        df = pd.read_excel(file_path, sheet_name='Catalogue', engine='openpyxl')
        print(f"\nColumns in the Catalogue sheet: {df.columns.tolist()}")
        
        # Check for EachesPerCase column
        if 'EachesPerCase' in df.columns:
            print("\nEachesPerCase column FOUND!")
            # Show first few values
            print(f"First 5 values: {df['EachesPerCase'].head().tolist()}")
        else:
            print("\nEachesPerCase column NOT found!")
            
            # Look for columns with 'case' in the name
            case_cols = [col for col in df.columns if 'case' in col.lower()]
            if case_cols:
                print(f"Found case-related columns: {case_cols}")
                # Show sample values from these columns
                for col in case_cols:
                    print(f"Sample values for {col}: {df[col].head().tolist()}")
            else:
                print("No case-related columns found.")
                
        # Show the first 5 rows for inspection
        print("\nFirst 5 rows of data:")
        print(df.head().to_string())
    else:
        # If 'Catalogue' not found, try other sheets
        for sheet in excel_file.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet, engine='openpyxl')
            print(f"\nColumns in sheet '{sheet}': {df.columns.tolist()}")
            
            # Check for EachesPerCase in this sheet
            if 'EachesPerCase' in df.columns:
                print(f"EachesPerCase column FOUND in sheet '{sheet}'!")
                break
        
except Exception as e:
    print(f"Error analyzing Excel file: {str(e)}")