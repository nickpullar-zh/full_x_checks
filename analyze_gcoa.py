import pandas as pd

# Read the GCoA Excel file
gcoa_path = r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks\test_data\GCoA file with 13106 Data rows on sheet GCoA Base account table.xlsx"

try:
    # Load the sheet
    gcoa_df = pd.read_excel(gcoa_path, sheet_name="GCoA Base account table")
    
    # Get basic info
    print("="*80)
    print("GCoA FILE STRUCTURE ANALYSIS")
    print("="*80)
    print("\nShape: {} rows x {} columns".format(gcoa_df.shape[0], gcoa_df.shape[1]))
    print("\nColumn Names:")
    for i, col in enumerate(gcoa_df.columns, 1):
        print("  {}. {}".format(i, col))
    
    # First 10 rows
    print("\n\nFIRST 10 ROWS:")
    print(gcoa_df.head(10).to_string())
    
    # Find Item Type column
    item_type_cols = [col for col in gcoa_df.columns if 'item' in col.lower() or 'type' in col.lower()]
    print("\n\nPOTENTIAL 'ITEM TYPE' COLUMNS: {}".format(item_type_cols))
    
    # If found, count QU
    if item_type_cols:
        col_name = item_type_cols[0]
        qu_count = (gcoa_df[col_name] == 'QU').sum()
        print("\n\nRows where '{}' = 'QU': {}".format(col_name, qu_count))
        
        # Show sample QU rows
        qu_rows = gcoa_df[gcoa_df[col_name] == 'QU'].head(10)
        print("\n\nSAMPLE QU ROWS (first 10):")
        print(qu_rows.to_string())
    
    # Find Account Number column
    acct_cols = [col for col in gcoa_df.columns if 'account' in col.lower()]
    print("\n\nACCOUNT-RELATED COLUMNS: {}".format(acct_cols))
    
except Exception as e:
    print("Error: {}".format(e))
    import traceback
    traceback.print_exc()
