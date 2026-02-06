import pandas as pd
import os

def clean_excel_file(input_path, output_path):
    print(f"--- Starting processing: {input_path} ---")
    
    try:
        # 1. Read file (Engine 'openpyxl' is safer for .xlsx)
        df = pd.read_excel(input_path, engine='openpyxl')
        initial_rows = len(df)
        
        # --- PHASE 1: DATA CLEANING (Core Logic) ---
        
        # Rule 1: Remove Duplicates (General, considering all columns)
        df = df.drop_duplicates()
        print(f"-> Duplicates removed: {initial_rows - len(df)} lines discarded.")
        
        # Rule 2: Remove 100% empty lines (Garbage Collection)
        df = df.dropna(how='all')
        
        # Rule 3: TRIM (The Secret Sauce). Removes whitespace before/after text in all columns.
        # This iterates through the whole DataFrame looking for text (object) to clean.
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        print("-> Extra whitespace removed (Trim applied).")
        
        # (Optional) Fill remaining empty cells with "N/A" or leave blank?
        # For pure cleaning, we usually leave them empty or fill with ""
        df = df.fillna("") 

        # --- PHASE 2: SAVING WITH POLISH (Enhanced formatting) ---
        print(f"Saving final file to: {output_path}")
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Save data
            df.to_excel(writer, index=False, sheet_name='CleanData')
            
            # Access sheet for formatting
            worksheet = writer.sheets['CleanData']
            
            print("Applying column auto-adjustment...")
            
            # Formatting Loop (Auto-width adjustment)
            for col in worksheet.columns:
                max_length = 0
                column_letter = col[0].column_letter 

                for cell in col:
                    try:
                        if cell.value:
                            cell_text = str(cell.value)
                            if len(cell_text) > max_length:
                                max_length = len(cell_text)
                    except:
                        pass
                
                # Set width (capped at 50 to prevent huge columns)
                adjusted_width = min(max_length + 2, 50) 
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
        print(">>> SUCCESS! Cleaned file saved.")
        return True

    except Exception as e:
        print(f">>> CRITICAL ERROR: {e}")
        return False

# --- TEST AREA (Simulates GUI behavior) ---
if __name__ == "__main__":
    # Create a dummy Excel file to test or use your 'vendas_baguncadas.xlsx'
    test_file = "vendas_baguncadas.xlsx" 
    
    # Check if file exists before running
    if os.path.exists(test_file):
        clean_excel_file(test_file, "sales_CLEANED.xlsx")
    else:
        print(f"Please create a file named '{test_file}' to test the script.")
