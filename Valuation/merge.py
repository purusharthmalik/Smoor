import pandas as pd

def merge_excel_files(input_files, output_file):
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for file in input_files:
            sheet_name = file.split('/')[-1].split('.')[0]
            print(sheet_name)
            df = pd.read_excel(file)
            df.to_excel(writer, 
                        sheet_name=sheet_name, 
                        index=False)

    print(f"All files have been merged into {output_file}")

input_files = ["C:/Users/purus/Documents/GitHub/Smoor/Valuation/Output_Lounge.xlsx",
               'C:/Users/purus/Documents/GitHub/Smoor/Valuation/Output_Factory.xlsx',
               'C:/Users/purus/Documents/GitHub/Smoor/Valuation/Output_HK.xlsx',
               'C:/Users/purus/Documents/GitHub/Smoor/Valuation/Output_WH.xlsx']

output_file = 'Inventory_Valuation_Consolidated.xlsx'

merge_excel_files(input_files, output_file)