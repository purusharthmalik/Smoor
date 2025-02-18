import pandas as pd

def merge_excel_files(input_files, output_file):
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for file in input_files:
            sheet_name = file.split('/')[-1].split('.')[0]
            print(sheet_name)
            df = pd.read_csv(file)
            df.to_excel(writer, 
                        sheet_name=sheet_name, 
                        index=False)

    print(f"All files have been merged into {output_file}")

month = 'Nov'
input_files = [f'C:/Users/purus/Documents/GitHub/Smoor/BLR_{month}.csv',
               f'C:/Users/purus/Documents/GitHub/Smoor/MUM_{month}.csv',
               f'C:/Users/purus/Documents/GitHub/Smoor/DEL_{month}.csv',
               f'C:/Users/purus/Documents/GitHub/Smoor/GUR_{month}.csv',
               f'C:/Users/purus/Documents/GitHub/Smoor/CHE_{month}.csv']
output_file = 'Cost_Center_Final_Oct_Revised.xlsx'

merge_excel_files(input_files, output_file)