import pandas as pd
import glob

print("Reading all the secondary files...")
rm_pm_issue = pd.read_excel("C:/Users/purus/Documents/GitHub/Smoor/Consumption/RM-PM Issue.xlsx")[['Item Code', 'Type']]
fg_sfg_issue = pd.read_csv("C:/Users/purus/Documents/GitHub/Smoor/Consumption/FG-SFG(FG-SFG).csv", encoding='latin1')[['Item Code', 'Type']]
type_master = pd.concat([rm_pm_issue, fg_sfg_issue], ignore_index=True)
category_master = pd.read_excel("C:/Users/purus/Documents/GitHub/Smoor/Valuation/Data/Data for Closing Valuations.xlsx", sheet_name="Category")

# Adding store issue reports
print("Working on the store issue reports...")
path = "C:/Users/purus/Documents/GitHub/Smoor/Consumption/ERP DUMPS Sept 2024/store issue report 2024 sept/*.xlsx"
all_files = glob.glob(path)

df_list = [pd.read_excel(file) for file in all_files]
store_issues = pd.concat(df_list, ignore_index=True)[['Posting Date', 'Item Name', 'Item Code', 'Item Group', 'UOM', 'Qty', 'Valuation rate', 'Amount', 'Issued To', 'Request Reference']]
store_issues['Source'] = ['Direct Issue'] * store_issues.shape[0]
store_issues['Department'] = store_issues['Issued To']
cat, typ = [], []
for idx, row in store_issues.iterrows():
    try:
        assert category_master[category_master['Item Code'] == row['Item Code']]['Category'].values[0]
        cat.append('Lounges')
    except:
        cat.append(row['Department'])
    try:
        typ.append(type_master[type_master['Item Code'] == row['Item Code']]['Type'].values[0])
    except:
        typ.append('')
store_issues['Category'] = cat
store_issues['Type'] = typ

# Adding purchase reports (bliss entity)
print("Working on the purchase receipts (bliss entity)...")
path = "C:/Users/purus/Documents/GitHub/Smoor/Consumption/ERP DUMPS Sept 2024/PURCHASE RECEIPT 2024 sept/purchase reciept(bliss entity)/*.xlsx"
all_files = glob.glob(path)

df_list = [pd.read_excel(file) for file in all_files]
purchase_reports = pd.concat(df_list, ignore_index=True)[['Date', 'Item Name', 'Item Code', 'UOM', 'Accepted Quantity', 'Rate', 'Amount', 'Accepted Warehouse', 'Purchase Receipt: Name']]
purchase_reports = purchase_reports[purchase_reports['Accepted Warehouse'].apply(lambda x: x not in ['HAL Stores - BCIPL', 'Jigani Stores - BCIPL'])]
purchase_reports.rename(columns={'Date':'Posting Date',
                                 'Accepted Quantity':'Qty',
                                 'Rate':'Valuation rate',
                                 'Accepted Warehouse':'Issued To',
                                 'Purchase Receipt: Name':'Request Reference'}, inplace=True)
purchase_reports['Item Group'] = [' '] * purchase_reports.shape[0]
purchase_reports['Source'] = ['Direct Purchase - Bliss'] * purchase_reports.shape[0]
purchase_reports['Department'] = ['Hot Kitchen'] * purchase_reports.shape[0]
purchase_reports['Category'] = ['Ready to Eat'] * purchase_reports.shape[0]
typ = []
for idx, row in purchase_reports.iterrows():
    try:
        typ.append(type_master[type_master['Item Code'] == row['Item Code']]['Type'].values[0])
    except:
        typ.append('')
purchase_reports['Type'] = typ

# Adding purchase reports (smoor entity)
print("Working on the purchase receipts (smoor entity)...")
path = "C:/Users/purus/Documents/GitHub/Smoor/Consumption/ERP DUMPS Sept 2024/PURCHASE RECEIPT 2024 sept/purchase reciept(smoor entity)/*.xlsx"
all_files = glob.glob(path)

df_list = [pd.read_excel(file) for file in all_files]
smoor_reports = pd.concat(df_list, ignore_index=True)[['Date', 'Supplier Name', 'Item Name', 'Item Code', 'UOM', 'Accepted Quantity', 'Rate', 'Amount', 'Accepted Warehouse', 'Purchase Receipt: Name']]
smoor_reports = smoor_reports[smoor_reports['Supplier Name'].apply(lambda x: x.split(' ')[0]) != 'Bliss']
smoor_reports.drop('Supplier Name', axis=1, inplace=True)
smoor_reports.rename(columns={'Date':'Posting Date',
                                 'Accepted Quantity':'Qty',
                                 'Rate':'Valuation rate',
                                 'Accepted Warehouse':'Issued To',
                                 'Purchase Receipt: Name':'Request Reference'}, inplace=True)
smoor_reports['Item Group'] = [' '] * smoor_reports.shape[0]
smoor_reports['Source'] = ['Direct Purchase - SMOOR'] * smoor_reports.shape[0]
smoor_reports['Department'] = ['Lounges'] * smoor_reports.shape[0]
smoor_reports['Category'] = ['Beverages'] * smoor_reports.shape[0]
typ = []
for idx, row in smoor_reports.iterrows():
    try:
        typ.append(type_master[type_master['Item Code'] == row['Item Code']]['Type'].values[0])
    except:
        typ.append('')
smoor_reports['Type'] = typ

# Concatenating all the files
final_df = pd.concat([store_issues, purchase_reports, smoor_reports], 
                     ignore_index=True)
final_df['Lounge Name'] = final_df['Issued To']
final_df.to_excel('Consumption_Output.xlsx', index=False)
print("File saved!")