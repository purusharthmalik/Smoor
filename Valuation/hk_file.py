import os
import pandas as pd

print("Loading master...")
name_master = pd.read_excel(r"C:\Users\purus\Documents\GitHub\Smoor\Valuation\Data\Data for Closing Valuations.xlsx",
                            sheet_name="Updated Name and city")
category_master = pd.read_excel(r"C:\Users\purus\Documents\GitHub\Smoor\Valuation\Data\Data for Closing Valuations.xlsx",
                            sheet_name="Category")

print("Loading dumps...")
dump = pd.read_excel(r"C:\Users\purus\Documents\GitHub\Smoor\Valuation\Closing Files Nov 24\HK\HK CLOSING STOCK 2024.xlsx")
dump = dump[['ERP Code', 'SKU Name', 'Item Type', 'Location', 'Sub-Location', 'City', 'Department', 'UOM', 'Qty']]
dump = dump[dump['Qty'] != 0]
dump.dropna(inplace=True)
dump.reset_index(drop=True, inplace=True)

print("Formatting dumps...")
final_df = pd.DataFrame(columns=['Con',
    'SKU Code',
    'SKU Name',         
    'Item Type',
    'Location',
    'Sub-Location',
    'Updated Name',
    'City',
    'Department',
    'Category',
    'UOM',
    'Qty'])

missing = set()
con, updated_name, cat = [], [], []

leave = []
item_type = dump['Item Type']
for idx, itype in enumerate(item_type):
    if itype == 'CM' or itype == 'PS':
        leave.append(idx)

# Removing the unnecessary indexes
idxs = list(range(dump.shape[0]))
for _ in leave:
    idxs.remove(_)

dump = dump.iloc[idxs]
sku_code = dump['ERP Code']
sku_name = dump['SKU Name']
loc = dump['Location']
sub_loc = dump['Sub-Location']
city = dump['City']
dep = dump['Department']
uom = dump['UOM']
qty = dump['Qty']
item_type = dump['Item Type']

for idx, row in dump.iterrows():
    match = name_master[name_master['Sub-Location'].apply(lambda x: x.strip().lower()) == row['Sub-Location'].strip().lower()]
    if len(match) != 0:
        con.append(row['City'] + sku_code[idx])
        updated_name.append(match['Updated Name'].values[0])
        cat.append('Ready to Eat')
    else:
        missing.add(row['Sub-Location'])
        con.append('')
        updated_name.append('')
        cat.append('Ready to Eat')

temp_df = pd.DataFrame({
    'Con': con,
    'SKU Code': sku_code,
    'SKU Name': sku_name,
    'Item Type': item_type,
    'Location': loc,
    'Sub-Location': sub_loc,
    'Updated Name': updated_name,
    'City': city,
    'Department': dep,
    'Category': cat,
    'UOM': uom,
    'Qty': qty
})
final_df = pd.concat([final_df, temp_df])
print(missing)

print("Loading the store issues...")
rate_master = r"C:\Users\purus\Documents\GitHub\Smoor\Valuation\Data\Rate Master till Nov.xlsx"
blr_issue_master = pd.read_excel(rate_master,
                            sheet_name=['Store issue report-BLR'],
                            header=1)['Store issue report-BLR'][['Item Code', 'Valuation rate']]
che_issue_master = pd.read_excel(rate_master,
                            sheet_name=['Store issue report-CHN'],
                            header=1)['Store issue report-CHN'][['Item Code', 'Valuation rate']]
mum_issue_master = pd.read_excel(rate_master,
                            sheet_name=['Store issue report-MUM'],
                            header=1)['Store issue report-MUM'][['Item Code', 'Valuation rate']]
gur_issue_master = pd.read_excel(rate_master,
                            sheet_name=['Store issue report-GUR'],
                            header=1)['Store issue report-GUR'][['Item Code', 'Valuation rate']]

print("Loading the purchase reciepts...")
blr_pr_master = pd.read_excel(rate_master,
                            sheet_name=['Bangalore PR'],
                            header=1)['Bangalore PR'][['Item Code', 'Rate']]
mum_pr_master = pd.read_excel(rate_master,
                            sheet_name=['Mumbai PR'],
                            header=1)['Mumbai PR'][['Item Code', 'Rate']]
che_pr_master = pd.read_excel(rate_master,
                            sheet_name=['Chennai PR'],
                            header=1)['Chennai PR'][['Item Code', 'Rate']]
gur_pr_master = pd.read_excel(rate_master,
                            sheet_name=['Gurgaon PR'],
                            header=1)['Gurgaon PR'][['Item Code', 'Rate']]

print("Loaded. Filling in the sheet...")
val_rates = []
for code, city, item_type in zip(final_df['SKU Code'], final_df['City'], final_df['Item Type']):
    if item_type == 'RM' or item_type == 'PM':
        try:
            if city == 'Bangalore':
                val_rates.append(blr_issue_master.loc[blr_issue_master['Item Code'] == code]['Rate'].values[0])
            elif city in ['Mumbai', 'Pune']:
                val_rates.append(mum_issue_master.loc[mum_issue_master['Item Code'] == code]['Rate'].values[0])
            elif city in ['Gurgaon', 'Delhi']:
                val_rates.append(gur_issue_master.loc[gur_issue_master['Item Code'] == code]['Rate'].values[0])
            elif city == 'Chennai':
                val_rates.append(che_issue_master.loc[che_issue_master['Item Code'] == code]['Rate'].values[0])
            else:
                val_rates.append(0)
        except:
            try:
                if city == 'Bangalore':
                    val_rates.append(blr_pr_master.loc[blr_pr_master['Item Code'] == code]['Rate'].values[0])
                elif city in ['Mumbai', 'Pune']:
                    val_rates.append(mum_pr_master.loc[mum_pr_master['Item Code'] == code]['Rate'].values[0])
                elif city in ['Gurgaon', 'Delhi']:
                    val_rates.append(gur_pr_master.loc[gur_pr_master['Item Code'] == code]['Rate'].values[0])
                elif city == 'Chennai':
                    val_rates.append(che_pr_master.loc[che_pr_master['Item Code'] == code]['Rate'].values[0])
                else:
                    val_rates.append(0)
            except:
                val_rates.append(0)
    else:
        val_rates.append(0)

final_df['Rate'] = val_rates
final_df['Valuation'] = final_df['Rate'] * final_df['Qty']

final_df.to_excel('Output_HK.xlsx', index=False)
print("File saved!")