import os
import numpy as np
import pandas as pd

print("Loading master...")
name_master = pd.read_excel(r"C:\Users\purus\Documents\GitHub\Smoor\Valuation\Data\Data for Closing Valuations.xlsx",
                            sheet_name="Updated Name and city")
category_master = pd.read_excel(r"C:\Users\purus\Documents\GitHub\Smoor\Valuation\Data\Data for Closing Valuations.xlsx",
                            sheet_name="Category")

print("Loading dumps...")
folder ="C:/Users/purus/Documents/GitHub/Smoor/Valuation/Closing Files Nov 24/Factory/"
files = os.listdir(folder)

dumps = [folder + file for file in files]

columns=['Revised SKU Code', 'SKU Name', 'Item Type', 'Location', 'Sub-Location', 'City', 'Category', 'UOM', 'Qty']
final_dump = pd.DataFrame(columns=columns)

for dump in dumps:
    try:
        df = pd.read_excel(dump)[columns]
    except:
        df = pd.read_excel(dump, header=1)[columns]
    final_dump = pd.concat([final_dump, df])

final_dump.dropna(subset=['Revised SKU Code', 'Qty', 'City'],
          inplace=True)
final_dump = final_dump[final_dump['Qty'] != 0]
final_dump.reset_index(drop=True, inplace=True)

print("Formatting dumps...")
final_df = pd.DataFrame(columns=[
    'Con',
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
con, updated_name, qty = [], [], [],

print("Filling in the item details...")
for row in final_dump.iterrows():
    idx, row_data = row
    con.append(row_data['City'] + str(row_data['Revised SKU Code']))
    try:
        qty.append(float(row_data['Qty']))
    except:
        qty.append(np.nan)

    # Sub Location rules
    if row_data['Sub-Location'].startswith('HAL') or row_data['Sub-Location'].startswith('Jigani'):
        updated_name.append('Bangalore')
    else:
        updated_name.append(row_data['Sub-Location'])

temp_df = pd.DataFrame({
    'Con': con,
    'SKU Code': final_dump['Revised SKU Code'],
    'SKU Name': final_dump['SKU Name'],
    'Item Type': final_dump['Item Type'],
    'Location': final_dump['Location'].apply(lambda x: x.title()),
    'Sub-Location': final_dump['Sub-Location'].apply(lambda x: x.title()),
    'Updated Name': updated_name,
    'City': final_dump['City'].apply(lambda x: x.title()),
    'Department': final_dump['Category'],
    'Category': final_dump['Category'],
    'UOM': final_dump['UOM'],
    'Qty': qty,
})
final_df = pd.concat([final_df, temp_df])

leave = []
for idx, code in enumerate(final_df['SKU Code'].values):
    if code.startswith('SFG') or code.startswith('SF'):
        final_df.loc[idx, 'Item Type'] = 'SFG'
    elif code.startswith('CM') or code.startswith('PS'):
        leave.append(idx)

idxs = list(range(final_df.shape[0]))
for _ in leave:
    idxs.remove(_)

final_df = final_df.iloc[idxs]

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

print("Loading the smoor and b2b sheet...")
smoor_sheet = pd.read_excel(rate_master,
                            sheet_name=['Smoor Product'],
                            header=1)['Smoor Product'][['FG Code', 'FnP (At Factory Level)']]
b2b_sheet = pd.read_excel(rate_master,
                          sheet_name=['B2B Products'],
                          header=1)['B2B Products'][['FG Code', 'FnP (At Factory Level)']]

print("Loading the FG-SFG sheet...")
fg_sfg_sheet = pd.read_excel(rate_master,
                          sheet_name=['FG-SFG'])['FG-SFG'][['SFG Code', 'Yield']]

print("Loaded. Filling in the sheet...")
val_rates = []
for code, city, item_type, category in zip(final_df['SKU Code'], final_df['City'], final_df['Item Type'], final_df['Category']):
    if item_type == 'RM' or item_type == 'PK' or item_type == 'PM':
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
    elif item_type == 'FG':
        if city == 'Bangalore':
            try:
                val_rates.append(smoor_sheet[smoor_sheet['FG Code'] == code]['FnP (At Factory Level)'].values[0])
            except:
                try:
                    val_rates.append(b2b_sheet[b2b_sheet['FG Code'] == code]['FnP (At Factory Level)'].values[0])
                except:
                    val_rates.append(0)
        elif city in ['Mumbai', 'Pune']:
            if category in ['Cakes', 'Bakery', 'Tea cakes']:
                try:
                    val_rates.append(smoor_sheet[smoor_sheet['FG Code'] == code]['FnP (At Factory Level)'].values[0])
                except:
                    try:
                        val_rates.append(b2b_sheet[b2b_sheet['FG Code'] == code]['FnP (At Factory Level)'].values[0])
                    except:
                        val_rates.append(0)
            else:
                try:
                    val_rates.append(mum_pr_master.loc[mum_pr_master['Item Code'] == code]['Rate'].values[0])
                except:
                    val_rates.append(0)
        else:
            if category == 'Cakes':
                try:
                    val_rates.append(smoor_sheet[smoor_sheet['FG Code'] == code]['FnP (At Factory Level)'].values[0])
                except:
                    try:
                        val_rates.append(b2b_sheet[b2b_sheet['FG Code'] == code]['FnP (At Factory Level)'].values[0])
                    except:
                        val_rates.append(0)
            else:
                try:
                    val_rates.append(mum_pr_master.loc[mum_pr_master['Item Code'] == code]['Rate'].values[0])
                except:
                    val_rates.append(0)
    elif item_type == 'SF' or item_type == 'SFG':
        val_rates.append(fg_sfg_sheet[fg_sfg_sheet['SFG Code'] == code]['Yield'].sum())
    else:
        val_rates.append(0)

final_df['Rate'] = val_rates
final_df['Valuation'] = final_df['Rate'] * final_df['Qty']

final_df.to_excel('Output_Factory.xlsx', index=False)
print("File saved!")