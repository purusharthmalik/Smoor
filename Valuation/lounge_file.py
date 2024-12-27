import os
import pandas as pd

print("Loading master...")
name_master = pd.read_excel(r"C:\Users\purus\Documents\GitHub\Smoor\Valuation\Data\Data for Closing Valuations.xlsx",
                            sheet_name="Updated Name and city")
category_master = pd.read_excel(r"C:\Users\purus\Documents\GitHub\Smoor\Valuation\Data\Data for Closing Valuations.xlsx",
                            sheet_name="Category")

print("Loading dumps...")
dump = pd.read_excel(r"C:\Users\purus\Documents\GitHub\Smoor\Valuation\Closing Files Nov 24\Lounge/30-Nov-2024 - Audit Items By Transaction.xlsx")
dump = dump[dump['Audit Quantity'] != 0]
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
con, sku_code, sku_name, item_type, loc, sub_loc, updated_name, city, dep, cat, uom, qty = [], [], [], [], [], [], [], [], [], [], [], []

for row_data in dump.iterrows():
    idx, row = row_data
    match = name_master[name_master['Sub-Location'].apply(lambda x: x.strip().lower()) == row['Branch Name'].strip().lower()]
    if len(match) != 0:
        con.append(match['City'].values[0] + str(row['SKU']))
        sku_code.append(row['SKU'])
        sku_name.append(row['Item Name'])
        item_type.append(row['Type'])
        loc.append('Lounges')
        sub_loc.append(match['Sub-Location'].values[0])
        updated_name.append(match['Updated Name'].values[0])
        city.append(match['City'].values[0])
        dep.append('Lounges')
        try:
            cat.append(category_master[category_master['Item Code'] == row['SKU']]['Category'].values[0])
        except:
            cat.append('')
        uom.append(row['Measuring Unit'])
        try:
            qty.append(float(row['Audit Quantity']))
        except:
            if 'G' in row['Audit Quantity']:
                qty.append(float(int(row['Audit Quantity'][:-1]) / 1000))
            else:
                assert("Non numerical value in the quantity column!")
    else:
        missing.add(row['Branch Name'])
        con.append('')
        sku_code.append(row['SKU'])
        sku_name.append(row['Item Name'])
        item_type.append(row['Type'])
        loc.append('Lounges')
        sub_loc.append('')
        updated_name.append('')
        city.append('')
        dep.append('Lounges')
        try:
            cat.append(category_master[category_master['Item Code'] == row['SKU']]['Category'].values[0])
        except:
            cat.append('')
        uom.append(row['Measuring Unit'])
        try:
            qty.append(float(row['Audit Quantity']))
        except:
            if 'G' in row['Audit Quantity']:
                qty.append(float(int(row['Audit Quantity'][:-1]) / 1000))
            else:
                assert("Non numerical value in the quantity column!")

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
            if category in ['Cakes & Pastries', 'Bakery', 'Tea Cakes & Muffins', 'Teacake & Muffins']:
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
            if category == 'Cakes & Pastries':
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

final_df.to_excel('Output_Lounge.xlsx', index=False)
print('File Saved!')