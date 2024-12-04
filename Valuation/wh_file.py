import pandas as pd

print("Loading dumps...")
dumps = [
    "C:/Users/purus/Documents/GitHub/Smoor/Valuation/Closing Files Nov 24/Warehouse/Closing Stock Nov 24 Warehouse - Bangalore .xlsx",
    "C:/Users/purus/Documents/GitHub/Smoor/Valuation/Closing Files Nov 24/Warehouse/Closing Stock Nov 24 Warehouse - Chennai.xlsx",
    "C:/Users/purus/Documents/GitHub/Smoor/Valuation/Closing Files Nov 24/Warehouse/Closing Stock Nov 24 Warehouse - Gurgaon 2nd.xlsx",
    "C:/Users/purus/Documents/GitHub/Smoor/Valuation/Closing Files Nov 24/Warehouse/Closing Stock Nov 24 Warehouse - Mumbai .xlsx",
    "C:/Users/purus/Documents/GitHub/Smoor/Valuation/Closing Files Nov 24/Warehouse/Closing Stock Nov 24 Warehouse - Pune .xlsx"
]

columns=['Item', 'Item Name', 'Warehouse', 'UOM', 'SOH', 'City', 'Valuation Rate']
final_dump = pd.DataFrame(columns=columns)

print("Formatting dumps...")
for dump in dumps:
    df = pd.read_excel(dump)[columns]
    final_dump = pd.concat([final_dump, df])

final_dump = final_dump[final_dump['SOH'] != 0]
final_dump.dropna(subset=['SOH'],
                  inplace=True)

item_type = []
leave = []
for idx, code in enumerate(final_dump['Item'].values):
    if code.startswith('RM'):
        item_type.append('RM')
    elif code.startswith('PKG'):
        item_type.append('PM')
    elif code.startswith('FG'):
        item_type.append('FG')
    elif code.startswith('SFG'):
        item_type.append('SFG')
    elif code.startswith('CM') or code.startswith('PS'):
        leave.append(idx)
    else:
        item_type.append('FG')

idxs = list(range(final_dump.shape[0]))
for _ in leave:
    idxs.remove(_)

final_dump = final_dump.iloc[idxs]
final_dump.reset_index(drop=True,
                       inplace=True)

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
    'Qty',
    'Rate'])

con, loc, sub_loc, updated_name, rate, val_rate = [], [], [], [], [], []

print("Filling in the item details...")
for row in final_dump.iterrows():
    idx, row_data = row
    con.append(row_data['City'] + row_data['Item'])
    loc.append('Warehouse')

    # Sub Location rules
    if row_data['Warehouse'].startswith('HAL'):
        sub_loc.append('HAL')
    elif row_data['Warehouse'].startswith('Jigani'):
        sub_loc.append('Jigani')
    elif row_data['Warehouse'].startswith('Chennai'):
        sub_loc.append('Chennai')
    elif row_data['Warehouse'].startswith('Pune'):
        sub_loc.append('Pune')
    else:
        sub_loc.append(row_data['Warehouse'])
    updated_name.append(row_data['City'])
    rate.append(row_data['Valuation Rate'])

temp_df = pd.DataFrame({
    'Con': con,
    'SKU Code': final_dump['Item'],
    'SKU Name': final_dump['Item Name'],
    'Item Type': item_type,
    'Location': loc,
    'Sub-Location': sub_loc,
    'Updated Name': updated_name,
    'City': final_dump['City'],
    'Department': loc,
    'Category': loc,
    'UOM': final_dump['UOM'],
    'Qty': final_dump['SOH'],
    'Rate': rate
})
final_df = pd.concat([final_df, temp_df])

if final_df['Rate'].isna().shape[0] != 0:
    print("Getting the missing valuation rates...")
    print("Loading the smoor sheet...")
    rate_master = r"C:\Users\purus\Documents\GitHub\Smoor\Valuation\Data\Rate Master till Nov.xlsx"

    smoor_sheet = pd.read_excel(rate_master,
                                sheet_name=['Smoor Product'],
                                header=1)['Smoor Product'][['FG Code', 'FnP (At Factory Level)']]
    
    print("Loaded. Filling in the sheet...")
    for row_data in final_df[final_df['Rate'].isna()].iterrows():
        idx = row_data[0]
        row = row_data[1]
        try:
            final_df.loc[idx, 'Rate'] = smoor_sheet[smoor_sheet['FG Code'] == row['SKU Code']]['FnP (At Factory Level)'].values[0]
        except:
            final_df.loc[idx, 'Rate'] = 0
            
final_df['Valuation Rate'] = final_df['Qty'] * final_df['Rate']

final_df.to_excel('Output_WH.xlsx', index=False)
print("File saved!")