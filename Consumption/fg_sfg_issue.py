import numpy as np
import pandas as pd
import glob

print("Reading all the secondary files...")
fg_sfg_issue = pd.read_csv("C:/Users/purus/Documents/GitHub/Smoor/Consumption/FG-SFG(FG-SFG).csv", encoding='latin1')[['Item Code', 'Type', 'Category', 'Grammage', 'Over head Cost']]
name_master = pd.read_excel("C:/Users/purus/Documents/GitHub/Smoor/Valuation/Data/Data for Closing Valuations.xlsx", sheet_name="Updated Name and city")
category_master = pd.read_excel("C:/Users/purus/Documents/GitHub/Smoor/Valuation/Data/Data for Closing Valuations.xlsx", sheet_name="Category")

print("Loading the smoor and b2b sheet...")
rate_master = "C:/Users/purus/Documents/GitHub/Smoor/Valuation/Data/Rate Master till Nov.xlsx"
smoor_sheet = pd.read_excel(rate_master,
                            sheet_name=['Smoor Product'],
                            header=1)['Smoor Product'][['FG Code', 'FnP (At Factory Level)']]
b2b_sheet = pd.read_excel(rate_master,
                          sheet_name=['B2B Products'],
                          header=1)['B2B Products'][['FG Code', 'FnP (At Factory Level)']]
mum_pr_master = pd.read_excel(rate_master,
                            sheet_name=['Mumbai PR'],
                            header=1)['Mumbai PR'][['Item Code', 'Rate']]

print("Loading the FG-SFG sheet...")
fg_sfg_sheet = pd.read_excel(rate_master,
                          sheet_name=['FG-SFG'])['FG-SFG'][['SFG Code', 'Yield']]

# Loading lounge fulfillment reports
print("Loading the lounge fulfillment reports...")
path = "C:/Users/purus/Documents/GitHub/Smoor/Consumption/ERP DUMPS Sept 2024/Lounge order fullfilment report SEPT 2024/*.xlsx"
all_files = glob.glob(path)

df_list = [pd.read_excel(file) for file in all_files]
lf_reports = pd.concat(df_list, ignore_index=True)[['Lounge', 'Item Code', 'Item Name', 'Weight Per Unit',	'UOM', 'Rate', 'Ordered Qty', 'Ordered Amount', 'Delivered Qty', 'Delivered Amount', 'Received Qty', 'Reason', 'Order Date', 'Deliver Date', 'Purchase Order']]

updated_name, actual_qty, new_rate, total_fnp, city, typ, category, grammage, overhead_cost, ohv, ftp_per_unit, ftp_value = [], [], [], [], [], [], [], [], [], [], [], []
for idx, row in lf_reports.iterrows():
    name_match = name_master[name_master['Sub-Location'].apply(lambda x: x.strip().lower()) == row['Lounge'].strip().lower()]
    # Remaining columns
    code = row['Item Code']
    type_match = fg_sfg_issue[fg_sfg_issue['Item Code'] == code]
    if len(type_match) != 0:
        # Updated name and city
        if len(name_match) != 0:
            updated_name.append(name_match['Updated Name'].values[0])
            city_name = name_match['City'].values[0]
            city.append(city_name)

        else:
            updated_name.append(None)
            city.append(None)
        # Actual Quantity
        try:
            assert row['Reason'].isalpha()
            print(row['Reason'])
            aq = float(row['Received Qty'])
            actual_qty.append(aq)
        except:
            aq = float(row['Delivered Qty'])
            actual_qty.append(aq)
        # Type
        item_type = type_match['Type'].values[0]
        typ.append(item_type)
        # Category, Grammage, Overhead Cost
        temp_cat = type_match['Category'].values[0]
        category.append(temp_cat)
        temp_gram = type_match['Grammage'].values[0]
        if type(temp_gram) == str:
            try:
                temp_gram = float(temp_gram.strip().replace(',', ''))
            except:
                temp_gram = 0
        else:
            temp_gram = float(temp_gram)
        if temp_gram == None or temp_gram == np.nan:
           grammage.append(0)
        else:
            grammage.append(temp_gram)
        try:
            temp_ohc = float(type_match['Over head Cost'].values[0])
        except:
            temp_ohc = 0
        if temp_ohc == None or temp_ohc == np.nan:
           overhead_cost.append(0)
        else:
            overhead_cost.append(temp_ohc)

        # OHV
        print(temp_gram, temp_ohc, aq)
        ohv.append(temp_gram * temp_ohc * aq)
        # New rate
        if item_type == 'FG':
            if city_name == 'Bangalore':
                try:
                    new_rate.append(smoor_sheet[smoor_sheet['FG Code'] == code]['FnP (At Factory Level)'].values[0])
                except:
                    try:
                        new_rate.append(b2b_sheet[b2b_sheet['FG Code'] == code]['FnP (At Factory Level)'].values[0])
                    except:
                        new_rate.append(0)
            elif city_name in ['Mumbai', 'Pune']:
                if temp_cat in ['Cakes & Pastries', 'Bakery', 'Tea Cakes & Muffins', 'Teacake & Muffins']:
                    try:
                        new_rate.append(smoor_sheet[smoor_sheet['FG Code'] == code]['FnP (At Factory Level)'].values[0])
                    except:
                        try:
                            new_rate.append(b2b_sheet[b2b_sheet['FG Code'] == code]['FnP (At Factory Level)'].values[0])
                        except:
                            new_rate.append(0)
                else:
                    try:
                        new_rate.append(mum_pr_master.loc[mum_pr_master['Item Code'] == code]['Rate'].values[0])
                    except:
                        new_rate.append(0)
            else:
                if temp_cat == 'Cakes & Pastries':
                    try:
                        new_rate.append(smoor_sheet[smoor_sheet['FG Code'] == code]['FnP (At Factory Level)'].values[0])
                    except:
                        try:
                            new_rate.append(b2b_sheet[b2b_sheet['FG Code'] == code]['FnP (At Factory Level)'].values[0])
                        except:
                            new_rate.append(0)
                else:
                    try:
                        new_rate.append(mum_pr_master.loc[mum_pr_master['Item Code'] == code]['Rate'].values[0])
                    except:
                        new_rate.append(0)
        elif item_type == 'SF' or item_type == 'SFG':
            new_rate.append(fg_sfg_sheet[fg_sfg_sheet['SFG Code'] == code]['Yield'].sum())
        else:
            new_rate.append(0)
        # FTP per unit
        try:
            ftp_per_unit.append(new_rate[-1] + (ohv[-1] / aq))
        except:
            ftp_per_unit.append(0)
        # FTP value
        ftp_value.append(ftp_per_unit[-1] * aq)
        # Total FNP value
        total_fnp.append(aq * new_rate[-1])
    else:
        typ.append(None)
    
# Adding the columns to dataframe
lf_reports['Type'] = typ
lf_reports.dropna(subset=['Type'], inplace=True)
lf_reports['Updated Name'] = updated_name
lf_reports['Actual Qty'] = actual_qty
lf_reports['New Rate'] = new_rate
lf_reports['Total FNP Value'] = total_fnp
lf_reports['City'] = city
lf_reports['Category'] = category
lf_reports['Grammage'] = grammage
lf_reports['Over head Cost'] = overhead_cost
lf_reports['Over head Value'] = ohv
lf_reports['FTP per unit'] = ftp_per_unit
lf_reports['FTP Value'] = ftp_value

lf_reports.to_excel('FG-SFG_Consolidated.xlsx', index=False)
print("File saved!")