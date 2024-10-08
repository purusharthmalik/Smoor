import numpy as np
import pandas as pd
import openpyxl as xl

print("Loading the workbooks...")
workbook = xl.load_workbook(r"C:\Users\purus\Downloads\Bangalore Costcenter july -24.xlsx")
tally = workbook[workbook.sheetnames[0]]

mis_master = pd.read_excel(r"S:\smoor\data\Master MIS.xlsx",
                          skiprows=[0,1])
mis_master.drop_duplicates(subset=['Code'], inplace=True)

# Renaming the columns
print("Renaming the headers...")
mis_names = pd.read_excel(r"S:\smoor\data\names_blr.xlsx")
mis_names['Name'] = mis_names['Name'].apply(lambda x: x.strip())
headers = [c.value for c in next(tally.iter_rows(min_row=3, max_row=3, min_col=13))]
new_headers = []
for header in headers:
    try:
        rename = mis_names[mis_names['Name'] == header.strip()]['Rename']
        new_headers.append(rename.values[0])
    except:
        print(header)
        new_headers.append(header)

# Getting the GL codes and names
print("Extracting the GL codes...")
gl_codes, names = [], []
for col in tally.iter_cols(min_row=4, min_col=5, max_col=5, values_only=True):
    for val in col:
        try:
            gl_code, name = list(map(lambda x: x.strip(), val.split('|')))
            gl_codes.append(int(gl_code))
            names.append(name)
        except:
            gl_codes.append(val)
            names.append(val)

# Getting the Party A/c
print("Extracting the Party A/c...")
for col in tally.iter_cols(min_row=4, min_col=6, max_col=6, values_only=True):
    party_ac = col

# Getting the vch and led
print("Extracting the VCH and LED Narration...")
vch_narration, led_narration = [], []
for col_idx, col in enumerate(tally.iter_cols(min_row=4, min_col=11, max_col=12, values_only=True), start=11):
    if col_idx == 11:
        vch_narration.extend(col)
    elif col_idx == 12:
        led_narration.extend(col)

# Getting columns F-L
print("Extrating the master columns...")
f, g, h, i, j, k, l = [], [], [], [], [], [], []
lists = [f, g, h, i, j, k, l]
for code in gl_codes:
    try:
        vals = mis_master[mis_master['Code'] == code].values[0][4:-1]
        for val, list_ in zip(vals, lists):
            list_.append(val)
    except:
        for list_ in lists:
            list_.append(None)

# Getting the cell values
print("Populating the rest of the sheet...")
temp = []
for idx, col in enumerate(tally.iter_cols(min_row=4, min_col=13)):
    # Getting the column name
    col_name = new_headers[idx]
    for cell in col:
        if cell.value == None:
            temp.append(0)
        elif cell.number_format[-3:-1] == "Dr":
            temp.append(cell.value)
        elif cell.number_format[-3:-1] == "Cr":
            temp.append(-1*int(cell.value))
    if idx == 0:
        value_df = pd.DataFrame({col_name: temp})
    else:
        value_df[col_name] = temp
    temp = []

# Creating the total column
value_df['Total'] = value_df.sum(axis=1)

master_cols = pd.DataFrame(np.array([gl_codes, names, party_ac, vch_narration, led_narration, f, g, h, i, j, k, l]).T,
                           columns=['GL Code', 'Name', 'Party A/c', 'Vch Narration', 'Led Narration', 'A/c Grp 5', 'A/c Grp 4', 'A/c Grp3 (Alloc)', 'A/c Grp 2 (MIS)', 'A/c Grp 1', 'Verical Grp', 'Final Grp in P&L'])
final_df = pd.concat([master_cols, value_df],
                     axis=1)
final_df.to_csv("Generated BLR Final.csv", index=False)
print("Files saved!")