import numpy as np
import pandas as pd
import openpyxl as xl

workbook = xl.load_workbook(r"S:\smoor\data\cost center Blr Aug -24.xlsx")
tally = workbook[workbook.sheetnames[0]]

mis_master = pd.read_excel(r"S:\smoor\data\Master MIS.xlsx",
                          skiprows=[0,1])
mis_master.drop_duplicates(subset=['Code'], inplace=True)

# Getting the GL codes and names
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
for col in tally.iter_cols(min_row=4, min_col=6, max_col=6, values_only=True):
    party_ac = col

# Getting the vch and led
vch_narration, led_narration = [], []
for col_idx, col in enumerate(tally.iter_cols(min_row=4, min_col=11, max_col=12, values_only=True), start=11):
    if col_idx == 11:
        vch_narration.extend(col)
    elif col_idx == 12:
        led_narration.extend(col)

# Getting columns F-L
f, g, h, i, j, k, l = [], [], [], [], [], [], []
for code in gl_codes:
    try:
        vals = mis_master[mis_master['Code'] == code].values[0][4:-1]
        for val, list_ in zip(vals, [f, g, h, i, j, k, l]):
            list_.append(val)
    except:
        for list_ in [f, g, h, i, j, k, l]:
            list_.append(None)

df = pd.DataFrame(np.array([gl_codes, names, party_ac, vch_narration, led_narration, f, g, h, i, j, k, l]).T)
df.to_csv("sample.csv", index=False)