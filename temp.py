import pandas as pd

df = pd.read_excel('temp.xlsx')

print(df['Generated'].value_counts()[df['Generated'].value_counts().values != df['Previous'].value_counts().values])