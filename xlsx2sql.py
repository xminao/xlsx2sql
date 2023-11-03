import pandas as pd

# 用read_excel函数读取Excel文件
df = pd.read_excel('test.xlsx')

# 获取第一行的列名
column_names = df.columns.tolist()

# 打印列名
print(column_names)