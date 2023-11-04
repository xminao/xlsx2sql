import openpyxl

# 提示用户输入xlsx文件名
xlsx_file = input("请输入包含表结构信息的xlsx文件名（包括文件扩展名）：")

# 读取xlsx文件
wb = openpyxl.load_workbook(xlsx_file)
ws = wb.active

# 读取表结构数据
data = []
for row in ws.iter_rows(min_row=2, values_only=True):
    schema_user, table_name, column_name, data_type, nullable, db_pk, column_comment = row
    data.append((schema_user, table_name, column_name, data_type, nullable, db_pk, column_comment))

print('表数据：')
print(data)

# 读取SQL模板文件
with open('sql_template.sql', 'r') as template_file:
    sql_template = template_file.read()

# 生成SQL脚本
sql_script = ''
for row in data:
    schema_user, table_name, column_name, data_type, nullable, db_pk, column_comment = row
    sql_script += sql_template.replace('[schema_user]', schema_user or '')\
                                .replace('[table_name]', table_name or '')\
                                .replace('[column_name]', column_name or '')\
                                .replace('[data_type]', data_type or '')\
                                .replace('[nullable]', 'NULL' if nullable == 'Y' else 'NOT NULL')\
                                .replace('[db_pk]', 'PRIMARY KEY' if db_pk == 'Y' else '')\
                                .replace('[column_comment]', f'-- {column_comment}' if column_comment else '')\
                                + '\n\n'

# 提示用户输入输出SQL文件名
output_sql_file = input("请输入要保存生成的SQL脚本的文件名（包括文件扩展名）：")

# 将SQL脚本保存到文件
with open(output_sql_file, 'w') as sql_file:
    sql_file.write(sql_script)

print(f"SQL脚本已生成并保存到{output_sql_file}")
