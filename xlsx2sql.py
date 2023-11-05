import openpyxl
import sys

# 提示用户输入xlsx文件名
xlsx_file = input("请输入包含表结构信息的xlsx文件名（包括文件扩展名）：")

# 提示用户输入要读取的工作表名称
#sheet_name = input("请输入要读取的工作表名称：")

# 读取xlsx文件
wb = openpyxl.load_workbook(xlsx_file)
#ws = wb[sheet_name]
ws = wb.active

# 读取DB_Design_TableColumn的模板文件列名
xlsx_template_table_column = openpyxl.load_workbook('[TEMPLATE]_DB_Design.xlsx')['TableColumn_XXX']
template_title = []
for row in xlsx_template_table_column.iter_rows(min_row=2, max_row=2, values_only=True):
    for col in row:
        template_title.append(col)
print('模板列名:')
print(template_title)

# 读取输入文件的文件列名
title = []
for row in ws.iter_rows(min_row=2, max_row=2, values_only=True):
    for col in row:
        title.append(col)
print('输入文件列名:')
print(title)

# 对比模板数据
if template_title != title:
    print("输入文件与模板不一致")
    sys.exit()

# 读取表结构数据
data = []
for row in ws.iter_rows(min_row=2, values_only=True):
    schema_user, table_name, column_name, data_type, nullable, db_pk, column_comments, business_pk, column_description, sample_data, StatusReason, active_status = row
    data.append((schema_user, table_name, column_name, data_type, nullable, db_pk, column_comments, business_pk, column_description, sample_data, StatusReason, active_status))

print('输入表数据：')
print(data)

# 读取表中schema, table_name, column
# {'dbo': {'table1': [['col1', ['Y', 'N', 'COMMENT']], ['col2', ['Y', 'N', 'COMMENT']]]}}
schema_dict = {}
for row in data:
    # if schmea not in dict, init
    schema = row[0]
    table_dict = schema_dict.get(schema, {})
    table = row[1]
    col_list = table_dict.get(table, [])
    col_attr = [row[3], row[4], row[5], row[6]]
    col = [row[2], col_attr]
    # update col_list
    col_list.append(col)
    # update table_dict
    table_dict[table] = col_list
    # update schema_dict
    schema_dict[schema] = table_dict
print("schema_dict:")
print(schema_dict)

# 读取表中schema, table_name, column
# ['dbo':['table_1':['column1':['Y', 'N', 'COMMENT']]]]
# list = {}
# for row in data:
#     print(row[0], row[1], row[2])
#     schema = list.get(row[0], {})
#     list[row[0]] = schema

#     table = list[row[0]].get(row[1], {})
#     list[row[0]][row[1]] = table

#     col_attri = [row[3], row[4], row[5], row[6]]
#     list[row[0]][row[1]][row[2]] = col_attri
#     # list[row[0]][row[1]][row[2]] = col_attri
# print(list['dbo'])
# for tb_k, tb_v in list['dbo'].items():
#     print(f"表：{tb_k}，列：")
#     for col_k, col_v in tb_v.items():
#         print(f"- {col_k} : {col_v}")

# 读取SQL模板文件
with open('template_sql.sql', 'r') as template_file:
    sql_template = template_file.read()

# 生成SQL脚本
sql_script = ''

for sch_k, sch_v in schema_dict.items():
    print(f"schema:{sch_k}")
    for tb_k, tb_v in sch_v.items():
        print(f"- table:{tb_k}")
        script = ''
        content = ''
        pk = 'PRIMARY KEY ('
        for index, col in enumerate(tb_v):
            print(f"\t- {index} col:{col}")
            nullable = ''
            if index != len(tb_v) - 1:
                nullable = ',' if col[1][1] == 'Y' else ' NOT NULL,'
            else:
                nullable = '' if col[1][1] == 'Y' else ' NOT NULL'
            comment = ''
            if col[1][3] is not None:
                comment = '' if col[1][3] == '' else ' -- ' + col[1][3]
            content += col[0] + ' ' + col[1][0] + nullable + comment + '\n\t'
            if col[1][2] == 'Y':
                pk += col[0] + ', '
        pk += ')'
        script = sql_template.replace('[schema_user]', '[' + sch_k + ']' or '')\
                                .replace('[table_name]', '[' + tb_k + ']' or '')\
                                .replace('[content]', content)\
                                .replace('[primary_key]', pk)\
                                .replace(', )', ')')
        sql_script += script + '\n\n'

# for row in data:
#     schema_user, table_name, column_name, data_type, nullable, db_pk, column_comment = row
#     sql_script += sql_template.replace('[schema_user]', schema_user or '')\
#                                 .replace('[table_name]', table_name or '')\
#                                 .replace('[column_name]', column_name or '')\
#                                 .replace('[data_type]', data_type or '')\
#                                 .replace('[nullable]', 'NULL' if nullable == 'Y' else 'NOT NULL')\
#                                 .replace('[db_pk]', 'PRIMARY KEY' if db_pk == 'Y' else '')\
#                                 .replace('[column_comment]', f'-- {column_comment}' if column_comment else '')\
#                                 + '\n\n'

# 提示用户输入输出SQL文件名
output_sql_file = input("请输入要保存生成的SQL脚本的文件名（包括文件扩展名）：")

# 将SQL脚本保存到文件
with open(output_sql_file, 'w', encoding='utf-8') as sql_file:
    sql_file.write(sql_script)

print(f"SQL脚本已生成并保存到{output_sql_file}")
