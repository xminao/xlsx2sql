'''
-*- coding: utf-8 -*-
@File: xlsx2sql.py
@Author: Minhao
@Desc: 用于根据DB_Design模板文件生成对应建表语句
@Time: 2023/11/04

'''

import openpyxl
import sys
import os
import time
import yaml

# 生成的脚本头内容
TITLE_AUTHOR = ''

# 模板文件
TEMPLATE_DB_DESIGN = ''
TEMPLATE_SCRIPT = ''
TEMPLATE_MSSQL = ''
TEMPLATE_MYSQL = ''
TEMPLATE_ORACLE = ''
TEMPLATE_DATABRICKS = ''

# 输入输出目录
INPUT_PATH = '.\\input'
OUTPUT_PATH = '.\\output'

# 输入文件后缀，用于校验
INPUT_EXTENSION = '_DB_Design.xlsx'

# 输出日志(待完成)
LOG_PATH = '.\\logs'
LOG_FORMAT = '%Y%m%d-%H%M%S'


def load_config():
    """
        读取配置文件，初始化全局变量
    """
    # 读取 config 配置文件
    with open('./config.yml', encoding='UTF-8') as config_file:
        config_data = yaml.safe_load(config_file)

    # 脚本使用者
    global TITLE_AUTHOR
    TITLE_AUTHOR = config_data['author']
    # 校验的后缀名
    global INPUT_EXTENSION
    INPUT_EXTENSION = config_data['file_extension']
    # 模板信息
    global TEMPLATE_DB_DESIGN
    TEMPLATE_DB_DESIGN = f".\\template\\{config_data['templates']['db_design']}"
    global TEMPLATE_SCRIPT
    TEMPLATE_SCRIPT = f".\\template\\{config_data['templates']['sql_script']}"
    global TEMPLATE_MSSQL
    TEMPLATE_MSSQL = f".\\template\\{config_data['templates']['mssql']}"
    global TEMPLATE_MYSQL
    TEMPLATE_MYSQL = f".\\template\\{config_data['templates']['mysql']}"
    global TEMPLATE_ORACLE
    TEMPLATE_ORACLE = f".\\template\\{config_data['templates']['oracle']}"
    global TEMPLATE_DATABRICKS
    TEMPLATE_DATABRICKS = f".\\template\\{config_data['templates']['databricks']}"


def check_file_format(file_path: str):
    """
        校验文件的格式与模板是否一致：
        1. 列名以及顺序是否一致

        file_path: 输入文件的名（包括路径）
    """

    # 确认输入文件存在 Table_List页 和 Column_List页
    input_wb = openpyxl.load_workbook(file_path)
    if 'Table_List' not in input_wb.sheetnames:
        # raise Exception(f"请确认 {file_path} 存在 Table_List 和 Column_List Sheet页.")
        raise Exception(f"Table_List Sheet页不存在.")

    if 'Column_List' not in input_wb.sheetnames:
        # raise Exception(f"请确认 {file_path} 存在 Table_List 和 Column_List Sheet页.")
        raise Exception(f"Column_List Sheet页不存在.")

    # 验证Table_List列名以及顺序
    # 1. 读取模板 DB_Design.xlsx 的 Table_List页 列名
    template_table_ws = openpyxl.load_workbook(TEMPLATE_DB_DESIGN)['Table_List']
    template_table_ws_title = []
    for row in template_table_ws.iter_rows(min_row=2, max_row=2, values_only=True):
        for col in row:
            template_table_ws_title.append(col)

    # 2. 读取输入文件的 Table_List页 列名
    input_table_ws = openpyxl.load_workbook(file_path)['Table_List']
    input_table_ws_title = []
    for row in input_table_ws.iter_rows(min_row=2, max_row=2, values_only=True):
        for col in row:
            input_table_ws_title.append(col)

    # 3. 判断输入文件列名以及顺序是否与模板一致
    if template_table_ws_title != input_table_ws_title:
        # raise Exception(f"输入文件 {file_path} 的 Table_List 页格式与模板不一致.")
        raise Exception(f"Table_List 页格式与模板不一致.")

    # 验证Column_List列名以及顺序
    # 1. 读取模板 DB_Design.xlsx 的 Column_List页 列名
    template_column_ws = openpyxl.load_workbook(TEMPLATE_DB_DESIGN)['Column_List']
    template_column_ws_title = []
    for row in template_column_ws.iter_rows(min_row=2, max_row=2, values_only=True):
        for col in row:
            template_column_ws_title.append(col)

    # 2. 读取输入文件的 Column_List页 列名
    input_column_ws = openpyxl.load_workbook(file_path)['Column_List']
    input_column_ws_title = []
    for row in input_column_ws.iter_rows(min_row=2, max_row=2, values_only=True):
        for col in row:
            input_column_ws_title.append(col)

    # 3. 判断输入文件列名以及顺序是否与模板一致
    if template_column_ws_title != input_column_ws_title:
        # raise Exception(f"输入文件 {file_path} 的 Column_List 页格式与模板不一致.")
        raise Exception(f"Column_List 页格式与模板不一致.")


# def check_file_content(file_path: str):
#     """
#         检验输入文件内容：
#         1. Table_List的数据有效位有没有非法的（非数值）
#         2. Table_List有效的数据关键列是不是空（数据库名称，数据库，schema，表名）
#         3. Table_List有效的数据表名在Column_List表有没有有效的合法数据字段
#         file_path: 输入文件的名（包括路径）
#     """
#     # TODO 检查内容
#     # 读取DB_Design_TableList页
#     input_table_list_ws = openpyxl.load_workbook(file_path)['Table_List']
#     input_table_names = set()  # format: schema.table_name
#     for row_num, row in enumerate(input_table_list_ws.iter_rows(min_row=3, values_only=True), start=3):
#         _, _, _, schema_user, _, table_name, _, _, _, _, _, active_status = row
#         if all(cell is None for cell in row):
#             continue
#         if schema_user is None or table_name is None:
#             raise Exception(f"页 Table_List, 行 {row_num} 未定义 schema_user 或 table_name.")
#         input_table_names.add(schema_user + '.' + table_name)
#
#     # 读取DB_Design_Column_List页
#     input_table_column_ws = openpyxl.load_workbook(file_path)['Column_List']
#     for row_num, row in enumerate(input_table_column_ws.iter_rows(min_row=3, values_only=True), start=3):
#         schema_user, table_name, _, _, _, _, _, _, _, _, _, _ = row
#         if all(cell is None for cell in row):
#             continue
#         if schema_user is None or table_name is None:
#             raise Exception(f"页 Column_List, 行 {row_num} 未定义 schema_user 或 table_name.")
#         table_name = schema_user + '.' + table_name
#         if table_name not in input_table_names:
#             raise Exception(f"行 {row_num} 表名 {table_name} 未在 Table_List 中定义.")


def generate_sql_script(schema_dict: dict, table_comment: dict, filename: str, selected_template):
    """
        根据生成的schema字典数据集，生成SQL脚本
        schema_dict: 传入的字典数据集
    """
    # 读取SQL模板文件
    with open(selected_template[1], 'r') as template_file:
        sql_template = template_file.read()

    # {
    #     'schema1': {
    #         'table1': [['col1', ['int', 'Y', 'N', 'COMMENT']],
    #                    ['col2', ['int', 'Y', 'N', 'COMMENT']],
    #                    ...],
    #         'table2': [['col1', ['int', 'Y', 'N', 'COMMENT']],
    #                    ['col2', ['int', 'Y', 'N', 'COMMENT']],
    #                    ...],
    #         ...
    #     },
    #     'schema2': {
    #         'table1': [['col1', ['int', 'Y', 'N', 'COMMENT']],
    #                    ['col2', ['int', 'Y', 'N', 'COMMENT']],
    #                    ...],
    #         'table2': [['col1', ['int', 'Y', 'N', 'COMMENT']],
    #                    ['col2', ['int', 'Y', 'N', 'COMMENT']],
    #                    ...],
    #         ...
    #     },
    #     ...
    # }

    # 生成的SQL脚本
    sql_script = ''

    # 遍历schema字典
    for sch_k, sch_v in schema_dict.items():
        print(f"{sch_k}")
        # 遍历该schema 下的table字典
        for tb_k, tb_v in sch_v.items():
            content = ''  # 字段相关的内容
            pk_content = 'PRIMARY KEY('  # 主键内容
            pk_count = 0  # 主键数量
            for index, col in enumerate(tb_v):
                # 遍历该table下的所有col数组
                nullable = ''  # NOT NULL约束语句
                if index != len(tb_v) - 1:
                    nullable = ',' if col[1][1] == 'Y' or col[1][1] is None else ' NOT NULL,'
                else:
                    nullable = '&end' if col[1][1] == 'Y' or col[1][1] is None else ' NOT NULL&end'

                comment = ''  # 注释
                if col[1][3] is not None:
                    comment = '' if col[1][3] == '' else ' -- ' + col[1][3]

                # 追加content
                if index != len(tb_v) - 1:
                    content += col[0] + ' ' + col[1][0] + nullable + comment + '\n\t'
                else:
                    content += col[0] + ' ' + col[1][0] + nullable + comment

                if col[1][2] == 'Y':  # 是否是主键
                    pk_count += 1
                    pk_content += col[0] + ', '  # 主键内容加字段

            pk_content += ')'  # 给主键括号封起来

            # 开始根据建表模板替换内容,生成该表的建表语句
            script = sql_template.replace('[schema_user]', sch_k or '') \
                .replace('[table_name]', tb_k or '') \
                .replace('[content]', content) \

            if pk_count != 0:
                # 如果有主键就替换内容
                script = script.replace('[primary_key]', pk_content) \
                    .replace(', )', ')') \
                    .replace('&end', ',')  # 有主键就把最后一个字段的结尾换为逗号
            else:
                # 没有主键
                script = script.replace('[primary_key]', '') \
                    .replace('\t\n);', ');') \
                    .replace('&end', '')

            # 添加表注释
            #print(f"{sch_k}.{tb_k}")
            schema_table = f"{sch_k}.{tb_k}"
            script = script.replace('[table_comment]', table_comment.get(schema_table, f"表名：{tb_k}"))

            sql_script += script + '\n\n'

    # 输出的SQL文件名
    output_sql_file = OUTPUT_PATH + '.\\' + os.path.splitext(filename)[0] + '.sql'

    # 将创建表的SQL脚本保存到文件
    # with open(output_sql_file, 'w', encoding='utf-8') as sql_file:
    #     sql_file.write(sql_script)

    # print(f"SQL脚本已生成并保存到 {output_sql_file}")

    # 生成一个SCRIPT（带大注释头）
    # 读取SQL模板文件
    with open(TEMPLATE_SCRIPT, 'r', encoding='utf-8') as template_script:
        script_template = template_script.read()

    sql_scripts = ''
    sql_scripts = script_template.replace('[script]', sql_script) \
        .replace('[script_created_date]', time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())) \
        .replace('[script_content]', f"根据 {filename} 生成的SQL建表脚本") \
        .replace('[script_db]', f"{selected_template[0]}") \
        .replace('[script_author]', TITLE_AUTHOR) \
        .replace('[script_version]', "v1.0")

    with open(output_sql_file, 'w', encoding='utf-8') as sql_file:
        sql_file.write(sql_scripts)

    print(f"    SQL脚本已生成并保存到 {output_sql_file}")
    return


def load_and_generate():
    """
        加载所有的DB_Design文件并生成脚本
    """

    # 初始化本批次Log文件
    time_str = time.strftime('%Y%m%d%H%M%S', time.localtime())
    file_log_path = os.path.join(LOG_PATH, f"{time_str}.txt")

    # 遍历 ./input 目录，读取所有的 DB_Design 文件，校验所有格式有效的文件和无效的文件
    file_valid = []
    file_invalid = []
    for filename in os.listdir(INPUT_PATH):
        # 构建完整的文件路径
        file_path = os.path.join(INPUT_PATH, filename)
        if os.path.isfile(file_path) and file_path.endswith(INPUT_EXTENSION):
            # 校验文件格式是否与模板一致
            try:
                check_file_format(file_path)
                file_valid.append(f"{filename}")
            except Exception as ex:
                # print(f"文件格式校验错误，错误原因: {ex}")
                file_invalid.append(f"{filename}，校验失败原因：{ex}")

    print("\n    有效文件列表：")
    for file in file_valid:
        print(f"    * {file}")
    print("\n    无效文件列表：")
    for file in file_invalid:
        print(f"    * {file}")
    print("\n")

    # 如果不存在有效文件，直接关闭脚本
    if len(file_valid) == 0:
        print("    无有效文件，10秒后自动退出脚本.")
        time.sleep(10)
        sys.exit(1)

    # 给用户确认
    input("    按任意键继续执行脚本...")

    # 遍历所有格式校验有效的文件
    for filename in file_valid:
        # 输入当前文件要生成的数据库建表脚本
        print(f"\n    当前文件：{filename}")
        selected_db = input(f"    选择要生成的脚本版本(1: MSSQL 2:MySQL 3:Oracle 4:Databricks，默认为1:MSSQL)：")

        # 判断当前文件要生成的建表语句版本
        selected_template = []
        if selected_db == 2:
            selected_template = ['MySQL', TEMPLATE_MYSQL]
        elif selected_db == 3:
            selected_template = ['Oracle', TEMPLATE_ORACLE]
        elif selected_db == 4:
            selected_template = ['Databricks', TEMPLATE_DATABRICKS]
        else:
            selected_template = ['MSSQL', TEMPLATE_MSSQL]

        # 构建完整的DB_Design文件路径
        file_path = os.path.join(INPUT_PATH, filename)

        # 读取当前输入文件的Column数据
        input_file_ws = openpyxl.load_workbook(file_path)['Column_List']
        # 生成的数据
        valid_data = []
        # 非法数据
        illegal_data = []
        # 失效的数据
        invalid_data = []
        # 遍历 Column_List Sheet页的行，从数据行第三行开始
        for row_num, row in enumerate(input_file_ws.iter_rows(min_row=3, values_only=True), start=3):
            # 读取数据
            no, schema_user, table_name, column_name, data_type, nullable, db_pk, column_comments, \
            business_pk, column_description, sample_data, status_reason, active_status = row

            # 校验，是否能写入数据结构
            if all(cell is None for cell in row):
                # 全空的行直接pass
                pass

            if schema_user is not None and table_name is not None and column_name is not None and data_type is not None:
                # 关键列不能为空，schema, table_name, column_name, data_type
                if nullable is not None and nullable != 'Y' and nullable != 'N':
                    # 非空列nullable值要合法：Y / N / None
                    illegal_data.append((row_num, 'nullable列 值非法，只能为：Y/N/空'))
                elif db_pk is not None and db_pk != 'Y' and db_pk != 'N':
                    # 主键列db_pk值要合法：Y / N / None
                    illegal_data.append((row_num, 'db_pk列 值非法，只能为：Y/N/空'))
                elif active_status != 0:
                    # 数据有效位表示为非0，即有效
                    valid_data.append((no, schema_user, table_name, column_name, data_type, nullable,
                                       db_pk, column_comments, business_pk, column_description, sample_data,
                                       status_reason, active_status))
                else:
                    invalid_data.append(row_num)
            else:
                illegal_data.append((row_num, '关键列(schema_user, table_name, column_name, data_type)存在空值'))

        # print('有效表数据：')
        # for row in data:
        #     print(row)
        # print('非法数据行：')
        # for row in illegal_data:
        #     print(f"- {row}")
        # print('失效数据行：')
        # print(invalid_data)

        # 读取表中schema, table_name, column(包含属性)
        # {'dbo': {'table1': [['col1', ['Y', 'N', 'COMMENT']], ['col2', ['Y', 'N', 'COMMENT']]]}}

        # {
        #     'schema1': {
        #         'table1': [['col1', ['int', 'Y', 'N', 'COMMENT']],
        #                    ['col2', ['int', 'Y', 'N', 'COMMENT']],
        #                    ...],
        #         'table2': [['col1', ['int', 'Y', 'N', 'COMMENT']],
        #                    ['col2', ['int', 'Y', 'N', 'COMMENT']],
        #                    ...],
        #         ...
        #     },
        #     'schema2': {
        #         'table1': [['col1', ['int', 'Y', 'N', 'COMMENT']],
        #                    ['col2', ['int', 'Y', 'N', 'COMMENT']],
        #                    ...],
        #         'table2': [['col1', ['int', 'Y', 'N', 'COMMENT']],
        #                    ['col2', ['int', 'Y', 'N', 'COMMENT']],
        #                    ...],
        #         ...
        #     },
        #     ...
        # }

        # row: (no, schema_user, table_name, column_name, data_type, nullable,
        #       db_pk, column_comments, business_pk, column_description, sample_data,
        #       status_reason, active_status)
        # 所有的schema，字典
        schema_dict = {}
        for row in valid_data:
            # if schema not in dict, init
            schema = row[1]
            # 对应schema的所有table
            table_dict = schema_dict.get(schema, {})

            # if table not in schema dict, init
            table = row[2]
            # 对应table的所有column
            col_list = table_dict.get(table, [])
            col_attr = [row[4], row[5], row[6], row[7]]
            col = [row[3], col_attr]
            # update col_list
            col_list.append(col)
            # update table_dict
            table_dict[table] = col_list
            # update schema_dict
            schema_dict[schema] = table_dict

        # print("生成的数据:")
        # for sch_k, sch_v in schema_dict.items():
        #     print(f'Schema: {sch_k}')
        #     for tb_k, tb_v in sch_v.items():
        #         print(f'- 表名: {tb_k}')
        #         for col in tb_v:
        #             print(f'\t- 列: {col}')

        # 读取Table_List，获取表名对应注释
        table_comment = {}
        table_list_ws = openpyxl.load_workbook(file_path)['Table_List']
        for row in table_list_ws.iter_rows(min_row=3, values_only=True):
            no, server_info, db_name, schema_user, classification, table_name, table_description, \
            table_comments_detail, data_scope, num_rows, status_reason, active_status = row
            if active_status != 0:
                table_comment[f"{schema_user}.{table_name}"] = table_description

        with open(file_log_path, "a") as f:
            f.write(f"file_name: {filename}\n")
            f.write('表列表：\n')
            for tc_k, tc_v in table_comment.items():
                f.write(f"\t- 表名:{tc_k}, 注释:{tc_v}\n")
            f.write("字段列表:\n")
            for sch_k, sch_v in schema_dict.items():
                f.write(f'Schema: {sch_k}\n')
                for tb_k, tb_v in sch_v.items():
                    f.write(f'- 表名: {tb_k}\n')
                    for col in tb_v:
                        f.write(f'\t- 列: {col}\n')
            f.write("\n\n")
            f.close()

        # 生成SQL脚本
        generate_sql_script(schema_dict, table_comment, filename, selected_template)


def main():
    """
        主函数，执行整个过程
    """
    print("""
    =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    =
    =    @Name：xlsx2sql                                  
    =    @Version：v2.1                                    
    =    @Author: Minhao
    =    @Desc: Generate table-creating statement 
    =           by DB_Design.xlsx.
    =
    =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=""")

    # 加载配置文件，初始化全局变量
    load_config()

    # 读取DB_Design文件并生成脚本
    load_and_generate()

    # 暂停控制台，显示信息
    print("\n    生成完毕，10秒后自动关闭控制台...")
    time.sleep(10)


if __name__ == "__main__":
    main()
