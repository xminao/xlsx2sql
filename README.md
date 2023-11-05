# xlsx2sql
xlsx file to sql file



1. 判断文件格式是否与模板一致（判断列名）
2. 判断数据如果active_status有值，而关键位（schema, table_name, column_name等）没有值，则记作非法行，不作为转换数据。如果数据合法，有效位为0，记录为失效数据
3. 非法的情况：非空列不是'Y'或者'N'；主键约束列的值不是'Y'或者'N'或者None；缺重要列；主键约束和非空约束冲突；