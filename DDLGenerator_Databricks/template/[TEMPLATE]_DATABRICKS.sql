/*------------------------------------------------------------------------------------------
* @Script Name: DDL_[table_name]
* @Description: [description]
* @Author: [author]
* @Created Date: [created_date]
* @Version: v1.0
* @Modified Log:
* - [created_date] v1.0 [author]   初始化
*------------------------------------------------------------------------------------------*/

CREATE TABLE IF NOT EXISTS [schema].[table_name] (
    [content]
)
USING DELTA
COMMENT '[table_comment]'
LOCATION '[location]';