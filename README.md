# VBA_ORM
The m_SchemaBuilder.bas file is a VBA module that maps an Access database (mdb/accdb) and creates a class module for each database table.
Each mapped table contains properties that return the table's column names as a string.

The QueryDesigner class modules can be used to return a properly formatted string for use as an SQL query, thus avoiding working directly with strings in your code.
