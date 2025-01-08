# 1. Intro
This code presents a template to dynamically extract tables from Excels stored in Sharepoint onto SQL Server tables, using a SQL Server Integration Services (SSIS) package.

We here assume the user has access to an OLEDB connection to an SQL Server database. Note that this connection is absent in this template, so you need to create it.

The extraction automation includes a retry mechanism, which premise may be useful in other SSIS applications, and it is constructed in two main procedures written as C# scripts:

- Sharepoint/Excel->Local Server extraction (output path defined in local package variables):
  
![image](https://github.com/user-attachments/assets/b4602751-2c75-424c-b3cf-d36aea69623c)

- Local Server->SQL Server extraction:
  
![image](https://github.com/user-attachments/assets/7715dbd1-7b85-4b78-909b-251d9415d8e9)


The extraction of the Excel tables is within a foreach with the following parameters:

- Path to the Source Excel within Sharepoint 
- Sheet name where the table is stored in the source Sharepoint Excel
- Destination Schema name the user wants
- Destination table name the user wants

The parameters above are given from a query to a parameter table in SQL Server:

![image](https://github.com/user-attachments/assets/739f2c34-1d62-41cf-a7a8-ba7b02c5e1f3)

![image](https://github.com/user-attachments/assets/b034b196-d495-480a-b554-dcd52f4e2239)

But of course, the user may change the package so that the way these inputs are processed differ. This is just a useful example.

If this package value is set to True:

![image](https://github.com/user-attachments/assets/c549b6ef-49b3-4d34-b7c4-e2c7795b72b7)

the package will retry up until 1000 times (you can change this threshold under MaximumErrorCount properties), otherwise it will retry up until the value set in the RetryCount value:

![image](https://github.com/user-attachments/assets/e35587c1-6e4d-4096-98e2-bf27ffb7007a)

