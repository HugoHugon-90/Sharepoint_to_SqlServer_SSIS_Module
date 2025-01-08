# 1. Intro
This code presents a template to dynamically extract tables from Excels stored in Sharepoint onto SQL Server tables, using a SQL Server Integration Services (SSIS) package.

We here assume the user has access to an OLEDB connection to an SQL Server database. Note that this connection is absent in this template, so you need to create it.

The extraction automation includes a retry mechanism, which premise may be useful in other SSIS applications, and it is constructed in two main procedures written as C# scripts:

- Sharepoint/Excel->Local Server extraction:
  
![image](https://github.com/user-attachments/assets/b4602751-2c75-424c-b3cf-d36aea69623c)

- Local Server->SQL Server extraction:
  
![image](https://github.com/user-attachments/assets/7715dbd1-7b85-4b78-909b-251d9415d8e9)


The 
