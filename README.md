# 1. Package Description

![image](https://github.com/user-attachments/assets/2ac7e7ce-fbb1-49c4-add3-6faee410296e)


This code presents a template to dynamically extract tables from Excels stored in Sharepoint onto SQL Server tables, using a SQL Server Integration Services (SSIS) package. It takes a given table in an Excel sheet with headers and transforms it into an SQL Server table. This module is not optimal for big data, but for reference tables hand-written in Excel.

We here assume the user has access to an OLEDB connection to an SQL Server database. Note that this connection is absent in this template, so you need to create it.

The extraction automation includes a retry mechanism, which premise may be useful in other SSIS applications, and it is constructed in two main procedures written as C# scripts:

- Sharepoint/Excel->Local Server extraction (output path defined in local package variables):
  
![image](https://github.com/user-attachments/assets/b4602751-2c75-424c-b3cf-d36aea69623c)

- Local Server->SQL Server extraction:
  
![image](https://github.com/user-attachments/assets/7715dbd1-7b85-4b78-909b-251d9415d8e9)


The extraction of the Excel tables is within a foreach with the following parameters:

- Path to the Source Excel within Sharepoint 
- Sheet name where the source table is stored in the source Sharepoint Excel
- Destination Schema name the user wants
- Destination table name the user wants

The parameters above are given from a query to a parameter table in SQL Server defined in the following manner: 

![image](https://github.com/user-attachments/assets/739f2c34-1d62-41cf-a7a8-ba7b02c5e1f3)

![image](https://github.com/user-attachments/assets/b034b196-d495-480a-b554-dcd52f4e2239)

where each field is a VARCHAR. Of course, the user may change the package so that the way these inputs are processed differ. This is just a useful example.

If this package value is set to True:

![image](https://github.com/user-attachments/assets/c549b6ef-49b3-4d34-b7c4-e2c7795b72b7)

# 2. Important Notes

Destination tables in SQL Server arrive with the datatype NVARCHAR (3000) always. This is hardcoded in the script-task "Extraction Server-SqlServer" and comes from the premise that it will be difficult to maintain data coherence within a hand-populated Excel for months or years. 

The Excel path is found by looking at the Excel properties in the Sharepoint folder. This is the one you shall use to parametrize your input:

![image](https://github.com/user-attachments/assets/5ddef99b-1569-4fef-939b-97759236e53e)

By default, the OLEDB connection used in the "Extraction Server-SqlServer" tries to guess the data-types of each Excel header with the first 8 entries. This results in null or empty values in situations where there are mixed values (e.g. first 8 values are Int and 9th value is a VARCHAR). In order to solve this issue and load everything as a string we need to:

1- Set Windows Registry TypeGuessRows as 10000 (or more if table has more entries):

![image](https://github.com/user-attachments/assets/b816ca17-5783-48c3-84aa-d496bbd973c0)


![image](https://github.com/user-attachments/assets/efa9734d-6de7-493a-9611-f0f7d86ac3d2)

2- Set parameters "HDR" and "IMEX" as "YES" and "1", respectively, in the Connection String used for the Excel. This is already hardcoded in the script task "Extraction Server-SqlServer":

![image](https://github.com/user-attachments/assets/c004fb0f-b8b0-471e-8f67-c99d2230dd75)

Empirically we noted that in order to reduce retries and transient errors in between extractions, it was important to kill any Excel applications and delete cache files:

![image](https://github.com/user-attachments/assets/07c0f999-f9e1-4ace-bad1-b26ee98a2551)


The package will retry up until 1000 times (you can change this threshold under MaximumErrorCount properties), otherwise it will retry up until the value set in the RetryCount value:

![image](https://github.com/user-attachments/assets/e35587c1-6e4d-4096-98e2-bf27ffb7007a)

