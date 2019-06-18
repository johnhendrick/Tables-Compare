# Tables-compare
This script validates 2 excel/csv files

### Rationale:
The day-to-day tasks of Business Analyst, Functional Consultants, and Data Analyst demand validating datas being communicated between cross-functional teams or data extracts from enterprise systems. Validating 2 or more excel files is a time-consuming manual task in excel and VBA macro is not always the most reliable and scalable method for large excel files. This Python script compare 2 excel files and output the delta between the two files. 

### Application:
- Post ETL load & extract validation
- Post ETL extract & raw file validation
- Checking delta between Excel Versions
- General purpose Excel table comparison

### The code:
The code converts 2 .xls files into pandas dataframes and perform validations against rows, columns and returns delta analysis & mismatches in each rows and columns in another excel output.xlsx

(detail descriptions WIP)



### Dependencies:
Python > 2.7
pandas
numpy 
Tkinter
tkFileDialog
