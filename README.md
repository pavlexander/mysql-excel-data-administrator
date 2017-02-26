# MySQL data administration in Excel
Excel which allows to administrate data for MySQL tables (or MariaDB, for example)

## Intro

I think that the "MySQL Workbench" is a great tool for managing tables, table diagrams (relations), indexes, views and other objects, but populating and managing data is the nightmere there, due to some bugs and not-so-friendy user interface.. so I created an Excel document just for that

## Use cases

1. if you want to load data (import) into some "MySQL Workbench" table
2. if you want to create INSERT statements for "PhpMyAdmin"

In both of the cases - you will be managing data in Excel document directly.
This is the tool that I have created for a personal use and I never use option 1. So some limitations are applied there. 

## How to use

### Generating the output file

In case 1:

1. open Excel
2. set "file name"
3. set "file extension" to "csv"
4. set "use SQL statement.." to "No"
5. click on "Generate file" button

In case 2:

1. open Excel
2. set "file name"
3. set "file extension" to "txt"
4. set "use SQL statement.." to "Yes"
5. click on "Generate file" button

A new file will be created in the same directory, where Excel file is. The file can then be imported in MySQL workbench (for different tables you shall create different files!!), or you can run all SQL inserts in PhpMyAdmin, for example.

### Creating worksheets (tables)

The problem is that you have to know exact columns that you have in tables, when populating the data in Excel. So I created a helper button, which will generate the new "Worksheet" in Excel, along with column names. So you have a metter managing capabilities.

1. go to "Mysql Workbench"
2. go to diagram view of your table
3. right click on table and use the option called "Copy Insert Template to Clipboard"
4. open Excel document
5. paste the insert template into field "Insert statement"
6. click on "Add Worksheet(Table)" button

>  P.S.: A button "Copy Insert Template to Clipboard" always generates data in format:
```
INSERT INTO `DATABASE`.`TABLE` (`COLUMN1`, `COLUMN2`, ...) VALUES (NULL, NULL, NULL, ...);
```
> So if you for some reason can not use the Workbench - you can either create Excel worksheets yourself (using the conventions below OR create an "insert statement template" youself.

## Conventions

- Excel shall always have a "main" worksheet and it shall be first in the list
- Worksheet (non-main) should have the matching name with the table name in order to generate data correctly
- In non-main worksheet - row 1 can have a value: NUMBER. This will force to not-use single quote for the generated value
- In non-main worksheet - row 2 is the default value for data wors. So if you are not specifying any values in data rows -> the default value will be taken from "row 2"
- In non-main worksheet - row 3 are just column names in MySQL table. This is only for your own convenience. It has no affect to anything.
- In non-main worksheet - row 4+. These are rows with your data.
- For all data rows where you are not putting "NUMBER" in row 1 - the script will automatically enclose the value in single quote characters
- Is you are using a double quote in your data cell value - it will automatically escape using the \ symbol so " will become \"

## Contributing

If you are a developer who wants to contribute to the project - after you change the VBA in Excel itself - don't forget to export VBA module into a sepaarte BAS file and commit it as well, so we keep the track of changes.




