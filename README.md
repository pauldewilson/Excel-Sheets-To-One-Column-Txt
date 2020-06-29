# Excel-Sheets-To-One-Column-Txt
Turns the contents of spreadsheets into a single columns separated by category (column) with a header category (index)

The purpose of Spreadsheet Values To Single Row In Text File is to transform an xlsx file's data to one big row in a text file.
This was done to enable easier qualitative data analysis coding (the process of highlighting qualitative data for later analysis - 
not programming coding).

It provides some very basic formatting based upon the category (column).

An example would be a spreadsheet table that looks like this (from CSV format) where each row is formatted in the same way:

`country,region,continent,population,gdp,gdppc<br>`
`Monaco,Western Europe,Europe,39242,6468000877,164823`

and turn it into:

####################
Monaco
####################

----------region: 
Western Europe

----------continent: 
Europe

----------population: 
39242

----------gdp: 
6468000877

----------gdppc: 
164823

There are two options when running the script:

You can go 'fast' which will rinse every single xlsx file within a folder.
This will open the file and iterate over each sheet whereby whatever is in A:A will function as the index (in the above example, Monaco)
and will then output all the other columns as above region, continent, population etc.
It then outputs a .txt file named xlsxname_sheetname.txt where xlsxname is the name of the workbook and sheetname is the name of the worksheet.
Each worksheet gets its own textfile.

The second option is to go 'slow' whereby it will iterate over every xlsx in a folder but ask the user to select:
	1. The sheets to keep
	2. The columns within the kept sheets to use and,
	3. The index column (which can be in any column, not just A:A) which ultimately will be the primary cateogry (above example, country)
