# Excel-Sheets-To-One-Column-Txt
Admittedly, this is my first github forray. Sorry for any errors.<br>
Turns the contents of spreadsheets into a single columns separated by category (column) with a header category (index)

The purpose of Excel-Sheets-To-One-Column-Txt is to transform an xlsx file's data to one big row in a text file.<br>
This was done to enable easier qualitative data analysis coding (the process of highlighting qualitative data for later analysis - 
not programming coding).

It provides some very basic formatting based upon the category (column).

An example would be a spreadsheet table that looks like this (from CSV format) where each row is formatted in the same way:

`country,region,continent,population,gdp,gdppc`<br>
`Monaco,Western Europe,Europe,39242,6468000877,164823`

and turn it into:

####################
<br>
Monaco
<br>
####################

----------region:
<br>
Western Europe

----------continent:
<br>
Europe

----------population: 
<br>
39242

----------gdp: 
<br>
6468000877

----------gdppc: 
<br>
164823

There are two options when running the script:

You can go `fast` which will rinse every single xlsx file within a folder.<br>
This will open the file and iterate over each sheet whereby whatever is in A:A will function as the index (in the above example, country, Monaco)
and will then output all the other columns within the sheet (in above example, region, continent, population etc).<br>
It outputs a .txt file named `xlsxname_sheetname.txt` where `xlsxname` is the name of the workbook and `sheetname` is the name of the worksheet.<br>
Each worksheet gets its own textfile.

The second option is to go `slow` whereby it will iterate over every xlsx in a folder but ask the user to select:

<ol>
	<li>The sheets to keep
	<li>The columns within the kept sheets to use and,
	<li>The index column (which can be in any column, not just A:A) which ultimately will be the primary category (above example, country)
</ol>