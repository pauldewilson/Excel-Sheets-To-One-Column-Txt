# Excel-Sheets-To-One-Column-Txt
Admittedly, this is my first github forray. Sorry for any errors :)<br>
Turns the contents of spreadsheets within a folder into a text file where each column is separated into 'one big column' where column names become categories and one column serves as a master category (termed: index).

The purpose for this script was to enable easier qualitative data analysis coding (the process of highlighting qualitative data for later analysis - 
not programming coding).

It provides some very basic formatting based upon the category (column).

Below is an example table from CSV format but presently the script only supports .xlsx.<br>
Beneath the sample csv data is an example of how each row will be represented (for example, if the next csv row was Germany with all its respective data, it would be formatted in the same way).

`example.xlsx` with sheetname `monacostat`

`country,region,continent,population,gdp,gdppc`<br>
`Monaco,Western Europe,Europe,39242,6468000877,164823`

will turn into into `example_monacostat.txt`<br>
The above would look like this:

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

(example end)

There are two options when running the script:

You can go `fast` which will rinse every single xlsx file within a folder.<br>
This will open the file and iterate over each sheet whereby whatever is in column A will function as the main header 'index' (in the above example the index was country, Monaco)
and will then output all the other columns within the sheet (in above example, region, continent, etc) as subheaders.<br>
It outputs a .txt file named `xlsxname_sheetname.txt` where `xlsxname` is the name of the workbook and `sheetname` is the name of the worksheet.<br>
Each worksheet gets its own textfile.

The second option is to go `slow` whereby it will iterate over every xlsx in a folder but ask the user to select:

<ol>
	<li>The sheets to keep
	<li>The columns within the kept sheets to use and,
	<li>The index column (which can be in any column, not just A:A) which ultimately will be the primary category (above example, country)
</ol>

For some context, one 465kb .xlsx workbook with a single 10,000x row and 6x column sheet takes approximately 10 seconds to process on a mid-spec laptop and will generate a 2.1mb textfile with 189,982 lines.<br>
This means the output filesize will be approx 4.5x greater and the number of rows will increase by approx 18x.
