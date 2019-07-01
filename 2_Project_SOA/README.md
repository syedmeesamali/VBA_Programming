# Statemnts of Accounts for various Accounts
*This VBA project involves SAP Database to Statement of Accounts*

An input data is taken from a SAP database onto the sheet named as "SAP". This sheet is then used to create "Statement of Accounts" for various unique accounts. Full code is given inside the file "wellness_main.xlsm" as well as a separate ".bas" file called "Final_code.bas".

Some cleaning of input data is required before creating the output e.g. if there are some empty cells in the SAP Data then the same should be filled by appropriate means. One way to fill a column having empty values is as below.

1. Make filter of the column
2. Select the filtered values and press "Ctrl + G". In the dialog box select "Visible Cells" only.
3. After selection type suitable value to fill all the cells and then press "Ctrl + Enter" to fill all the cells simultaneously. 

This data cleaning will ensure consistency of output and won't mess with the output layout of the "Statement of Account" sheets for each customer. 
