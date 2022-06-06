# #objectives:
#
# target.xlsx has 2 sheets:
#
# 1. "errors" - target errors that need to be sorted
# 2. "inv" - raw inventory data from salsify
#
# #we need to do the following::
#
# 1. remove columns B + C from errors, and then insert two more to the left
# 2. rename columns B, C, D, E to "Product Status (Computed),	Total Quantity,	salsify:parent_id, Action Taken"
# 3. put in VLOOKUPs into columns B, C, D which look like this:
# 	=VLOOKUP(A2,inv!A:D,2,0)
# 	=VLOOKUP(A2,inv!A:D,3,0)
# 	=VLOOKUP(A2,inv!A:D,4,0)
# 4. convert those formulae to plain text/integers
# 5. convert "errors" sheet to JSON (https://anthonydebarros.com/2020/09/07/json-from-excel-using-python-openpyxl/)
# 6. filter the data in "errors" in the following ways (https://stackoverflow.com/questions/27189892/how-to-filter-json-array-in-python):
# 	a. remove all SKUs with a "Product Status" of "Resourcing"
# 	b. remove all SKUs with a "Product Status" of "Do Not Reorder" and Total Quantity of < 130
# 	c. remove all SKUs with a "Item Type" of "Patio Umbrellas" or "Powered Riding Toys"
# 	d. remove all SKUs with a "Reason" from a list of reasons (need to create)
# 7. save the spreadsheet!
#
#
# #resources
#
# openpyxl - https://openpyxl.readthedocs.io/en/stable/usage.html
# convert to JSON - https://anthonydebarros.com/2020/09/07/json-from-excel-using-python-openpyxl/
# filter JSON - https://stackoverflow.com/questions/27189892/how-to-filter-json-array-in-python
# map two excel worksheets / VLOOKUP - https://www.youtube.com/watch?v=OB_ixIlQbWI
# excel cleaning - https://www.youtube.com/watch?v=o8uGCFZfBmM
# pandas filtering https://www.youtube.com/watch?v=Lw2rlcxScZY
#
# #finish this tutorial first
#
# https://realpython.com/openpyxl-excel-spreadsheets-python/#before-you-begin

# Note - is now working properly, only issue is this: an item such as SKY5047, which has multiple errors:
#   1. "UPC Not in WERCs. Supplier action needed." - not caught by the logic
#   2. "Image might be considered to be RACY" error - caught by the logic.
#   3. "Illustrations/Logos are not acceptable images." - not caught by the logic
# essentially what is happening is that the program is filtering out these SKUs entirely because they are caught by the "Racy" error.
# however, they should still be displayed because they are technically actionable - the UPC error is something that should show up!
# my suspicion is that the & / OR logic can be tweaked to make this work.
# https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.query.html