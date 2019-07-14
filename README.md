# Invetory-Age-Automation
A program in python that automates the Dirty Dozen report. The Dirty Dozen is a list of the 12 oldest cars at the dealerships.

The script is named ‘dirtydozen.py’. The input file is a excel file with inventories from six dealerships named 'DDC.xlxs'. The output is the same ‘DDC.xlsx’ file with the first three tabs filled: ‘Dirty Dozen’, ‘Excluded’ and ‘All Stores’. ‘dirtydozen.py’ and ‘DDC.xlsx’ should be in the same folder before running the program. The program takes approximately 30 seconds to run. You'll notice that the output tabs are not formatted. This is because I wrote the output back into the original spreadsheet. The output tables can be formatted if the program produces a new file as the output. This can be implemented in a later version.


The program is written in python3 and requires you to install the following packages:

pandas

bs4

requests

openpyxl


The program sends a request to each dealerships website to see if the car is online. I put in a 0.5 second delay in between requests. The request is only run for 15 or so vehicles.


The program first reads in the 6 separate dealership specific inventories and then combines them into one All Stores dataframe. A dataframe is a 2-dimensional labeled data structure.  Next, All Stores is sorted by age and the age is updated according to the current date. The dataframe is then run into a function that checks if the vehicle is listed online and has less than 1000 miles. Next, the Dirty Dozen and Excluded dataframes are compiled and the excluded vehicles are removed from the All Stores dataframe. Finally the three completed dataframes are written to the excel sheet.
