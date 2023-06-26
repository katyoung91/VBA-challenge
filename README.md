# VBA-challenge
Module 2 Homework Assignment for Katherine Young.
This repo includes the export of the VBA script, screenshots of each page in the Excel file where I ran the code, and this Read me.

A couple notes based on the assignment:
- My module script contains two Subs, Stonks and All_Stonks. Use All_Stonks to run the script on all sheets in the wrokbook. Stonks is the main code that runs everything in the assignment on a single sheet. All_Stonks calls the Stonks Sub to be run on each sheet in the workbook.
- For applying conditional formatting, I did this through VBA but not through a conditional formatting tool or function. The instructions aren't clear if you wanted us to do this using say ".FormatConditions" or methods that were gone over in class. I opted for what we went over in class and got the desired result for the assignment so I hope this approach works.
- As a follow up, the grading rubric describes "Conditional formatting is applied correctly and appropriately to the percent change column (10 points)" but there were no conditional formatting instructrions in the assignment for this column. I did apply general formatting like making it a percent out to 2 decimal places in my code and hope that's what you are looking for.
-  I actually found 3 different ways to get the yearly change and percent change values, 2 of which used nested for loops, but these took hours to run with how big the data and summary tables are. The single loop I have for pulling all initial values was the fastest way to make this all work. 
