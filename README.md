# VBA of Wall Street
Building on my knowledge of VBA (Visual Basic for Applications) by examining DAQO/DQ stocks for both 2017 and 2018, and then comparing them to other green energy stock data.
## Overview of Project
I created a workbook following the module that added the total daily volumes for all the DAQO/DQ green energy stocks.  Afterwards, I added the total daily volumes for all the other various green energy stocks.  The total daily volumes for each of the green energy stocks were then used to find the yearly returns for each company.  This was all accomplished through VBA coding.

The Deliverable 1 portion of this module saw me take the workbook that I had created in the original module and attempt to edit and refractor the code by switching the nesting order of the for loops.  I had to do this while creating a new variable (tickerIndex) and using it as a counter to keep track of which ticker I was working on in sections 7a) to 7c) of Deliverable 1.
### Purpose
Using my knowledge of VBA to examine if I could successfully refractor the code from the original module worksheet to make the green energy stocks analysis more efficient and faster, while also potentially being able to add more green energy stock market information from previous years (prior to 2017). 
## Analysis and Challenges
I was able to create both a functioning module workbook, and a functioning Deliverable using refractored code by mostly using the following module lessons: 

[2.2.3: Find Total Daily Volume for DQ in 2018] (https://bootcampspot.instructure.com/courses/193/pages/2-dot-2-3-find-total-daily-volume-for-dq-in-2018?module_item_id=52537)

[2.2.4: Get DQ's Yearly Return] (https://bootcampspot.instructure.com/courses/193/pages/
2-dot-2-4-get-dqs-yearly-return-for-2018?module_item_id=52538)

[2.3.2: Loop Over All Tickers] (https://bootcampspot.instructure.com/courses/193/pages/2-dot-3-2-loop-over-all-tickers?module_item_id=52541)

[2.3.3: Reuse Code] (https://bootcampspot.instructure.com/courses/193/pages/2-dot-3-3-reuse-code?module_item_id=52542)

The challenging part for me was declaring a new tickerIndex variable as a counter in Deliverable 1, while I was refractoring the code from the original module worksheet.  Trying to understand the purpose of the tickerIndex in the program and how to incorporate it in the subsequent coding was not easy.  

Creating the tickerIndex and using it to increase the total volume, as well as creating the if-then statements to check if current rows were the first and last rows with the selected tickerIndex differed from the original module worksheet.  The original worksheet had no tickerIndex and used if-then statements to find if current rows were the previous and following rows only.  Trying to incorporate these changes into the Deliverable 1 code was very time consuming.

### Analysis Original Worksheet (Tables and Timer Results)

./Resources/Original_Module_2017_Returns 
./Resources/Original_Module_2018_Returns
./Resources/Original_Module_Timer_for_2017
./Resources/Original_Module_Timer_for_2018

### Analysis Refractored Worksheet (Timer Results)

./Resources/VBA_Challenge_2017
./Resources/VBA_Challenge_2018

## Results
When examining the yearly returns from the original script for all 12 green energy stocks in both 2017 and 2018, we can see that the yearly returns for the stocks mostly increased in 2017, and mostly decreased in 2018.  As seen in the png files above, the only stock that had a yearly drop in 2017 was TERP with a 7.2% decline.  The other 11 green energy stocks all had yearly increases ranging from RUN's 5.5% increase to DQ's 199.4% increase.

That completely changed in 2018.  Only 2 green energy stocks showed increases in 2018.  Those were ENPH's 81.9% increase and RUN's 84.0% increase.  The other 10 green energy stocks declined in 2018.  Those declines ranged from VSLR's 3.5% decline to DQ's 62.6% decline.  Code formatting allowed me to display the stocks with yearly increases in green and the stocks with yearly declines in red.

For the original worksheet, the execution times from the timers were almost identical for both 2017 and 2018.  Each original worksheet (2017 and 2018) took between 1.21 - 1.32 seconds to complete as shown in the original worksheet timer results screenshots above.  The refractored coding worksheets (2017 and 2018) ran roughly 1 second faster as also shown in the VBA_Challenge png files.  The refractored timers for both 2017 and 2018 were roughly 0.25 seconds.

### Summary
As mentioned in the deliverable background of the assignment, refractoring improves both functionality and efficiency since it requires fewer lines of code to successfully run a program.  The refractored code also allows the program to run faster.  For a beginner in programing, trying to understand how the refractoring and the tickerIndex variable fit into the coding was difficult and probably the major disadvantage of this assignment.
