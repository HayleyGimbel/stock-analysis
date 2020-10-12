# Analyzing Stocks Using VBA

## Overview of Project
In this project I used VBA to perform an analysis on stock data for our client, Steve.  Steve works in finance and is advising his parents on the best stocks to invest in.  My analysis will provide Steve with inights on the Total Daily Volume and Yearly Return for each stock he is considering.  These insights will allow his parents to make more successful investments.  Additionally, the VBA code that I will provide for Steve will allow him to easily run similar analysis in the future.

## Results
My analysis of the stock data focuses on each stock's Starting Price, Ending Price, and Return in order to provide Steve with the information necessary to advise his parents on their investment. Using VBA (for) loops, I was able to calculate the Total Daily Volume from the Starting Price and Ending Price for each stock.  This calculation lead to the conclusion that in 2017 FPWR had the highest Total Daily Volume, while in 2018 ENPH had the highest Total Daily Volume.  I then calculated the Return for each stock, and determined that in 2017 DQ had the largest Return and in 2018 RUN had the largest Return.  Another important insight is that 2017 was statistially a better year for investments than 2018.  Steve can now use these insights to advise his parents based on their investment goals. 

![VBA_Challenge_2017](/Resources/VBA_Challenge_2017.png)

![VBA_Challenge_2018](/Resources/VBA_Challenge_2018.png)

## Summary

###
The primary advantage of refactoring code is that the framework for the code is provided, so the initial effort of conceptualizing and building code is not required.  The disadvantage here is that the code may not include comments or could be difficult to interpret.  In this case, more time and effort may be spent trying to decifer code than would be spent writing the code.  In this analysis, a combination of both methods were required to get the best result. 

### There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script
For my analysis, the main advantage of both the original and refactored VBA script was that the code could be used to easily run calculations on the data even when the data set is changed.  I could quickly run analysis on both worksheets in the provided data set, and I was able to set up buttons so that Steve could easily run the analysis in the future.  One of the disadvantages of the original VBA script was that it was not set up to run efficiently on larger data sets.  I was able to improve the code efficiency slightly with the refactored code, however the efficiency of the code seemed to be a main disadvantage of the VBA script. 
