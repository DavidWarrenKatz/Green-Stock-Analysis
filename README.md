#Stock Analysis -Challenge 2 Overview

This project contains VBA macros that analyze a list of stock information ordered by ticker and day. The output is the total daily volume and yearly return of the 12 stocks in the data set. The outputted annual yields are then colored green or red based on profitability. 

#Results
The All Stocks Analysis Sheet has two buttons, one that initiates the standard analysis and one that initiates the refactored analysis. Both of these macros perform the same calculations. The difference is that the refactored analysis only iterates through the data once, while the standard analysis iterates through it 12 times - once for each stock ticker. As expected, the refactored analysis runs about 12 times faster. The only disadvantage of the refactored analysis is that it requires slightly more memory space since it uses three additional arrays to keep track of data as it iterates through the loop.

#Summary
The standard analysis is simpler conceptually and requries less memory. The refactored analysis is 12 times fact and does not require nested for loops since it only iterates through the data once. 


