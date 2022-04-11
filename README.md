# Stock-Analysis
## Overview of Project:
### Purpose: 
 make this code cleaner, in order to run through larger data sets in the future. While the initial code achieved the desired analysis goal, it ran much slower than the refactored code.  
### Background: 
Steve’s Parent’s bought 
1st 	
We helped Steve analyze a range of stocks so that he can inform his parents whether they have made a good decision in the stock market. They can then use this information to help guide their decisions going forward. 
## Results:
### 2017 performance: 
	2017 was a good year for this group of stocks. All but 1 of our 12 stocks ended the year in the green. “DQ” the Stocks Steve’s Parents originally invested in, did particularly well with a 199.4% return.  
	(picture of 2017 output table and picture of 2017 Runtime)
### 2018 performance: 
	Bad year for stocks, only 2 of the 12 stocks ended the year in the green. It was particularly bad 
	(picture of 2017 output table and picture of 2017 Runtime)
### 2017 vs 2018:
	Our Analysis 
### Refactored code 

## Summary:
### Advantages and Disadvantages of Refactoring Code in General:
	Refactored code can make the code cleaner and easier to use in practice in the long run. While it may be easier and faster to execute, it can take longer time to think out and write the code. When you write code one way, it can be difficult to step back and rethink what could make this code better. The time taken to make code better and easier to read, does not always mean that it will save you time on that particular project. However, it will allow you to apply the code to a larger dataset with more ease. 
### Advantages and Disadvantages of the original and Refactored VBA Script:
	
	According to the Timer, the refactored code ran much faster than the original code from the module. That said, the amount of time taken and the run time displayed in the “AllStocksAnalysis” code do not match up. The displayed runtime for 2017 was 61390.29 seconds and the displayed runtime for 2018 was 61352.54. 
(display photos of original run times)
However, I believe the original code may be running this in milliseconds, since it did not take over 1,000 minutes to display the results. This may be due to the fact that the original stock analysis uses a nested loop and the refactored code does not. The refactored code not only runs faster, it clears, and formats the output table as part of one code. In the original “AllStocksAnalysis” code, this was split between 3 different buttons, which looks a little clunky and can be confusing if you do not format and/or clear the sheet before running a second analysis. For Example, if you run analysis for 2017 and then run it for 2018, without formatting first. You could see the green values from 2017 remain on the negative outcomes from 2018. This could potentially lead someone to draw the wrong conclusions for the outcomes, if they are quickly referencing the data.  
The time it took to refactor the code, does not seem like it saved time overall on the analysis. The refactored code is more difficult to fully understand from the original code in the module. Splitting up the Analysis, Formatting, and Clear sheet codes make it easy to edit each one separately.
