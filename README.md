# Stock-Analysis
## Overview of Project
### Purpose 

 make this code cleaner, in order to run through larger data sets in the future. While the initial code achieved the desired analysis goal, it ran much slower than the refactored code.  
### Background 
Steve’s Parent’s bought ticker "DQ" becasue the ticker reminded them of the Dairy Queen they met at. With our help Steve  analyzed a series of 12 different stocks to determine wether "DQ" has been the best investment for the family, or if they should be looking at other stocks. We first looked at the year 2018, but then made some changes so we could look ath both 2017 and 2018. We did this by adding the below Inputbox: 
yearValue = InputBox("What year would you like to run the analysis on?")	
This allowed us to input the year that we want to anlyze. 


added a Once we completed the initial analysis, Steve realized that we would like 
 	
## Results

### 2017 performance

2017 was a good year for this group of stocks. All but 1 of our 12 stocks ended the year in the green. “DQ” the Stocks Steve’s Parents originally invested in, did particularly well with a 199.4% return.  
![2017_fullsheet](https://user-images.githubusercontent.com/101226991/162647138-7dcc0a5a-fe50-483e-9d9e-721b8e3506d2.png)

	
![VBA_Challange_2017](https://user-images.githubusercontent.com/101226991/162647039-f57cbb37-3d01-4191-9144-4c1a12c6f5ba.png)
	
	
	
### 2018 performance 
Bad year for stocks, only 2 of the 12 stocks ended the year in the green. It was particularly bad 

![2018_fullsheet](https://user-images.githubusercontent.com/101226991/162647157-f8f1da58-9d87-41b0-8ea8-e7978b56b911.png)
![VBA_Challange_2018](https://user-images.githubusercontent.com/101226991/162647178-5ecc0689-569f-412d-9a02-cc6f09c41499.png)


	
### 2017 vs 2018
	Our Analysis 
### Refactored code 

## Summary
### Advantages and Disadvantages of Refactoring Code in General

Refactored code can make the code cleaner and easier to use in practice in the long run. While it may be easier and faster to execute, it can take longer time to think out and write the code. When you write code one way, it can be difficult to step back and rethink what could make this code better. The time taken to make code better and easier to read, does not always mean that it will save you time on that particular project. However, it will allow you to apply the code to a larger dataset with more ease. 
	
### Advantages and Disadvantages of the original and Refactored VBA Script
	
According to the Timer, the refactored code ran much faster than the original code from the module. That said, the amount of time taken and the run time displayed in the “AllStocksAnalysis” code do not match up. The displayed runtime for 2017 was 61390.29 seconds and the displayed runtime for 2018 was 61352.54. 
![Module_2017_time](https://user-images.githubusercontent.com/101226991/162646932-66274538-f248-4673-b97b-d53b073238e2.png)
![Module_2018_time](https://user-images.githubusercontent.com/101226991/162646906-45d8783e-e75e-40ec-a177-f0f0482f44a8.png)
However, I believe the original code may be running this in milliseconds, since it did not take over 1,000 minutes to display the results. This may be due to the fact that the original stock analysis uses a nested loop and the refactored code does not. The refactored code not only runs faster, it clears, and formats the output table as part of one code. In the original “AllStocksAnalysis” code, this was split between 3 different buttons, which looks a little clunky and can be confusing if you do not format and/or clear the sheet before running a second analysis. For Example, if you run analysis for 2017 and then run it for 2018, without formatting first. You could see the green values from 2017 remain on the negative outcomes from 2018. This could potentially lead someone to draw the wrong conclusions for the outcomes, if they are quickly referencing the data.  
The time it took to refactor the code, does not seem like it saved time overall on the analysis. The refactored code is more difficult to fully understand from the original code in the module. Splitting up the Analysis, Formatting, and Clear sheet codes make it easy to edit each one separately.
