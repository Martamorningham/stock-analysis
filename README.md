# Stock-Analysis
## Overview of Project
### Purpose 
 To make this code cleaner, in order to run through larger data sets in the future. While the initial code achieved the desired analysis goal, it ran much slower than the refactored code. Steve realized that he will be looking at much larger data sets using this code, so he wanted to refactor it and then compare the elapsed run times for both the refactored and orginal code. From this we can see weather our refactored code would be more effiencent on a larger scale then our original code.   
 
### Background 
Steve’s Parent’s bought ticker "DQ" becasue the ticker reminded them of the Dairy Queen they met at. With our help Steve  analyzed a series of 12 different stocks to determine wether "DQ" has been the best investment for the family, or if they should be looking at other stocks. We first looked at the year 2018, but then made some changes so we could look ath both 2017 and 2018. We did this by adding the Inputbox below
#### InputBox:
	yearValue = InputBox("What year would you like to run the analysis on?")
	
This allowed us to input the year that we want to anlyze. With this addtion we can now run analysis for either year. We then wrote a VBA script to format the output table that we created. 

Steve wanted to see how effiecnt the code was so we set a timer to run from the begining of the script to the end. We could then compare the initial code written in the module with the refactored code we worked on with steve. 
 	
## Results

### 2017 performance

2017 was a good year for this group of stocks. All but 1 of our 12 stocks ended the year in the green. “DQ” the Stocks Steve’s Parents originally invested in, did particularly well with a 199.4% return. This was the best return of all the stocks studied in 2017.

![2017_fullsheet](https://user-images.githubusercontent.com/101226991/162647138-7dcc0a5a-fe50-483e-9d9e-721b8e3506d2.png)	
![VBA_Challange_2017](https://user-images.githubusercontent.com/101226991/162647039-f57cbb37-3d01-4191-9144-4c1a12c6f5ba.png)
			
### 2018 performance 
2018 was a bad year for stocks, only 2 of the 12 stocks ended the year in the green. While the whole dataset did poorly in 2018, DQ did the wort of the whole data set. It lost over half of its value, as it fell 62.6% in 2018.

![2018_fullsheet](https://user-images.githubusercontent.com/101226991/162647157-f8f1da58-9d87-41b0-8ea8-e7978b56b911.png)
![VBA_Challange_2018](https://user-images.githubusercontent.com/101226991/162647178-5ecc0689-569f-412d-9a02-cc6f09c41499.png)

	
### 2017 vs 2018
	2017 was a much better year for stocks than 2018. This could be the result of the 2018 market crash, but our data does not provide any insight into the crash and how it effected our stock data for the year, outside of the market showing as red in 2018 compared to the green in 2017. Steves parents still did a good job with DQ, as the nearly 200% rise in 2017, helped to offset the huge drop that came in 2018. However, if they had invested in ENPH a stock with much more daily trading volume than DQ, they would have reaped an almost 130% return in 2017 and an 82% return in 2018. This would have been the best stock for them to incest in, as the only other stock in the green for 2018 "RUN" only rose 5.5% in 2017. 
![Final_Stock-Table_2017](https://user-images.githubusercontent.com/101226991/162648074-6ff6b680-878b-4a8f-8740-d8a21a846ce9.png)

![Final_Stock-Table_2018](https://user-images.githubusercontent.com/101226991/162648067-55b80a23-7680-4b43-9146-60d41c705dd8.png)

The elapsed runtimes for 2017 and 2018 were very similiar, the only the major difference between the run times was the code run. The refactored code was far faster than ther the orginal AllStocksAnalysis code.

## Summary
### Advantages and Disadvantages of Refactoring Code in General

Refactored code can make the code cleaner and easier to use in practice in the long run. While it may be easier and faster to execute, it can take longer time to think out and write the code. When you write code one way, it can be difficult to step back and rethink what could make this code better. The time taken to make code better and easier to read, does not always mean that it will save you time on that particular project. However, it will allow you to apply the code to a larger dataset with more ease. 
	
### Advantages and Disadvantages of the original and Refactored VBA Script
	
According to the Timer, the refactored code ran much faster than the original code from the module. That said, the amount of time taken and the run time displayed in the “AllStocksAnalysis” code do not match up. The displayed runtime for 2017 was 61390.29 seconds and the displayed runtime for 2018 was 61352.54. 
![Module_2017_time](https://user-images.githubusercontent.com/101226991/162646932-66274538-f248-4673-b97b-d53b073238e2.png)
![Module_2018_time](https://user-images.githubusercontent.com/101226991/162646906-45d8783e-e75e-40ec-a177-f0f0482f44a8.png)
However, I believe the original code may be running this in milliseconds, since it did not take over 1,000 minutes to display the results. This may be due to the fact that the original stock analysis uses a nested loop and the refactored code does not. The refactored code not only runs faster, it clears, and formats the output table as part of one code. In the original “AllStocksAnalysis” code, this was split between 3 different buttons, which looks a little clunky and can be confusing if you do not format and/or clear the sheet before running a second analysis. For Example, if you run analysis for 2017 and then run it for 2018, without formatting first. You could see the green values from 2017 remain on the negative outcomes from 2018. This could potentially lead someone to draw the wrong conclusions for the outcomes, if they are quickly referencing the data.  
The time it took to refactor the code, does not seem like it saved time overall on the analysis. The refactored code is more difficult to fully understand from the original code in the module. Splitting up the Analysis, Formatting, and Clear sheet codes make it easy to edit each one separately.
