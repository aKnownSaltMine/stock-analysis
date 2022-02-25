# Green Stocks Analysis
## Overview of Project
With the environmental conciousness becoming evermore real for many consumers, they are looking to invest their money into companies that are building the infrastruture and taking part in the transition to renewable energy from fossil fuels. It is for this person that our client commissioned us to create an Excel workbook in order to analyze multiple stocks performance over the years 2017 and 2018 specifically in the realm of solar. Twelve companies were looked at in the original run of the project, and they are the following:
 - Atlantica Sustainable Infrastructure (AY)
 - Canadian Solar (CSIQ)
 - Daqo New Energy Corp (DQ)
 - Enphase Energy (ENPH)
 - First Solar (FSLR)
 - Hannon Armstrong Sustainable Infrastructure (HASI)
 - JinkoSolar Holding Co (JKS)
 - Sunrun (RUN)
 - Solar Edge (SEDG)
 - SunPower (SPWR)
 - TerraForm Power (TERP)
 - Vivint Solar (VSLR)
We were able to create a simple but powerful workbook that can run analysis on all the data given and output it into a format that is quickly understandable utilizing VBA code and macro capability natively available in Excel. 
### Purpose
The overall purpose of this project is to compare the end of year performance of the stocks of the companies listed in a given year and output it in a format that is quickly readable. The output was focused on the total volume that the stock was traded in and how it was valued in comparison to a year before. By putting it in this readable format, our client will be able to make informed financial decisions with much more ease. Also because the VBA code has been futher optimized, then it should be able to be adapted to run through thousands of stocks without much issue. 
## Results
A link to the final workbook including the macros can be found [here](https://github.com/aKnownSaltMine/stock-analysis/blob/main/VBA_Challenge.xlsm)
### Comparison of 2017 and 2018 Stock Performance
 ![Stocks Analysis 2017](https://github.com/aKnownSaltMine/stock-analysis/blob/main/Resources/Stocks_Analysis_2017.PNG) ![Stocks Analysis 2018](https://github.com/aKnownSaltMine/stock-analysis/blob/main/Resources/Stocks_Analysis_2018.PNG)
 
 Many of the companies examined in the breakdown did well in 2017 with all except TerraForm Power (-7.2%) showing positive gains. Some like Sunrun showed modest gains of 5.5%, while companies like Daqo showing growth of almost 200% though the daily volume was considerably lower than the rest of the stocks. 
 
 However, 2018 tells another story for Solar companies. All aside from two (Enphase Energy and Sunrun) did post negative numbers in growth. Even the primary stock of interest of the Client, Daqo, posted a -62.6% on return. It is difficult to recommend taking up such a stock after such a drop in performance from the previous year. After examining year over year performance, however, the safest stock to invest in would be Enphase Energy. In 2017, it had a return of 129.5% with a daily trading volume of almost a quarter of a million which is on par with the rest of the stocks examined. But also it is one of the only ones to show positive returns in 2018 while the rest of the stocks contracted. So with all of this in consideration, the best stock for our client to focus on would be ENPH.
### Comparison of VBA Code Used to Run Analysis
The major difference between the [original code](https://github.com/aKnownSaltMine/stock-analysis/blob/main/Resources/AllStocksAnalysis()) (Lines 50-84) and the [refactored code](https://github.com/aKnownSaltMine/stock-analysis/blob/main/Resources/AllStocksAnalysisRefactored()) (Lines 50-94) was that the original only used a single array in order to keep track of the actual ticker when going through the for-loop. Because of this, at the start of each loop the macro would have to activate the worksheet of the year that we input, and then would also have to output the value before moving into the next loop. Because the refactored code was able to use multiple arrays in order to contain all of the data, it only had to activate the the workbook once, loop through and gather all the needed data for each ticker. Then it only needed to output once. This alone was able to speed up the macro considerably. 
### Code Performance After Refactoring
On the [original code](https://github.com/aKnownSaltMine/stock-analysis/blob/main/Resources/AllStocksAnalysis()) the two different years would run just under a second for analyzing all of the stocks in the dataset.

![2017](https://github.com/aKnownSaltMine/stock-analysis/blob/main/Resources/VBA_Challenge_2017%20(unrefactored).PNG) ![2018](https://github.com/aKnownSaltMine/stock-analysis/blob/main/Resources/VBA_Challenge_2018%20(unrefactored).PNG)

Though after refactoring, the macro was able to run for both years in just over 10th of a second for a total reduction in runtime of 83.26%.

![2017](https://github.com/aKnownSaltMine/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG) ![2018](https://github.com/aKnownSaltMine/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

## Summary
The main advantage of refactoring code that it improves in efficency by improving on the logic and the amount of commands that the computer would have to execute. The major disadvantage of refactoring a block of code is that the code already works, if not optimized. So you would be spending your time fixing what is not broken. Also there is no guaruntee that the time spent refactoring the code will actually yield any noticable results, making the time spent, for nothing. 

Because of the use of multiple arrays as explained above, the refactored code had a major advantage in speed over the original code, while delivering the same output in the end. Although one of the major disadvantages of both macros is that they are not that adaptable. If we wanted to run an analysis on different stocks, we would have to modify the actual VBA code of the macro rather than presenting an End User friendly version of inputing the stocks that they want to run analysis on for the macro to take and then output into a similar table. 
