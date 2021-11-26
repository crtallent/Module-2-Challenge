# VBA Stocks Analysis Challenge
## Overview

The purpose of this analysis is to provide information to Steve, regarding the value of his parents' investments in the stock market.  Steve is concerned that his parents are not getting a return on their investment in the stocks they purchased for DQ, and would like to show his parents the value of their returns.  When performing the initial analysis for Steve, it was noted that DQ performed at a rate of nearly a -60% return https://github.com/crtallent/stock-analysis/blob/main/green_stocks.xlsm!  This was, of course, very concerning to Steve, who wanted to find out how other stocks fared in relation to DQ.  

## Results

Ten addtional tickers were analyzed to determine whether these stocks performed with a better return rate than DQ.  During the analysis, it was found that both ENPH and RUN had high return rates (nearly 82% and 84% respectively).  In order to perform these analyses, nested loops were created, as well as a flexible Macro ("Run" Button) to ensure that different sets of stocks could be ran by Steve by just pressing a button in the future.  The code was then run for 2017 and 2018, and a timer was added to calculate the time it took for Excel to perform the update.  The timer indicated that the scripts ran in a little over half a second for both [2017](https://github.com/crtallent/stock-analysis/blob/main/green_stocks%202017.png) and [2018](https://github.com/crtallent/stock-analysis/blob/main/green_stocks%202018.png).

While Steve appreciated this new tool, a concern was brought up.  What would happen if Steve wanted to run this analysis in the future for years to come?  What if Steve wanted to run the analysis for more stocks than initially coded?  This would surely take a lot longer to process a much larger amount of data.  It was then determined to refactor the existing code in the macro to allow for expansion in the future, and to cut down on the time it would take to run the analysis.  The code was then refactored to run more efficiently and to loop through all of the data at one time to speed up the process.

## Summary

