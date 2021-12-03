# VBA Stocks Analysis Challenge

## Overview of Project

This project consisted of analyzing different stock market indexes to determine how they performed.  The first analysis included information for only one index.  The second and more detailed analysis included a total of 12 stock market indexes and their performance (return rates) for two years.

### Purpose

The purpose of this analysis is to provide information to Steve, regarding the value of his parents' investments in the stock market.  Steve is concerned that his parents are not getting a return on their investment in the stocks they purchased for DQ, and would like to show his parents the value of their returns.  When performing the initial analysis for Steve, it was noted that DQ performed at a rate of nearly a -60% return https://github.com/crtallent/stock-analysis/blob/main/green_stocks.xlsm!  This was, of course, very concerning to Steve, who wanted to find out how other stocks fared in relation to DQ.  

## Results

Ten addtional tickers were analyzed to determine whether these stocks performed with a better return rate than DQ.  During the analysis, it was found that both ENPH and RUN had high return rates (nearly 82% and 84% respectively).  In order to perform these analyses, nested loops were created, as well as a flexible Macro ("Run" Button) to ensure that different sets of stocks could be ran by Steve by just pressing a button in the future.  The code was then run for 2017 and 2018, and a timer was added to calculate the time it took for Excel to perform the update.  The timer indicated that the scripts ran in a little over half a second for both [2017](https://github.com/crtallent/Module-2-Challenge/blob/main/Resources/VBA_Challenge_2017.png) and [2018](https://github.com/crtallent/Module-2-Challenge/blob/main/Resources/green_stocks%202018.png).

While Steve appreciated this new tool, a concern was brought up.  What would happen if Steve wanted to run this analysis in the future for years to come?  What if Steve wanted to run the analysis for more stocks than initially coded?  This would surely take a lot longer to process a much larger amount of data.  It was then determined to refactor the existing code in the macro to allow for expansion in the future, and to cut down on the time it would take to run the analysis.  The code was then refactored to run more efficiently and to loop through all of the data at one time to speed up the process.  While the times taken to run the code did not decrease as much for both years, it did decrease to a little over .004 seconds for [2017](https://github.com/crtallent/Module-2-Challenge/tree/main/Resources), indicating that a well-written code would ensure that Steve is able to run his analysis for years to come.

## Summary

In summary, using a refactored code can save time as many of the components are already present in the first code you created.  For example, you can cut down on the steps taken to get the code to work.  You can also add additional comments that will make the code easier to read for those that may read it in the future.  You can also find new and more efficient ways of accomplishing a task when you run back through a code that you have already created.  This can make it run much faster and more efficiently in the future.  A con of using refactored code would be that there may be a better way to accomplish the task at hand, so you may find that you need to just create a new routine entirely.  It also may be difficult to refactor someone else's code if you are unsure of the logic they used, especially if clear notes are not present in the code.

For this task, I found using the refactored code easier at times, as I did not need to type in everything that was still relevant for each task.  On the other hand, I felt that refactoring an existing code would be easier for more experienced coders, as they would be able to bring in knowledge they possessed already of more efficient ways to change the code going forward.  Another challenge I found in refactoring the code was determining where to input the changes, as I had a lot of break errors come up.
