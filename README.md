# VBA of Wall Street

## Overview of Project

### Purpose
Performing analysis on green energy stock data to determine whether refactoring the Module 2 solution code successfully made the VBA script run faster.

### Background
Steve want to expand the dataset to include the entire stock market over the last few years to perform his analysis on how actively a stock is traded and each stock performed over the last few years. We will refactor the Module 2 solution code by switching the nesting order of our for loops. Then, we'll determine whether refactoring our code successfully made the VBA script run faster. This will ensure our code will not take a long time to execute if Steve expand the dataset to include the entire stock market over the last few years.

## Results

### Analysis of Yearly Volume

We generated a table with Total Daily Volume for each green stock to help us understand how active the stock is traded in a particular year, see Total Daily Volume column in the result images below. The total daily volumne ranges from 35 to more than 684 million in year 2017 and ranges from 83 to more than 607 million in year 2018. We can see that DQ has the lowest volume in year 2017, but it is still at more than 35 million. Since all of these stocks are traded often, the prices will accurately reflect the value of the stocks.

![VBA_Challenge_ResultTable_2017](https://github.com/grwon/VBA_Challenge/blob/master/resources/VBA_Challenge_ResultTable_2017.png).
![VBA_Challenge_ResultTable_2018](https://github.com/grwon/VBA_Challenge/blob/master/resources/VBA_Challenge_ResultTable_2018.png).

### Analysis of Stock Performance Based on Yearly Returns

We populated the yearly returns for each stock in tabluar data format. If we look at the images of results above, we can visualize the return results using the return columns. Most of 2017 is green and most of 2018 returns are red. We noted that all of the stocks performed better in 2017 than 2018, except for stock tickers TERP and RUN. We also noted that ticker ENPH outperformed all other stocks on average for the two years in 2017 and 2018, with a return of 129.5% in 2017 and 81.9% in 2018. DQ had the highest return in 2017 compared to the other green stocks, however, they suffered a return of -62.6% in 2018. Having said that, DQ still has a very high favourable return over the two years in 2017 and 2018 and on average over the two years is one of the top 3 performers amongst these green stocks. 

### Analysis of Execution Times of Original vs Refactored Scripts

We can look at the execution times of the original script and the refactored script to determine whether refactoring our code successfully made the VBA script run faster. The execution time for year 2017 data is only 0.0703125 seconds with the refactored code compared to 0.4648438 seconds using the original code. The execution time for year 2018 data is only 0.0859375 seconds with the refactored code compared to 0.4414063 seconds using the original code. These results confirmed that refactored code successfully made the VBA script run faster.

![VBA_Challenge_2017](https://github.com/grwon/VBA_Challenge/blob/master/resources/VBA_Challenge_2017.png).
![VBA_Challenge_2018](https://github.com/grwon/VBA_Challenge/blob/master/resources/VBA_Challenge_2018.png).

## Summary

- What are the advantages or disadvantages of refactoring code?
  
  Refactoring process has the following advantages:
  1. improve performance by reducing the steps, using less memory, making the script run faster
  2. find bugs in original scripts
  3. improve the design and logic of the code to make it easier to understand
  4. maintain and organize to make the code cleaner

  Refactoring process has the following disadvantages:
  1. may introduce new bugs when you alter your script
  2. may be costly if the original code is poorly written and it would take too much effort to fix bugs
  3. may be risky if there is a tight deadline

- How do these pros and cons apply to refactoring the original VBA script?
By performing refactoring process to the original VBA script, we have improved performance by looping over all the rows only once instead of looping over all the rows for each ticker. The original script loops through all the rows 12 times to get the Total Daily Volume and return for a specific ticker each time, whereas the refactored code loops through all the rows only once and store the Total Daily Volume and return data to the arrays tickerVolumes(), tickerStartingPrices() and tickerEndingPrices(). This reduces the steps and uses less memory, thus making the refactored script run faster than original. The refactored code is cleaner and easier to understand. The cons did not apply to our original VBA script as it is not a very complex script or big project. However, it is possible that it would not be feasible to perform refactoring process for our VBA script if the deadline is very tight and it would be risky to start refactoring process if we do not have enough time to re-run checks.
