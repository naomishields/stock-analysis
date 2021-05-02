# Stock Analysis with in Excel with VBA

## Overview of Project 

The purpose of this project was to help Steve analyze stock data to help advise his parents on what stocks to invest in. This goal was to be met through refactoring a previous solution. By refactoring this previous solution, the new analyze process should be executed faster and work well for larger amounts of data. 

## Results

Through this analysis, we were able to obtain the returns for the stocks from both 2017 and 2018. In 2017, most of the stocks had positive return rates, with only TERP having a negative return. However, in 2018 all but two of the stocks had negative return rates. Ultimately, the stocks performed better in 2017 as illustrated below. 
![2017 Stocks](https://github.com/naomishields/stock-analysis/blob/main/images/2017%20stocks.png)
![2018 Stocks](https://github.com/naomishields/stock-analysis/blob/main/images/2018%20stocks.png)

When looking at performance times, the refactored script was much faster than the original script. The original script took about 0.7734 seconds to run the analysis for the 2017 stocks, whereas the refactored script only took about 0.2969 seconds. Although the original script worked well enough for just 12 stocks, it would be better to use the refactored script for datasets with a large number of stocks. The main difference that helped speed up the analysis was treating the outputs as arrays. We created three arrays for our purposes:
```
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
```
These arrays allowed us to create more efficient for loops that would loop through all the rows in the spreadsheet. 

## Summary
 
- What are the advantages or disadvantages of refactoring code?

A clear advantage to refactoring code is coming up with a more efficient way to address a problem. Refactoring also allows for easier readability due to the fewer steps and improved logic. However, one disadvantage that I can think of for a coder would be having to think about addressing the same problem differently. Once you have already come up with a solution, sometimes it can be hard to step back and think of a different approach.

- How do these pros and cons apply to refactoring the original VBA script?

By refactoring the script, the analysis process was definetly made more efficient, as can be seen by the difference in times. Overall, it also seemed like a more elegant solution. The con doesn't really apply in this situation since we were given a starter code, so I didn't meet that roadblock. 
