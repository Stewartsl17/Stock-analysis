# Stock-analysis

## Overview of Project

Our client, Steve, is analyzing stock data from the last couple of years, specifically 2017 and 2018, to look at trends among green energy companies for his parents who are interested in investing. He has created a [excel file](https://github.com/Stewartsl17/Stock-analysis/blob/master/VBA_Challenge.xlsm) with different stocks including the one that his parents are most interested in, DAQO New Energy Corp, which has the ticker "DQ". He has asked for our help with his analysis due to the large size of the data set. We have utilized VBA in order to automate the task and help him to look into various years in a quicker timeframe. However, after seeing the speed of the initial analysis of DQ, he wants to expand the dataset and find trends there. 

## Purpose 

The purpose of the second task was to not only provide the same information for other stocks as we had for DQ, but to test refactoring of initial VBA script to assist Steve with his interest into diversifying his parents' portfolio. 

## Results

As stated before, we utilized VBA programming to visualize and analyze stock data from 2017 and 2018. Our worksheets include the following columns of data: Stock Name (by ticker), Total Volume of stock, and Stock Return. 

Two main analyses were conducted during our exercise: 

*2017 vs. 2018 stock performance 
*Execution times between our initial VBA script of all stocks that we explored versus the refactored VBA script

### 2017 vs. 2018 Stock Performance

#### Image 1: 2017 Stocks
![](https://github.com/Stewartsl17/Stock-analysis/blob/master/Resources/Stock_Performance_2017.png)

#### Image 2: 2018 Stocks
![](https://github.com/Stewartsl17/Stock-analysis/blob/master/Resources/Stock_Performance_2018.png)

Based on the stock performances in 2017 and 2018, which can be seen in the images above, Steve should first let his parents know that the stock that they were initially interested in, "DQ", should be traded as it's trending downwards after 2017. They were seeing a huge positive return (+199.4%), but that became a fairly negative (-62.6%) return. IF he is looking to help them replace their investment with other stocks, the data suggests that "ENPH" and "RUN" are stocks that are either holding steady or positively trending over these two years. From a general perspective, by 2018, most of the stocks that he was looking into were trending negatively.

### Execution times between initial VBA script and refactored script

#### Image 1: Execution time for initial script for 2017 stock data
![](https://github.com/Stewartsl17/Stock-analysis/blob/master/Resources/VBA_Analysis_2017.png)

#### Image 2: Execution time for refactored script for 2017 stock data
![](https://github.com/Stewartsl17/Stock-analysis/blob/master/Resources/VBA_Challenge_2017.png)

#### Image 3: Execution time for initial script for 2018 stock data
![](https://github.com/Stewartsl17/Stock-analysis/blob/master/Resources/VBA_Analysis_2018.png)

#### Image 4: Execution time for refactored script for 2018 stock data
![](https://github.com/Stewartsl17/Stock-analysis/blob/master/Resources/VBA_Challenge_2018.png)

As seen in the images for 2017 above, the initial script for the stock data ran in .66 seconds. Comparatively, the refactored script ran in about .11 seconds. Looking at the 2018 execution times, the initial script ran in .68 seconds. However, the refactored script ran in .10 seconds. 

## Summary 

### Advantages and Disadvantages of Refactoring Code

Advantages: Refactoring code is done in order to make the code simpler and cleaner. Another advantage of this process is the code typically performs when it's refactored. 

Disadvantages: Refactoring code can be problematic if errors and bugs are not taken care of correctly. This can lead to the overall process becoming more time consuming. Moreover, these problems become exponentially more problematic as the code becomes more complex. 

### Applying the Advantages and Cons to Refactoring this VBA script

When it comes to this code, the advantages and disadvantages noted above definitely occurred during the refactoring process. Errors and bugs we consistently present as we move different pieces of code around. Moreover, having an understanding of how different parts of the code fit next to each to each other is crucial. On the other hand, once the refactoring was complete, this code worked far faster than before. Not only this, but the code looked cleaner in comparison.
