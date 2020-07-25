# Stock-analysis

## Overview of Project

Our client, Steve, is analyzing stock data from the last couple of years, specifically 2017 and 2018, to look at trends among green energy companies for his parents who are interested in investing. He has created a [excel file](https://github.com/Stewartsl17/Stock-analysis/blob/master/VBA_Challenge.xlsm) with different stocks including the one that his parents are most interested in, DAQO New Energy Corp, which has the ticker "DQ". He has asked for our help with his analysis due to the large size of the data set. We have utilized VBA in order to automate the task and help him to look into various years in a quicker timeframe. However, after seeing the speed of the inital analysis of DQ, he wants to expand the dataset and find trends there. 

## Purpose 

The purpose of the second task was to not only provide the same information for other stocks as we had for DQ, but to test refactoring of initial VBA script to assist Steve with his interest into diversifying his parents' portfolio. 

## Results

As stated before, we utilized VBA programming to vizualize and analyze stock data from 2017 and 2018. Our worksheets includes the following colums of data: Stock Name (by ticker), Total Volume of stock, and Stock Return. 

Two main analyses were conducted during our exercise: 

*2017 vs. 2018 stock performance 
*Execution times between our intial VBA script of all stocks that we explored versus the refactored VBA script

### 2017 vs. 2018 Stock Performance

#### Image 1: 2017 Stocks
![](https://github.com/Stewartsl17/Stock-analysis/blob/master/Resources/Stock_Performance_2017.png)

#### Image 2: 2018 Stocks
![](https://github.com/Stewartsl17/Stock-analysis/blob/master/Resources/Stock_Performance_2018.png)

Based on the stock performances in 2017 and 2018, which can be seen in the images above, Steve should first let his parents know that the stock that they were intially interested in, "DQ", should be traded as it's trending downwards after 2017. They were seeing a huge positive return (+199.4%), but that became a fairly negagtive (-62.6%) return. IF he is looking to help them replace their investment with other stocks, the data suggests that "ENPH" and "RUN" are stocks that are either holding steady or positively trending over these two years. From a general perspective, by 2018, most of the stocks that he was looking into were trending negatively.

### Execution times between inital VBA script and refactored script

#### Image 1: Exection time for inital script for 2017 stock data
![](https://github.com/Stewartsl17/Stock-analysis/blob/master/Resources/VBA_Analysis_2017.png)

#### Image 2: Exection time for refactored script for 2017 stock data
![](https://github.com/Stewartsl17/Stock-analysis/blob/master/Resources/VBA_Challenge_2017.png)

#### Image 3: Exection time for inital script for 2018 stock data
![](https://github.com/Stewartsl17/Stock-analysis/blob/master/Resources/VBA_Analysis_2018.png)

#### Image 4: Exection time for refactored script for 2018 stock data
![](https://github.com/Stewartsl17/Stock-analysis/blob/master/Resources/VBA_Challenge_2018.png)

As seen in the images for 2017 above, the intial script for the stock data ran in .66 seconds. Comparatively, the refactored script ran in about .11 seconds. Looking at the 2018 execution times, the intial script ran in .68 seconds. However, the refactored script ran in .10 seconds. 

## Summary 


