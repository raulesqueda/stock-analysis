# stock-analysis
Analysis and exercises part 2

# **Overview of Project: Stock analysis for Steve**

## Overview of Project

In this project, we’ll help to Steve's parents to analyze the stock market in 2017 and 2018. For this, we use the VBA in excel to get a report of the value of the stock and the performance during the year, one way to measure this is to calculate the yearly return for daily stock. Also, to help in this analysis, we use conditional format to know more faster the stock who has best and the worst performance and adding buttons to start the analysis and other to clear the analysis.

## Coding in VBA

We start this analysis with the code called “AllStockAnalysis_Final”, were we create three variables “Ticker”, “Total Daily Volume” and “Return” for eleven different of stocks. 

![Code 1 Part 1](https://github.com/raulesqueda/stock-analysis/blob/main/Code%201%20part%201.PNG)

Then we calculate the starting Price and the ending Price of each stock we define in the last step. After that, we start to calculate the total Volume and the return of the stocks, to decide which had the best performance.

![Code 1 Part 2](https://github.com/raulesqueda/stock-analysis/blob/main/Code%201%20part%202.PNG)
 
For 2017, the result was this:

![Results 2017 Code 1](https://github.com/raulesqueda/stock-analysis/blob/main/Results%202017%20Code%201.png)
 
For 2018, the result was:

![Results 2018 Code 2](https://github.com/raulesqueda/stock-analysis/blob/main/Results%202018%20Code%201.png)

We had the information we need to evaluate the performance of the daily stocks, but the presentation lacks format to help to Steve’s parents. For this reason, we add to the code we start to program conditional formatting for the results, and we calculate an index for the stock market. Using the code in the previous steps, we add a code for the index:
 
![Code 2 Part 1](https://github.com/raulesqueda/stock-analysis/blob/main/Code%202%20part%201.PNG)
 
Then we create three more arrays to the analysis, this includes the Volume of the stock, and de Starting and Ending Prices of each:
 
![Code 2 Part 2](https://github.com/raulesqueda/stock-analysis/blob/main/Code%202%20part%202.PNG)
 
Then we started a loop to increase the volume of the stock:

![Code 2 Part 3](https://github.com/raulesqueda/stock-analysis/blob/main/Code%202%20part%203.PNG)
 
After this, we stablish assigning the starting price in the first row of the Index:

![Code 2 Part 4](https://github.com/raulesqueda/stock-analysis/blob/main/Code%202%20part%204.PNG)
 
Then, we stablish assigning the ending price in the last row of the Index, in the cause the row doesn’t match, we increase the Index:

![Code 2 Part 5](https://github.com/raulesqueda/stock-analysis/blob/main/Code%202%20part%205.PNG)
 
To fill the values of the three variables we create in steps before, we create a loop:

![Code 2 Part 6](https://github.com/raulesqueda/stock-analysis/blob/main/Code%202%20part%206.PNG)
 
Finally, we are coding the conditional formatting for the return:

![Code 2 Part 7](https://github.com/raulesqueda/stock-analysis/blob/main/Code%202%20part%207.PNG)

## Results
By adding this code, we continue with the evaluation of the code and the return of the stocks, the performance of 2017 was:

![Results 2017 Code 2](https://github.com/raulesqueda/stock-analysis/blob/main/Results%202017%20Code%202.PNG)

The performance for 2018 was:

![Results 2018 Code 2](https://github.com/raulesqueda/stock-analysis/blob/main/Results%202018%20Code%202.PNG)
 
Now we can easily analyze the stocks we have for the two years using the code we generate. The stocks called ENPH and RUN in both years had a positive performance, maybe Steve’s parents could invest in these two stocks to gain some money. Also we can see that the part we add to the code sightly increase the amount of time to the code to be executed, but we think this is not a problem if we increase the amount to information in the future.

## Summary

As result, we develop a better code for analyzing the stock market for Steve’s parents, we consider that using the first code help with the analysis and the results because we use the previous work and we adapt what we had, reducing the time to make a new code and, knowing the structure of the code, we could adapt the code.

In the case of coding in VBA, VBA is part of a well-known software, this is an advantage because almost everyone know it, this is Microsoft Excel, also, many computers has this software, people only have to learn to use it. 

The disadvantage is the time you must use to learn how to code in VBA, we have to remember that many projects lacks time and the team had to use the time efficiently. 
