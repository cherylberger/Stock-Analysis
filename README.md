# Stock-Analysis
## Overview of the Project
### Steve enjoys using the last tool that our team provided to analyze the performance of stock ticker "DQ" and has been able to adivse his parents on this investment.  His parents have inquired about expanading their portfolio and want to know if our team can help.  The purpose of this project was to develop a tool for Steve to analyze the entire stock market in reacent years.  

## Results
### Evalute the data 
#### Using the dataset provided we can analyze performance of up to 12 different stocks from January through December in 2017 and 2018.  Instead of creating a seperate tool for each stock type, we will refractor the code to loop through the data once and collect all of the information.  The refractored code will run faster than running the code in Steves QQ Analysis tool. 
### Prepare the Analysis file
#### Locate the green stocks dataset and save the file as VBA_Challenge.xlsm and add a new worksheet.  Title the worksheet 'All Stocks Analysis".  Open the file AllStocksAnalysisReafractored.vbs using the VBA editor and rename in 'AllstocksAnalysisRecfratoredChallenge.vbs'.
### Create the ticker index variable and output arrays for the analysis
#### There are several features in the original subroutine that do not need to be changed.  The code used to calculate the length of time to run can retained.  The start time and end time have been dimensioned and the timer is active in the existing code.  The 'All Stocks Analysis' worksheet has been activated and the code to output the year for all stocks has been udpated to remove the hard coding for a specific year (2018) to the use of the 'yearvalue' variable.  Tickers is dimensioned As String and there are 12 tickers from index 0 to 11 initialized for the tickers array. After activiting the Year value worksheet, the code to analyze all data in the worksheet can be kept from the original script.  Instead of running an individual analysis of each ticker as in the original script  (                  )  we will create a ticker index variable which will allow us to perform the same calculations for every stock ticker in the dataset.   Next create the output arrays needed to complete the analysis.  There are three.  Dimension TickerVolumes as Long and Ticker StartingPrices and TickerEndingPrices as Single The new code can be seen below in Figure 1: ![image](https://user-images.githubusercontent.com/94234511/144731625-93ae57b9-32c0-42b0-ab56-18ee6e4e0963.png)
 
### Create a for loop to intialize the tickervolumes array and set it to zero
### Create a for loop that will analyze Total volume for all stocks in the dataset
### Using conditional statements, calculate the TickerStartingPrice and TickerEndingPrice, indexing the ticker for each stock
### Create a for loop to analyze all data and output the Total Daily Volume and Percent Return for each Stock ticker
### Comparison of Green Stock performance in 2017 vs 2018
### Compare exection times of the original and refractored script
## Summary
### Advantages and disadvantages of refactoring code
### How does this apply to refactoring the original script.  
