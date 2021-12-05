# Stock-Analysis
## **Overview of the Project**
#### Steve enjoys using the last tool that our team provided to analyze the performance of stock ticker "DQ" and has been able to adivse his parents on this investment.  His parents have inquired about expanading their portfolio and want to know if our team can help.  The purpose of this project was to develop a tool for Steve to analyze the entire stock market in reacent years.  

## **Results**
### *Evalute the data* 
#### Using the dataset provided we can analyze performance of up to 12 different stocks from January through December in 2017 and 2018.  Instead of creating a seperate tool for each stock type, we will refractor the code to loop through the data once and collect all of the information.  The refractored code will run faster than running the code in Steves QQ Analysis tool. 
### *Prepare the Analysis using VBA in Microsoft Excel*
#### Locate the green stocks dataset and save the file as VBA_Challenge.xlsm and add a new worksheet.  Title the worksheet 'All Stocks Analysis".  Open the file AllStocksAnalysisReafractored.vbs using the VBA editor and rename in 'AllstocksAnalysisRecfratoredChallenge.vbs'.
### *Create the ticker index variable and output arrays for the Analysis*
####  There are several features in the original subroutine that do not need to be changed.  The code used to calculate the length of time to run can retained.  The start time and end time have been dimensioned and the timer is active in the existing code.  The 'All Stocks Analysis' worksheet has been activated and the code to output the year for all stocks has been udpated to remove the hard coding for a specific year (2018) to the use of the 'yearvalue' variable.  Tickers is dimensioned As String and there are 12 tickers from index 0 to 11 initialized for the tickers array. After activiting the Year value worksheet, the code to analyze all data in the worksheet can be kept from the original script.  
####  Instead of running an individual analysis of each ticker as in the original script  (                  )  we will create a ticker index variable which will allow us to perform the same calculations for every stock ticker in the dataset.   Next create the output arrays needed to complete the analysis.  There are three outputs that must be calculated to determine the total daily volume and percentage return for each stock, the total volume 'tickerVolumes', the starting price 'TickerStartingPrice' and the ending price 'TickerEndingPrice'.  Dimension TickerVolumes as Long and Ticker StartingPrices and TickerEndingPrices as Single.  Each array must be able to be analyzed for each ticker, therefore we tell the code to index 12 times using (12) for each array.  The new code can be seen below in Figure 1:
![image](https://user-images.githubusercontent.com/94234511/144731625-93ae57b9-32c0-42b0-ab56-18ee6e4e0963.png)
 
### *Create a for loop to intialize the tickervolumes array and set it to zero*
####  To perform the same analysis for each ticker, we must loop through all of the data for each ticker, adding the volume each time the stock is traded to calculate the total volume. Use i to define each ticker index.  Also, to make sure our analysis is correct each time the tool is used, we add the code to set the ticker volume for each ticker to zero prior to summing the total volume.  Figure 2 illustrates the code to initialize the ticker volumes array. 
![image](https://user-images.githubusercontent.com/94234511/144732615-e8a80872-b3d9-484f-a015-a4d15330c7b2.png)

### *Create a for loop that will analyze Total volume for all stocks in the dataset*
####  Set the code to loop over all the rows in the spreadsheet so that each ticker is analyzed.  In order to determine the total daily volume for each stock, we must define the formula to calculate the total daily volume by adding the values of all cells in column H (Volume) for each ticker index "i". Column H is position 8 in the 2017 and 2018 worksheets. The correct code to 
![image](https://user-images.githubusercontent.com/94234511/144732679-cd218bc3-4dd9-4185-b071-572d7288e977.png)

### *Using conditional statements, calculate the TickerStartingPrice and TickerEndingPrice, indexing the ticker for each stock*
####  Conditional statements make performing calculations quickly and use logical statements to return the output.  For example, we want to determine the starting price and ending price for each ticker and then use those values to calculate the percent return for each ticker.  Conditional statements may be used within the for loop to perform calculations using the ticker index variable such that matching ticker values are used in the formula but those that do not match are excluded from the calculation of each output array.  Using the 'If then' conditonal statement, the code to calculate starting price and ending price for each ticker can be seen in Figure 4. Note that the ticker index must be incremented to the next ticker (tickerindex +1) or the code will only run for one row.  
![image](https://user-images.githubusercontent.com/94234511/144733387-f9859d47-faf6-4967-83e4-53ba4daabe4b.png)

### *Create a for loop to analyze all data and output the Total Daily Volume and Percent Return for each Stock ticker*
####  Now, we are ready to analyze all the data and generate the output for each array. The Ticker column will list the ticker code for each green stock in the dataset.  The Total Daily Volume column will include the sum of all trading (tickerVolumes) for each stock. The Return column will reflect the profit or loss for each ticker based on the percentage of increase in ending price over starting price.  Activate the 'All Stocks Analysis' Worksheet to ensure the output is correctly placed in the file.  Using the for loop, loop through each array to output Ticker, Total Daily Volume and Return in the first three columns of the sheet, with headers placed in Row 3. The code will appear as illustrated in Figure 5 below. 
![image](https://user-images.githubusercontent.com/94234511/144733897-38ce04dc-f3ff-45ff-8aeb-5a632d89f9a3.png)

### *Comparison of Green Stock performance in 2017 vs 2018*
####  The formatting used in the original script is helpful in visualizing the data and Steve agreed that we use similar formatting for this project, so no changes to the formatting code were needed.  Save the refactored code to VBA_Challenge.xlsm and click the forward arrow icon in the toolbar to run the subroutine 'AllStocksAnalysisRefactoredChallenge()'.  A window prompt will be visible asking the user "What year would you like to run the analysis on?".  Enter "2017" and press "Ok".  The output populated on the All stocks analysis can be seen in Figure 6 below. 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/94234511/144734351-a0179be5-9d78-4a7f-bbd2-40eceadcb06a.png)
###  The analysis indicates that all green stocks had a profitable return in 2017 except for "TERP".  The return for "TERP" indicates a -7.2% return, indicating the total endingPrice for this ticker was lower than the total startingPrice during 2017.  Total daily volume for "TERP" (139,402,800) was lower than average (263, 886, 592) in 2017. Of the remaining eleven (11) green stocks, four (4) tickers "DQ", "ENPH", "FSLR", and "SEDG" showed a greater than 100% return, more than doubing in price; while two (2) stocks yielded at least a 50% or greater return.  Stock tickers "AY" and "RUN" showed the lowest profit (<10%) on their returns in 2017.  The total daily volumes were near the average for "RUN" while "AY" had much lower trading volume. Overall, the analysis indicates that green stocks were successful in 2017.
#### When we repeat the analysis using the data from 2018, the results look much different for investments in the same green stocks compared to 2017.  See the analysis for all green stocks in 2018 in Figure 7 below.  
![VBA_Challenge_2018](https://user-images.githubusercontent.com/94234511/144735340-187d6d95-7fc7-4e91-b0d6-92595f371482.png)
####  With the exception of "ENPH" and "RUN" with profitable returns of 80% or better, all tickers showed a negative return in 2018. Returns of -62.6% and -60.5% were calculated using this method for tickers "DQ" and "JKS". The total daily volume for "ENPH" nearly tripled in 2018 and exceed all other tickers by 70 million shares in that same calendar year. Of the four tickers that doubled in price in 2017, only "ENPH" showed a positive return in both 2017 and 2018.  

####  Overall green stocks performed better in 2017 than 2018.  Total daily volumes for all stocks were similar in 2017 and 2018, although the differences in total volume from year to year varied for each ticker.  The assumption that there is a strong correlation between total daily volume and percent return was not realized in either 2017 or 2018 for the green stocks dataset. Additional analysis would be needed to determine the causes of the negative returns for most green stocks in 2018.    
### *Compare exection times of the original and refractored script*
####  The execution time was 0.1640625 seconds to analyze the output using the refactored code in the challenge exercise. The original script required XX seconds to perform the analysis of just the ticker "DQ" using the data for 2017, see Figure 8 below.  The refractored script outputs the data for each ticker in a fraction of a second compared to almost 2 seconds using the original code. Similarly, when the execution time for the 2018 analysis is compared, we find the refactored code running in 0.1484375 seconds compared to 1.xx seconds in the original script.  See the run time for the 2018 analysis for ticker "DQ" in Figure 9.    
## **Summary**
### *Advantages and disadvantages of refactoring code*
#### Refactored has two primary advantages, the data can be analyzed quickly providing more rapid analysis of outcomes. This may be useful if data must be analyzed frequently or when speed is needed to allow quick decision.  When looking to protect an investment, speed in analyzing the data would result in timely decision making for stock purchase or sale. For datasets that are updated and analysis that are repeated at some frequency, refatored code is preferred as hard coding is often removed in favor of readable code.  Readable code is easier to follow after time has passed or when others are reviewing the analysis tool and reading the code.  Using logical arguments to quickly index through multiple rows of data at one time rather than repeating a series of analysis for indvidual rows is more efficent. The only disadvantage is that refactoring code takes programming time and errors during refactoring need to be debugged prior to completing the analsyis.  
### *How does this apply to refactoring the original script.*  
#### After refactoring the code for this challenge exercise, I had two bugs to remove before it would run.  I experienced a couple of minor bugs after refactoring the code.  First, the error message "Run-time error 9: Subscript out of range" was recieved.  The code did not read the new worksheet, "All Stocks Analysis" until I closed and reopened Excel.  After a successful intiation of the macro and entry of 2017 in the prompt, another error was identified related to the placement of the End If in the first conditional statement of the for loop in step 3c.
