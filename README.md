# The VBA Of Wallstreet

##Overview of Project

This is an analysis of 12 select stocks performance between the fiscal years of 2017 and 2018 for the purpose of determining which stocks might be a viable investment.  

In addition, this analysis seeks to determine if using refactored vba code can enhance the performance of excel and speed up the processing time of data.

##Results
Daily opening price, closing price, and total volume data for each of these twelve stocks was analysed to determine, over the course of 2017 and 2018 the total volume of each stock that was moved and the percentage return for 1 year of investment.

The initial analysis attempts to do this by:

Looping through each ticker:

       For i = 0 To 11
     	   ticker = tickers(i)
     	   totalVolume = 0

The, within each ticker, looping through all the rows of data:

       For j = 2 To RowCount

To determine the Total Volume traded for each stock:

        If Cells(j, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value

The starting price:

        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            startingPrice = Cells(j, 6).Value

And the ending price: 

        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            endingPrice = Cells(j, 6).Value

The refactored analysis instead focused on creating arrays for the output values desired.  

    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Long
    Dim tickerEndingPrices(12) As Single

Rather than using nested loops to sort through the data, the data for each ticker is now saved to an array, meaning that the macro only has to perform on level of looping.

    For i = 2 To RowCount

        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If

         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If
       
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
         End If

Both methods provide the exact same results.  The data analysis indicates that, for the 12 selected stocks, performance was overall very positive for the year of 2017, with 10 of the stocks having significant return on investment.  Even for the two stocks that did not result in profit, the losses were fairly low.

2018 appears to represent a poor year across the market for the stocks in this analysis.  Only two stocks provided a return on investment, and many of the other stocks fell deep into the red.  

Based on these two years of data, it would appear that ENPH, with its year-over-year returns, and SEDG, with very large 188.8% returns in 2017, and only a -7.6% loss in 2018, are the safest investments.

![2017-2018 Analyses](https://github.com/rscalise88/stock-analysis/blob/main/Resources/2017_2018_Output.PNG)

While the two methods produced the same output, the time taken to conduct the refactored analysis was significantly lowered.  Timers were placed within each analysis, set to display the processing time for each run.  On average the initial methodology took 0.68s to complete.  However, the messages produced after the refactored analyses indicated that these were often finished over twice as fast.

![2017 Processing Time](https://github.com/rscalise88/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)

![2018 Processing Time](https://github.com/rscalise88/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

##Summary 
For smaller data sets, the extra steps required to develop the code for the refactoring do not result in a significant amount of time saved versus time put into generating the code.  For an amateur coder, the fractions of a second shaved off the processing time are lost in writing and potentially troubleshooting the additional lines of code.

However, were this to be applied to a significantly larger data set, the processing time would heavily factor into the decision to use the refactored methodology.  Were the entire stock market, instead of 12 individual stocks, to be analyzed, rather than fractions of a second, hours and minutes could potentially be saved from the processing time of the analysis.
