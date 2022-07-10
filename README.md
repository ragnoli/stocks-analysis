# stocks-analysis

# Overview of Project
   The purpose of this project is to refactor the code prepared in Visual Basic for Applications to analizes multiple stocks to find what is the best investment option for Steve's parents.
   Steve wants to expand the dataset to include the entire stock market over the last few years. Using the previous coding it may take a long time to execute, therefore, we will refactor the Sub "All Stock Analysis" to loop through all the data one time in order to collect the same information, however, the refactored code should be more efficient—by taking fewer steps and use less memory, as a consequence the macro should run faster.


# Results

### Refactored code explained:

In the refactored code we created arrays for the volume, starting and ending price of the tickers. An array for the tickers were already created in the original file.

    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

We also created the variable "tickerindex" used  to access the correct index across the arrays:

    tickerindex = 0
    
We first created a for loop that will loop over all the rows in the spreadsheet, then inside the loop we created scripts to determine the volume, starting and ending price of the tickers and determine the value for the arrays created before.

 tickerVolumes:
 
      tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
  
 tickerStartingPrices:
 
      If Cells(i - 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
           tickerStartingPrices(tickerindex) = Cells(i, 6).Value
    
tickerEndingPrices:

      If Cells(i + 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
           tickerEndingPrices(tickerindex) = Cells(i, 6).Value
           
 To alocate the value to the correct array index, we wrote a scrip that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker:

      If Cells(i + 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
            tickerindex = tickerindex + 1
           
Finally, to bring the results to the All Stock Analysis sheet, we first activtated the sheet and then we used a for loop to loop through the arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in the spreadsheet.
 
      Cells(4 + i, 1).Value = tickers(i)
      Cells(4 + i, 2).Value = tickerVolumes(i)
      Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
   
   
### Execution time
Refactoring the code to automate tasks decrease the chance of errors and reduce the time needed to run the analyses.
Creating arrays for the volume, starting and ending price made the code more efficient and fast.







