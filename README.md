# stock-analysis
## Module 2

### Overview of Project:
The purpose of this is data workbook is to help Steve analyze the enrtire set of data. We are expanding the dataset to include the entire stock market over the last few years. Hopefully this will help Steve read and understand the data better.

### Results:
![myTest](https://github.com/nfreeman19/stock-analysis/blob/main/resources/2017.png) 
![myTest](https://github.com/nfreeman19/stock-analysis/blob/main/resources/2018.png)

The images above are the results after running 2017 and 2018.

Execution times were both very quick. 

As you can see from the postive numbers in 🟢 and the negative numbers in 🔴 the returns were better ad higher in 2017. Returns in 2018 we're mostly negative.

When refactoring the results we needed to create three output arrays
     
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
        
When running these results the helpful code letting us loop through the arrays to output the Ticker, Total Daily Volume, and Return was:

        For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
    
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1


### Summary:
The advantages of refactoring code is code is new and most likely easier to understand and read. The code is less complex and easier to maintain. 
Disadvantages of refeactiong code is that is can take a good amount of time. You may not know how much time it will end up taking. You may also be confused on where to go with the data.

How do these pros and cons apply to refactoring the original VBA script?
We refactored the first data sheet we made for Steve. We determined that refactoring our code successfully made the VBA script run faster. We did not add any new functionality; we just made the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code. This made it easier for Steve to review and understand.
