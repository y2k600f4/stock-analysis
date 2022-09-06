# Stock Analysis-analysis (VBA of Wallstreet)

##Overview of Project

The purpose of this analysis was to refactor VBA code that originally analyzed 12 individual stock market stocks over (2) Years (2017 and 2018 data sets). The original code worked well, however it was inefficient and would take excessive time to analyze a larger dataset that would include the entire stock market over several years.  By refactoring the original code the goal was to successfully make the VBA script run faster to improve the efficiency.

##Results
###Original Code
In the original VBA script a loop was done for each ticker symbol which resulted in the code looping through all the rows for each ticker symbol to extract the appropriate data ( in this case 12 times). The sample code of interest can be seen below:

'4) Loop through tickers   
   For i = 0 To 11   
       ticker = tickers(i)       
       totalVolume = 0       
       '5) loop through rows in the data       
      'Worksheets("2018").Activate      
       'replace hard-coded year value
        Sheets(yearValue).Activate       
       For j = 2 To RowCount       
           '5a) Get total volume for current ticker           
           If Cells(j, 1).Value = ticker Then
               totalVolume = totalVolume + Cells(j, 8).Value
           End If           
           '5b) get starting price for current ticker           
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               startingPrice = Cells(j, 6).Value
           End If
           '5c) get ending price for current ticker           
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               endingPrice = Cells(j, 6).Value
           End If
       Next j

###Refactored Code
In the refactored code an array was created and all data was extracted into an array by looping through all the rows 1x (all stock symbols data was extracted without the need to loop through the data set multiple times). The sample code of interest can be seen below:
'2b) Loop over all the rows in the spreadsheet.    
    For i = 2 To RowCount 'i is cell reference    
        '3a) Increase volume for current ticker        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value    
            
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then            
        'open column is tickerStartingPrice-column3        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value            
        End If        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then        
        'close column is tickerEndingPrice-column 6            
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            '3d Increase the tickerIndex.            
        tickerIndex = tickerIndex + 1           
            
        End If    
    Next i

###Analysis of 2017 Stock Dataset
Below shows the time improvement from the original VBA script vs the refactored VBA script with a time saving of ~.75 or ~15% of the original time.


![Timing of Analysis  of 2017 Stock VBA Code](https://github.com/y2k600f4/stock-analysis/blob/main/Resources/time_2017.png)

![ Timing of Analysis  of 2017 Stock VBA Code-Refactored](https://github.com/y2k600f4/stock-analysis/blob/main/Resources/time_2017_refactored.png)

###Analyss of 2018 Stock Dataset

Below shows the time improvement from the original VBA script vs the refactored VBA script with a time saving of ~1.1s or ~11% of the original time.

![Timing of Analysis  of 2018 Stock VBA Code](https://github.com/y2k600f4/stock-analysis/blob/main/Resources/time_2018.png)

![ Timing of Analysis  of 2018 Stock VBA Code-Refactored](https://github.com/y2k600f4/stock-analysis/blob/main/Resources/time_2018_refactored.png)


###Summary

The advantages of refactoring code in general include improving programs to run faster, cleans up the code making it easier to understand (improves logic), helps find bugs and uses less memory. Some disadvantages of refactoring code in general is it time intensive (especially on larger program) and could be risky if not properly tested.  For this refactoring example the advantage has been clearly shown that the program runs significantly faster. One disadvantage of refactoring this VBA script is that further testing and error prevention was not done to prove the code’s flexibility.			







