# stock-analysis
stock-analysis
# Stock-Analysis
## Overview of Project
In this project I created an easy to view VBA so Steve can analyze some green stocks performance showing daily volume and annual return. I used macros with 2017 and 2018 data and highlighted rate of returns which should help him narrow down the stocks he would like to invest in. I refactored my code in order to increase the efficiency of my analysis. 
## Results 
I came to several conclusions based on the data given to me by using tickers to determine the type of stock that was being displayed. The stock categories from 2017 tended to have a successful return. However, in the year 2018, only two of the twelve stock types had a positive return.
### Refactoring the Code
I had to switch the nesting order of my loops in order to make my code more efficient. I created four arrays consisting of tickers, tickervolumes, tickerstartingprices, and tickerendingprices. I used tickers array to establish the ticker symbol of each stock. I matched up the remaining arrays with ticker array with the variable tickerIndex. Using this variable made it easier for me to assign tickervolumes, tickerstartingprices, and tickerendingprices to each stock’s ticker symbol. I found this method to be much more efficient compared to nested loop. 
#### Refactored Code
    '2)Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
     '1a) Create a ticker Index
      
    tickerIndex = 0

    '1b) Create three output arrays
    
    ReDim tickerVolumes(12) As Long
    ReDim tickerStartingPrices(12) As Single
    ReDim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
    
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
    
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
    
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
            

            '3d Increase the tickerIndex.
                
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerIndex = tickerIndex + 1
        
        End If
    
    Next i
    
      '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
### Original Code
      '2) Initialize array of all tickers
      
      Dim tickers(11) As String

      tickers(0) = "AY"
      tickers(1) = "CSIQ"
      tickers(2) = "DQ"
      tickers(3) = "ENPH"
      tickers(4) = "FSLR"
      tickers(5) = "HASI"
      tickers(6) = "JKS"
      tickers(7) = "RUN"
      tickers(8) = "SEDG"
      tickers(9) = "SPWR"
      tickers(10) = "TERP"
      tickers(11) = "VSLR"

      '3a) Initialize variables for starting price and ending price

      Dim startingPrice As Single
      Dim endingPrice As Single

      '3b) Activate data worksheet

      Worksheets("2018").Activate

      '3c) Get the number of rows to loop over

      RowCount = Cells(Rows.Count, "A").End(xlUp).Row

      '4) Loop through tickers
   
       For i = 0 To 11
      ticker = tickers(i)
      totalVolume = 0
         
      '5) loop through rows in the data
   
      Worksheets("2018").Activate
    
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
       
       '6) Output data for current ticker
       
       Worksheets("All Stocks Analysis").Activate
       
       Cells(4 + i, 1).Value = ticker
       
       Cells(4 + i, 2).Value = totalVolume
       
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

      Next i
   
## Comparing the Original Run Times to the Refactored Run Times

The run-times using the original code.

![2017 Original Run-time](https://github.com/meliscelikay/stock-analysis/blob/7e468c016fd27df4739faf0f7d6a51c2faac23dd/Resources/Original-2017.png)

![2018 Original Run-time](https://github.com/meliscelikay/stock-analysis/blob/7e468c016fd27df4739faf0f7d6a51c2faac23dd/Resources/Original-2018.png)

The run-times using the refactored code.

![2017 Refactored Run-time](https://github.com/meliscelikay/stock-analysis/blob/7e468c016fd27df4739faf0f7d6a51c2faac23dd/Resources/Refactoredcode-2017.png)

![2018 Refactored Run-time](https://github.com/meliscelikay/stock-analysis/blob/7e468c016fd27df4739faf0f7d6a51c2faac23dd/Resources/Refactoredcode-2018.png)

# Summary
## Advantages or Disadvantages of refactoring code
### Advantages
  Refactoring code is advantages because it is more efficient and easier to debug if there is an issue with the VBA code. It is making the run time faster. This can be very important especially when working with extremely large datasets.
### Disadvantages
  The disadvantage of refactoring code is it is not a new code so there is always the possibility of fixing additional bugs if it is refactored.  

# Advantages and disadvantages of the original and refactored VBA script 
  I noticed that refactored code reduced the number of loops which increased efficiency. Refactored code also maximized the VBA code which led to less run time if we compare it to the original code. Although I favored the refactored code, one can make a case for added runtime and more time spent on debugging if the data set and the analysis was more complex.  

