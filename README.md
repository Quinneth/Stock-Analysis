# Stock-Analysis
## Overview of Project
This project revisits previous stock analysis code using VBA macros.
1. Provide technical stock analysis and results report.
2. Refactor VBA code to broaden analsysis scope and attempt to improve performance measure.

### Purpose
The purpose of the analysis is to expand the dataset to include the entire stock market for years 2017 and 2018 and measures the duration of execution times for comparison to orignial script. Using knowledge of VBA and the already code provided; scritp must be written during to collect all the information looping through the dataset only once. Specifically, this project aim to:
- Compare stock performance between 2017 and 2018
- Document execution time comparisons between original and refactored script
- Evaluate the pros and cons of applying refactored code to original VBA script
 
## Analysis
This analsyis demonstrates stock performance via yearly returns for years 2017 and 2018. The yearly return is the percentage change in price from year start to end, indicating growth.

#Methodical steps with including code:

1. Activate the specified worksheet, determine the number of rows involved in the loop, then reate a tickerIndex variable and set it equal to zero before iterating over all the rows. The tickerIndex is used to access the correct index across the four different arrays: the tickers array and the three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.<br/>

    Activates data worksheet:
    
        Worksheets(yearValue).Activate
    
    Get the number of rows to loop over:
    
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    Creates a ticker Index:
    
        tickerIndex = 0

    Creates three output arrays:
        
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
        
    Creates a for loop to initialize the tickerVolumes to zero.
    
        For i = 0 To 11
        tickerVolumes(i) = 0         
        Next i
          
    Loops over all the rows in the spreadsheet:
    
         For j = 2 To RowCount
         
     To increase volumes for:
     
         If Cells(j, 1).Value = tickers(tickerIndex) Then
          
2. Creates a for loop to initialize the tickerVolumes to zero. If the next row's ticker doesn't match, increase the ticker index.
    Loop over all the rows in the spreadsheet:
       
        For j = 2 To RowCount
        
    Increases tickerIndex in next row does not match:
     
        If Cells(j, 1).Value = tickers(tickerIndex) Then


3. Creates a for loop that will loops and reads over the rows in the spreadsheet, then stores all data values from each array by finding starting price for the cufrrent ticker using multiple conditions.<br/>
- Writes conditional script, inside the for loop, that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticke and stores as ending price within the worksheet. Uses the tickerIndex variable as the index. Writes an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current closing price to the tickerStartingPrices variable:
    
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
          End If<br/>
        Store start price
          tickerStartingPrices(tickerIndex) = Cells(j, 6).Value <br/>
        
- Write an if-then statement to check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable
If so-->assigns current price to the ending price variable:

        If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
        Stores Ending Price Value
          tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
          tickerIndex = tickerIndex + 1 
          End I         
         Next j

4. Use a for loop to loop through your arrays (tickers, tickerVolumes, tickerStartingPrices, and 
tickerEndingPrices) to format the output the of “Ticker,” “Total Daily Volume,” and “Return” columns:

        Headers
        Worksheets("All Stocks Analysis").Activate
        "cells()" is easier than "range()"-->this prints on the 4th row + i
            For i = 0 To 11
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
         Next i
 5. Formatting:
   
            Range("A3:C3").Font.FontStyle = "Bold"
            Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
            Range("B4:B15").NumberFormat = "#,##0"
            Range("C4:C15").NumberFormat = "0.0%" 'Uses 1 digit percision in return amount
            
- Adds more zeros for more digits:
  
        Columns("B").AutoFit

- Creates loop to color data:

          dataRowStart = 4
          dataRowEnd = 15
            For i = dataRowStart To dataRowEnd
            'inside loop
            If Cells(i, 3) > 0 Then
            'Change cell color to green
            Cells(i, 3).Interior.Color = vbGreen
            
            ElseIf Cells(i, 3) < 0 Then
            'Change color of cell to red
            Cells(i, 3).Interior.Color = vbRed
            
            Else
            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone
            
        End If
      Next i

###Results
Improving perfomance efficiency involves reducing steps, using less memory, or improving logic to make it easier for future users to interpret.<br/>
#Comparison output of 2017 and 2018:
***Original 2017 "All Stocks Analysis" <br/> ![TheaterLaunchdate]
***Refactored 2017 "All Stocks Analysis" <br/> ![TheaterLaunchdate]
***Original 2017 "All Stocks Analysis" <br/> ![TheaterLaunchdate]
***Refactored 2017 "All Stocks Analysis" <br/> ![TheaterLaunchdate]
![TheaterLaunchdate](https://github.com/Quinneth/Kickstarter-analysis/blob/main/Theater_Outcomes_vs_Launch.png)

### Challenges and Difficulties Encountered



### Analysis of Outcomes Based on Goals


### Challenges and Difficulties Encountered
