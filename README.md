# Challenge 2 stock-analysis

Purpose
The purpose of this project is to analyze information on a group of stocks from 2017 and 2018.  The analysis was done in excel with VBA coding to automate the review of stocks to potentially invest in.  This challenge is to refactor the original code and make it more efficient by taking fewer steps, using less memory, or improving the logic of code.  
Results
The code for this project is listed below:
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
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
   For i = 0 To 11
        tickerIndex = tickers(i)

    '1b) Create three output arrays
    Dim tickerVolumes As Long
    Dim tickerStartingPrices, tickerEndingPrices As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
       Worksheets(yearValue).Activate
       tickerVolumes = 0
    
        
        '2b) Loop over all the rows in the spreadsheet.
        For j = 2 To RowCount

        'If the next row's ticker doesn't match, increase the tickerIndex
            If Cells(j, 1).Value = tickerIndex Then
    
               '3a) Increase volume for current ticker
               tickerVolumes = tickerVolumes + Cells(j, 8).Value
            End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
         
               tickerStartingPrices = Cells(j, 6).Value
            
          'End If
      End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
           If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then

               tickerEndingPrices = Cells(j, 6).Value

                            
            End If
    
       Next j
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
        
            Worksheets("All Stocks Analysis").Activate
        
            Cells(4 + i, 1).Value = tickerIndex
            Cells(4 + i, 2).Value = tickerVolumes
            Cells(4 + i, 3).Value = tickerEndingPrices / tickerStartingPrices - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
-	When comparing the original code to the new code there is a slight change in time needed to run the macro.  In 2017 the original code took 0.66 seconds, and the refactored code took 0.679 seconds. For 2018 the original code took 0.664 seconds, and the refactored code took 0.671 seconds.  Both run times are similar but if this was run on larger projects the time variance could become larger.
-	When comparing 2017 stock returns to 2018 stock returns, it looks like only 2 stocks have a positive return both years ENPH and RUN.  The stock with a negative return both years is TERP.
 ![image](https://user-images.githubusercontent.com/89753083/140803142-fa0e22cc-f377-4d66-9611-1b4978002d1c.png)
 ![image](https://user-images.githubusercontent.com/89753083/140803186-13e34970-72f3-4393-8a5c-d32768b41f02.png)
![image](https://user-images.githubusercontent.com/89753083/140803212-c28b08d3-3156-442e-b586-8c711e68f199.png)
![image](https://user-images.githubusercontent.com/89753083/140803245-dea5de17-8ce7-4f3b-ad81-ad3fab3fd42f.png)


Summary
-	Refactoring the code has an advantage of making the code more organized and easier to read.  I had some issues getting my refactored code to run correctly due to not indenting part of the code.  Once I got that figured out it ran great.  The disadvantage to refactoring the code is that it did not make the code run any faster than the original code.  
