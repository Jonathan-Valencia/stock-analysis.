# Stock-Analysis.
 - perform data analysis using VBA

 # Overview of the Project.
  - The overview of the project was to explore the Visual Basic of Aplications aka VBA to learn the coding concepts that will help us perform more complex analyses. As we used the example in our model, I came to understand that VBA was helping me learn the fundamental building blocks of programming languages. As well as Learning to code is essentially learning how to deconstruct a problem and translate the solution into simple instructions. At the end I was able to programmatically analyze multiple stocks with skills such as syntax recollection, pattern recognition, problem decomposition, and debugging.
  ### More General Purpose and Data 
  - The purpose of this project was to refactor a Microsoft Excel VBA code to collect certain stock information in the year 2017 and 2018 and determine whether or not the stocks are worth investing. This process was originally completed in a similar format, however, the goal for this round was to increase the efficiency of the original code.
  - The data that is presented includes two charts with stock information on 12 different stocks. The stock information contains a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The goal is to retrieve the ticker, the total daily volume, and the return on each stock.
# Results
### Analysis of results 
- Before refactoring the code, I began by copying the code that was needed to create the input box, chart headers, ticker array, and to activate the appropriate worksheet. The steps were then listed out in order to set the structure for the refactoring. Below is the instruction and code as written in the file.
### Code
    Sub AllStocksAnalysisRefactored()
    'start time/end time of timer variables

    Dim startTime As Single
    Dim endTime  As Single
    
    'ask for input of year to analysis

    yearValue = InputBox("What year would you like to run the analysis on?")
    
    'Timer Starts

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

    'Create a ticker Index
    tickerIndex = 0

    'Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    ' Create a for loop to initialize the tickerVolumes to zero. If the next row’s ticker doesn’t match, increase the tickerIndex.

    For i = 0 To 11
           tickerVolumes(i) = 0
           tickerStartingPrices(i) = 0
           tickerEndingPrices(i) = 0
    Next i

    Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount

    ' Increase volume for current ticker
     tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    ' Check if the current row is the first row with the selected tickerIndex.
    
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
    
    'check if the current row is the last row with the selected ticker
  
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If

        ' Increase the tickerIndex.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If

    Next i
    
    ' Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
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
    
    'positive returns green and negative returns red
    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
    ' how fast his VBA code will compile the results/message box pop up/timer ends
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
    End Sub 
### Continued Analyze
- This visaul shows us two output run times of the code one comeing from the un-factored program and the other from the re-fractored program. As we can see from the image the re-factored code has a faster time of initialization compared to the other concluding that refracting has a benefit of a better run time then non-refactoring your code. 

### Non - refactored
![PRE-17](https://user-images.githubusercontent.com/53058061/170409526-4acbef7b-e2bf-48bf-b0f0-d5d62f45d7d8.PNG)
### Refactored
![VBA_Challenge_2017](https://user-images.githubusercontent.com/53058061/170409627-2d8ada5d-b5d1-4ee4-92ee-da1eb9328c78.PNG)


# Summary
### Advantages and Disadvantges of Refactoring Code
- A few advantages of a refactoring is cleaner code which includes design and software improvement, debugging, and faster programming. It also benefits the reability for others who need to view the project. However, we do not always have the luxury to refactor our code due to disadvantages. These disadvantages may range from having applications that are too large to not having the proper test cases for the existing codes, which may ultimately pose some risk if we try to refactor our code.



### Detailed Statment
- To give an example of an advantage if refactoring we can view the viusal shown above. This visaul shows us two output run times of the code one comeing from the non-factored program and the other from the re-fractored program. As we can see from the image the re-factored code has a faster time of initialization compared to the other concluding that refracting has a benefit of a better run time then non-refactoring your code. 
