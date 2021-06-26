## DQ-Analysis

#### **Project Overview**
- The project was to refacter already existing code, into something that is more efficient to run. The code is to determine the best stocks to invest in, taking into consideration the opeing and closing prices of the stocks. 
#### **Results**
- **Code**
- - `Sub AllStocksAnalysisRefactored()
    
    Dim startTime As Single
    
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    
    Worksheets("AllStockAnalysis").Activate
    
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
    
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
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
         
         If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 3).Value

         End If
        
        
        '3c) check if the current row is the last row with the selected ticker
        
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            tickerIndex = tickerIndex + 1
           
        End If
            
            '3d Increase the tickerIndex.
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        Worksheets("AllStockAnalysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    
    Worksheets("AllStockAnalysis").Activate
    
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
        
    endTime = Timer
    
    Next i
    
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
        End Sub
`

- **Screenshots**
- - ![VBA_Challenge_2017](https://user-images.githubusercontent.com/85029175/123499192-88e23b00-d5fa-11eb-8a73-f55a0e0dc566.png) ![VBA_Challenge_2018](https://user-images.githubusercontent.com/85029175/123499196-94cdfd00-d5fa-11eb-80fb-48a832732643.png)

#### **Summary**
- **Advantages of Refactoring Code**
- - One of the advantages of refactering code is that the code is already there, and you just have to make changes to it. If the coder left sufficient comments, then the code will be easy enough to follow along and understand what is happening. This could really help speed up the process of coding, especially if working in teams and big groups. I would also help in the review process, if you send your code to someone to look at and they might refactor it to be more efficient (or correct).

- **Disadvantages of Refactoring Code**
- - While already having code there might be nice at times, sometimes the code might be so bad that it would be best to start from scratch. If someone keeps refactoring the code, they might end up getting lost as to what the initial result should be and how they are getting to where they are at. Also to many changes might make it difficult for someone else to come in and understand what is happening. Also if there aren't any comments then it might be hard to understand the code.

- **Adantage of the Original VBA Script**
- - The original code had great comments and it was easy to understand what the original coder was trying to do. Those comments made it easy enough to follow along and clean up the code a bit. 

- **Disadvantage of the Original VBA Script**
- - While the comments were great, the code itself was lacking a bit. There were certainly big chunks that were missing, and that made a bit hard to figure out what code was needed in order to make it run more efficient. 
