Attribute VB_Name = "Module1"
Sub wall_street()

    For Each ws In Worksheets

        ' Set a variable for holding the ticker name
        Dim ticker As String
        
        ' Set a variable for counting the number of rows in the workshet
        Dim lastRow As Long
        
        ' Set variables for storing the opening and closing prices for the year
        Dim openPrice, closePrice, yearChange, percentChange As Double
        
        ' Set a variable for the running tally of the stock volume
        ' Must be double as otherwise variable will overflow!
        Dim stockVolume As Double
        
        ' Set variables for maximum and minimum percent change and maximum total volume
        Dim maxPercentChange, minPercentChange, maxTotalVolume As Double
        
        ' Set variables for row numbers of maximum and minimum percent change and maximum total volume
        Dim maxRowNumber, minRowNumber, maxTVRowNumber As Long
        
        ' Set variabls for ticker values of maximum and minimum percent change and maximum total volume
        Dim maxTicker, minTicker, maxTVTicker As String
        
        ' Set initial value of stockVolume to 0
        stockVolume = 0
        
        ' Set a variable for the index of the current row in the results table
        Dim resultsTableRow As Long
        
        ' Set initial value of resultsTableRow to 2 (first row is headers)
        resultsTableRow = 2

        ' Count the number of rows to loop through
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Set up results tables
        
        ' Clear any previous results
        ws.Range("I:Q").ClearContents
        
        ' Clear any previous formatting
        ws.Range("I:Q").ClearFormats
        
        ' Reset fill colour of yearly change ws.Cells to blank
        ws.Range("J:J").Interior.Color = xlNone
        
        ' Fill column headers of results table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ' Label rows of maximum values table
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ' Label ws.Columns of maximum values table
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ' Autofit to display data
        ws.Columns("I:Q").AutoFit
        
        ' Set initial value of ticker
        ticker = ws.Cells(2, 1).Value
        
        ' Set initial value of stock opening price
        openPrice = ws.Cells(2, 3).Value

        ' Set initial value of stock volume
        stockVolume = 0
        
        ' Loop through all rows containing stock data
        For i = 2 To lastRow
        
            ' Check if we are still with the same ticker
            ' New ticker has been found
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' Set ticker name for results
                ticker = ws.Cells(i, 1).Value
                
                ' Set closing price of stock for year
                closePrice = ws.Cells(i, 6).Value
                
                ' Calculate the yearly change in the stock price
                yearChange = closePrice - openPrice
                
                
                ' Calculate the percentage change in the stock price
                ' Need to check for bad data i.e. 0 in the opening price
                ' Otherwise percent change division causes overflow error
                If (openPrice = 0) Then
                    percentChange = 0
                Else
                     percentChange = yearChange / openPrice
                End If
               
                ' Calculate the last value for the stock volume for the current ticker
                stockVolume = stockVolume + ws.Cells(i, 7).Value

                ' Fill in the results table with the correct values
                ws.Range("I" & resultsTableRow).Value = ticker
                ws.Range("J" & resultsTableRow).Value = yearChange
                ws.Range("K" & resultsTableRow).Value = percentChange
                ws.Range("L" & resultsTableRow).Value = stockVolume
                
                ' Format yearly change values
                ' Set cell colour to red if negative change
                If (percentChange < 0) Then
                    ws.Range("J" & resultsTableRow).Interior.ColorIndex = 3
                ' Otherwise for positive or zero change set cell to green
                Else
                    ws.Range("J" & resultsTableRow).Interior.ColorIndex = 4
                End If
                
                ' Increment the results table to the next row index
                resultsTableRow = resultsTableRow + 1
                
                ' Set stock opening price for new ticker
                openPrice = ws.Cells(i + 1, 3).Value
               
                ' Reset stock volume for new ticker
                stockVolume = 0
                
              ' Otherwise still calculating for same stock ticker
              Else
                
                ' Add to the running tally of the value for stock volume
                stockVolume = stockVolume + ws.Cells(i, 7).Value
                
              End If
        
        Next i
        
        ' Format Percent Change column with percent symbol and two decimal places
        ws.Range("K:K").NumberFormat = "0.00%"
        
        ' Find greatest percent increase
        maxPercentChange = WorksheetFunction.Max(ws.Range("K:K"))
        
        ' Find row number in results table of greatest percent increase
        maxRowNumber = WorksheetFunction.Match(maxPercentChange, ws.Range("K:K"), 0)

        ' Find ticker of greatest percent increase
        maxTicker = ws.Cells(maxRowNumber, 9).Value
        
        ' Set values of ticker and greatest percent increase in results table
        ws.Range("P2").Value = maxTicker
        ws.Range("Q2").Value = maxPercentChange

        ' Find greatest percent decrease
        minPercentChange = WorksheetFunction.Min(ws.Range("K:K"))
        
        ' Find row number in results table of greatest percent decrease
        minRowNumber = WorksheetFunction.Match(minPercentChange, ws.Range("K:K"), 0)

        ' Find ticker of greatest percent decrease
        minTicker = ws.Cells(minRowNumber, 9).Value
        
        ' Set values of ticker and greatest percent decrease in results table
        ws.Range("P3").Value = minTicker
        ws.Range("Q3").Value = minPercentChange

        ' Format Greatest Percent Increase and Decrease with percent symbol and two decimal places
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        ' Find greatest total stock volume
        maxTotalVolume = WorksheetFunction.Max(ws.Range("L:L"))
        
        ' Find row number in results table of greatest total volume
        maxTVRowNumber = WorksheetFunction.Match(maxTotalVolume, ws.Range("L:L"), 0)
        
        ' Find ticker of greatest total volume
        maxTVTicker = ws.Cells(maxTVRowNumber, 9).Value
        
        ' Set values of ticker and greatest total volume in results table
        ws.Range("P4").Value = maxTVTicker
        ws.Range("Q4").Value = maxTotalVolume
        
        ' Format greatest total volume into scientific notation with four decimal places
        ws.Range("Q4").NumberFormat = "0.0000E+00"
        
        ' Autofit results tables to display data properly
        ws.Columns("I:L").AutoFit
        ws.Columns("O:Q").AutoFit

    Next ws

End Sub

