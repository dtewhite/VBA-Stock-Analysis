Attribute VB_Name = "Module1"
Sub stock()
    
    ' Make header values
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    ' Define a variable so that each successive calculation prints on the next row
    Dim printrow As Long
    printrow = 2
        
    ' Define a variable to find the last row of the data set
    Dim lastrow As String
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Define a variable that will save the first row of stock data
    Dim stockopenrow As String
    stockopenrow = 2
    
    ' Build a loop that will search each row of the stock data
    For i = 2 To lastrow
        ' If the stock ticker changes in the next row, print the ticker in the print row
        If Cells(i, 1) <> Cells(i + 1, 1) Then
            Cells(printrow, 9).Value = Cells(i, 1).Value
            ' Print the closing amount on the last line less the opening amount on the first line w/ formatting
            Cells(printrow, 10).Value = Cells(i, 6) - Cells(stockopenrow, 3)
            If Cells(printrow, 10).Value < 0 Then
                Cells(printrow, 10).Interior.ColorIndex = 3
            ElseIf Cells(printrow, 10).Value > 0 Then
                Cells(printrow, 10).Interior.ColorIndex = 4
            ElseIf Cells(printrow, 10).Value = 0 Then
                Cells(printrow, 10).Interior.ColorIndex = 6
            End If
            ' Print the % change in stock price
            If Cells(stockopenrow, 3).Value <> 0 Then
                Cells(printrow, 11).Value = (Cells(i, 6) - Cells(stockopenrow, 3)) / Cells(stockopenrow, 3)
                Cells(printrow, 11).NumberFormat = "0.00%"
            ' I was getting an overflow error because there was a ticker with 0 as all prices
            Else
                Cells(printrow, 11).Value = 0
                Cells(printrow, 11).Interior.ColorIndex = 6
                Cells(printrow, 10).Value = 0
            End If
            ' Sum of trading volume
            Cells(printrow, 12).Value = Application.Sum(Range(Cells(stockopenrow, 7), Cells(i, 7)))
            Cells(printrow, 12).NumberFormat = "#,###"
            ' Mark the start of the next ticker
            stockopenrow = i + 1
            ' Move the printrow one row down
            printrow = printrow + 1
        End If
    Next i
    
    ' Bonus
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest total volume"
    Cells(2, 16).Value = GreatestIncrease
    Cells(3, 16).Value = GreatestDecrease
    
    Dim lastrowprinted As String
    lastrowprinted = Cells(Rows.Count, 9).End(xlUp).Row
        
    Cells(2, 16).Value = Application.WorksheetFunction.Max(Range(Cells(2, 11), Cells(lastrowprinted, 11)))
    Cells(2, 15).Value = WorksheetFunction.Index(Range("I:L"), WorksheetFunction.Match(Cells(2, 16), Range("K:K"), 0), 1)
    Cells(2, 16).NumberFormat = "0.00%"
    Cells(3, 16).Value = Application.WorksheetFunction.min(Range(Cells(2, 11), Cells(lastrowprinted, 11)))
    Cells(3, 15).Value = WorksheetFunction.Index(Range("I:L"), WorksheetFunction.Match(Cells(3, 16), Range("K:K"), 0), 1)
    Cells(3, 16).NumberFormat = "0.00%"
    Cells(4, 16).Value = Application.WorksheetFunction.Max(Range(Cells(2, 12), Cells(lastrowprinted, 12)))
    Cells(4, 15).Value = WorksheetFunction.Index(Range("I:L"), WorksheetFunction.Match(Cells(4, 16), Range("L:L"), 0), 1)
    Cells(4, 16).NumberFormat = "#,###"
    
    Columns("I:P").Select
    Selection.EntireColumn.AutoFit
    
End Sub

