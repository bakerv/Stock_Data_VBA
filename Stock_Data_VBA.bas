Sub stockdata():
    Dim index As Long
    Dim startingprice As Double
    Dim startingrow As Long
    Dim volume As Double
    Dim numberofsheets As Integer
    Dim compare As Double
    Dim compareticker As String
    Dim i As Long
    
    'Determine workbook size for looping
    numberofsheets = ActiveWorkbook.Worksheets.Count
    
    'Loop through all sheets of the workbook
    For j = 1 To numberofsheets
        'Activate the sheet, initialize variables
        Worksheets(j).Activate
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        index = 2
        startingrow = 2
        
        'Output headers to the sheet
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(2, 15).Value = "Greatest % increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        'Format sheet to improve presentation
        Range("J:J").ColumnWidth = 12
        Range("K:K").ColumnWidth = 13
        Range("K:K").NumberFormat = "0.00%"
        Range("O:O").ColumnWidth = 20
        Range("L:L").ColumnWidth = 16
        Range("Q2:Q3").NumberFormat = "0.00%"
        Range("A1:Q1").Font.Bold = True
        Range("O2:O4").Font.Bold = True
        
        'Loop through all rows of the sheet
        For i = 2 To lastrow
            'actions to perform before switching to a new ticker
            If Cells(i + 1, 1) <> Cells(i, 1) Then
                'Data retrieval and output
                Cells(index, 9).Value = Cells(i, 1) 'ticker name
                Cells(index, 10).Value = Cells(i, 6) - Cells(startingrow, 3) 'Yearly Change
                'Error handling for division
                If Cells(startingrow, 3) <> 0 Then
                    Cells(index, 11).Value = (Cells(i, 6) - Cells(startingrow, 3)) / Cells(startingrow, 3) 'Percent Change
                Else
                    Cells(index, 11).Value = ""
                End If
                volume = volume + Cells(i, 7).Value
                Cells(index, 12).Value = volume
                
                
                'Apply conditional formating to output columns
                If Cells(index, 10) > 0 Then
                    Cells(index, 10).Interior.ColorIndex = 4
                    Cells(index, 11).Interior.ColorIndex = 4
                Else
                    Cells(index, 10).Interior.ColorIndex = 3
                    Cells(index, 11).Interior.ColorIndex = 3
                End If
              
                'Set variables for the next ticker
                volume = 0
                startingrow = Cells(i + 1, 1).Row
                index = index + 1
                
            'actions to perform when looping through the same ticker
            Else
                'retrieval and summation of trade volume
                volume = volume + Cells(i, 7).Value
            End If
        Next i
        
        'Calculate and output maximum % increase
        For i = 2 To index - 1
            If Cells(i, 11) > compare Then
                compare = Cells(i, 11).Value
                compareticker = Cells(i, 9).Value
            End If
        Next i
        Cells(2, 17) = compare
        Cells(2, 16) = compareticker
        compare = 0
    
        'Calculate and output maximum % decrease
        For i = 2 To index - 1
            If Cells(i, 11) < compare Then
                compare = Cells(i, 11).Value
                compareticker = Cells(i, 9).Value
            End If
        Next i
        Cells(3, 17) = compare
        Cells(3, 16) = compareticker
        compare = 0
        
        'Calculate and output maximum total volume
        For i = 2 To index - 1
            If Cells(i, 12) > compare Then
                compare = Cells(i, 12).Value
                compareticker = Cells(i, 9).Value
            End If
        Next i
        Cells(4, 17) = compare
        Cells(4, 16) = compareticker
        compare = 0
    Next j
End Sub
