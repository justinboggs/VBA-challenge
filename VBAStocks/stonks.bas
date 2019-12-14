Attribute VB_Name = "stonks"
Sub stonks()

    'run script on each worksheet
    For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
    'create headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
'determine tickers and total vol
    
    'set variable for holding ticker
    Dim tickSym As Integer
    tickSym = 1
    
    'set variable for total vol
    Dim totalVol As Double
    totalVol = 0
    
    'set the summary table
    Dim summTable1 As Long
    summTable1 = 2
    
    'set variable to stop loop at last populated row
    stopRow1 = Cells(Rows.Count, 1).End(xlUp).Row
    
    'loop through tickers until blank cells
    For i = 2 To stopRow1
    
        'check if cells are within the same ticker
        If Cells(i + 1, tickSym).Value <> Cells(i, tickSym).Value Then
            'add the last vol total to the variable
            totalVol = totalVol + Cells(i, 7).Value
        
            'write the vol summary table
            Cells(summTable1, 9).Value = Cells(i, tickSym).Value
            Cells(summTable1, 12).Value = totalVol
        
            'reset total vol
            totalVol = 0
        
            'next table line
            summTable1 = summTable1 + 1
        Else
            'still on same ticker, add to the total vol
            totalVol = totalVol + Cells(i, 7).Value
        End If
        
    
    Next i
    
'determine yearly and percent change

    'set variable to find the first and last stock prices per ticker
    Dim firstPrice As Long
    Dim lastPrice As Long
    
    'set the summary table
    Dim summTable2 As Long
    summTable2 = 2
    
    'set variable to stop loop at last populated row
    stopRow2 = Cells(Rows.Count, 9).End(xlUp).Row
    
    'loop through tickers until blank cells
    For i = 2 To stopRow2
        
        'set the first and last prices per ticker
        firstPrice = Range("A:A").Find(what:=Cells(i, 9), after:=Cells(1, 1), LookAt:=xlWhole).Row
        lastPrice = Range("A:A").Find(what:=Cells(i, 9), after:=Cells(1, 1), LookAt:=xlWhole, SearchDirection:=xlPrevious).Row
            
        'Write the price change and percent change to the summary table
        Range("J" & summTable2).Value = Range("F" & lastPrice).Value - Range("C" & firstPrice).Value
        
        'divide by 0 error checker
        If Range("C" & firstPrice).Value = 0 Then
            Range("C" & firstPrice).Value = Null
        Else
            Range("K" & summTable2).Value = (Range("F" & lastPrice).Value - Range("C" & firstPrice).Value) / Range("C" & firstPrice).Value
        End If
        
        'next table line
        summTable2 = summTable2 + 1
        
    Next i
               
'determine greatest total vol
    
    'Set variable and range for total vol
    Dim volRng As Range
    Set volRng = Range("L:L")
    
    'Set variable and max vol
    Dim volMax As Double
    volMax = WorksheetFunction.Max(volRng)
    
    'set variable to stop loop at last populated row
    stopRow3 = Cells(Rows.Count, 12).End(xlUp).Row
    
    For i = 2 To stopRow3
        If Cells(i, 12).Value = volMax Then
            Cells(4, 16).Value = Cells(i, 9).Value
            Cells(4, 17).Value = Cells(i, 12).Value
        End If
    Next i
    
'determine greatest percentage decrease
    
    'set variable and range for max decrease
    Dim decRng As Range
    Set decRng = Range("K:K")
    
    'set variable for max decrease
    Dim decMax As Double
    decMax = WorksheetFunction.Min(decRng)
    
    'set variable to stop loop at last populated row
    stoprow4 = Cells(Rows.Count, 11).End(xlUp).Row
    
    For i = 2 To stoprow4
        If Cells(i, 11).Value = decMax Then
            Cells(3, 16).Value = Cells(i, 9).Value
            Cells(3, 17).Value = Cells(i, 11).Value
        End If
    Next i
    
'determine greatest percentage increase
    
    'set variable and range for max increase
    Dim incRng As Range
    Set incRng = Range("K:K")
    
    'set variable for max increase
    Dim incMax As Double
    incMax = WorksheetFunction.Max(incRng)
   
    For i = 2 To stoprow4
        If Cells(i, 11).Value = incMax Then
            Cells(2, 16).Value = Cells(i, 9).Value
            Cells(2, 17).Value = Cells(i, 11).Value
        End If
    Next i
  
    'autofit all columns
    Columns("A:Z").AutoFit
    
    'conditional color formatting for column J
    'setting variables for range and condition
    Dim colorRng As Range
    Dim condition1 As FormatCondition, condition2 As FormatCondition
    Set colorRng = Range("$J$2:$J$10000")
    
    'greater than 0
    Set condition1 = colorRng.FormatConditions.Add(xlCellValue, xlGreater, "=0")
    'less than 0
    Set condition2 = colorRng.FormatConditions.Add(xlCellValue, xlLess, "=0")
    
    With condition1
        .Interior.Color = vbGreen
    End With
    
    With condition2
        .Interior.Color = vbRed
    End With
    
    'formatting numbers
    With Range("L:L")
        .NumberFormat = "#,##0"
    End With
    
    With Range("Q4")
        .NumberFormat = "#,##0"
    End With
    
    With Range("K:K")
        .NumberFormat = "0.00%"
    End With
    
    With Range("Q2:Q3")
        .NumberFormat = "0.00%"
    End With

    'closing the command to run on all sheets
    Next ws

End Sub

