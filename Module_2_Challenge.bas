
Sub tickerInfo():
Application.ScreenUpdating = False

'Walk through each worksheet
'Dim ws As Worksheet

'For Each ws In ThisWorkbook.Worksheets
'    ws.Activate


'Headers
'----------------------------------------------------
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change ($)"
Range("K1").Value = "Percentage Change"
Range("L1").Value = "Total Stock Volume"

Range("O1").Value = "Ticker"
Range("P1").Value = "Value"

Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"


'Determine Total Rows Count (minus 1 for heading)
'----------------------------------------------------
'<This is required for the for loops as it is not efficient to change the range everytime there are different dataset rows>

Dim rowCount As Long
Dim tickerRow As Long
Dim tickerCount As Long
Dim rows As Long
Dim percentChange As Double

tickerRow = 2
rows = 2

rowCount = Range("A1").End(xlDown).Row

'Part 1
'----------------------------------------------------
For i = 2 To rowCount + 1

    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
        'Determine the Ticker Names
        Cells(tickerRow, 9).Value = Cells(i, 1).Value

        'Calculate Yearly Change
        Cells(tickerRow, 10).Value = Cells(i, 6).Value - Cells(rows, 3).Value
        

        'Calculate percent change
        If Cells(rows, 3).Value <> 0 Then
            percentChange = ((Cells(i, 6).Value - Cells(rows, 3).Value) / Cells(rows, 3).Value)
                    
        'Percent formating
            Cells(tickerRow, 11).Value = FormatPercent(percentChange)
                    
        Else
                    
            Cells(tickerRow, 11).Value = FormatPercent(0)
                    
        End If

        'Increase tickerRow count
        tickerRow = tickerRow + 1
        
        'Increase rows count
        rows = i + 1
        
    End If
        
Next i

'Total Volume (add the values in column g if column A is equal to ticker symbol)
'----------------------------------------------------
tickerCount = Range("I1").End(xlDown).Row

    For i = 2 To tickerCount + 1
        Cells(i, 12).Value = Application.WorksheetFunction.SumIf(Range("A:A"), Cells(i, 9).Value, Range("G:G"))
    Next i

'Greatest % Increase
'----------------------------------------------------
Dim percentIncrease As Double

percentIncrease = 0

    For i = 2 To tickerCount + 1

        If percentIncrease < Cells(i, 11).Value Then

            percentIncrease = Cells(i, 11).Value
            Range("O2").Value = Cells(i, 9).Value
            Range("P2").Value = FormatPercent(percentIncrease)

        End If
    Next i

'Greatest % Decrease
'----------------------------------------------------
Dim percentDecrease As Double

percentDecrease = 0

    For i = 2 To tickerCount + 1

        If percentDecrease > Cells(i, 11).Value Then

            percentDecrease = Cells(i, 11).Value
            Range("O3").Value = Cells(i, 9).Value
            Range("P3").Value = FormatPercent(percentDecrease)

        End If
    Next i

'Greatest Total Volume
'----------------------------------------------------
Dim totalVolume As Double

totalVolume = 0

    For i = 2 To tickerCount + 1

        If totalVolume < Cells(i, 12).Value Then

            totalVolume = Cells(i, 12).Value
            Range("O4").Value = Cells(i, 9).Value
            Range("P4").Value = totalVolume

        End If
    Next i

'Apply Conditional Formatting to column J
'----------------------------------------------------
For i = 2 To tickerCount + 1
    If Cells(i, 10).Value >= 0 Then
        Cells(i, 10).Interior.Color = RGB(124, 252, 0)
    Else
        Cells(i, 10).Interior.Color = RGB(255, 0, 0)
    End If
Next i

'Apply Conditional Formatting to column K
'----------------------------------------------------
For i = 2 To tickerCount + 1
    If Cells(i, 11).Value >= 0 Then
        Cells(i, 11).Interior.Color = RGB(124, 252, 0)
    Else
        Cells(i, 11).Interior.Color = RGB(255, 0, 0)
    End If
Next i

'Next ws
Application.ScreenUpdating = True

End Sub
