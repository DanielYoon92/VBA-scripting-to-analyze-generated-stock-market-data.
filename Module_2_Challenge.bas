Attribute VB_Name = "Module1"

Sub tickerInfo():
Application.ScreenUpdating = False

'Walk through each worksheet
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    ws.Activate


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

rowCount = Range("A1").End(xlDown).Row

'Find the Ticker Names
'----------------------------------------------------
'Create an arraylist object to store the different ticker symbols
'Source: https://analystcave.com/vba-arraylist-using-vba-arraylist-excel/)

Dim tickerCount As Integer

Set tickerArray = CreateObject("System.Collections.ArrayList")

    For i = 2 To rowCount + 1
    
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
            tickerArray.Add Cells(i, 1).Value

        End If
    Next i

tickerCount = tickerArray.Count


'Inputting the data in the worksheet
'----------------------------------------------------

'Ticker Symbol
'-------------
    For i = 2 To tickerCount + 1
        Cells(i, 9).Value = tickerArray(i - 2) 'minus 2 since array starts from 0
    Next i


'Yearly Change
'--------------
'Create an array object to store the open and closed prices for each tickers

Set openPrice = CreateObject("System.Collections.ArrayList")

     For i = 2 To rowCount + 1
        For J = 2 To tickerCount + 1
        
            If Cells(i, 2).Value = "20200102" And Cells(i, 1).Value = tickerArray(J - 2) Then
                openPrice.Add Cells(i, 3).Value
    
            End If
        Next J
    Next i

Set closePrice = CreateObject("System.Collections.ArrayList")

     For i = 2 To rowCount + 1
        For J = 2 To tickerCount + 1
        
            If Cells(i, 2).Value = "20201231" And Cells(i, 1).Value = tickerArray(J - 2) Then
                closePrice.Add Cells(i, 6).Value
    
            End If
        Next J
     Next i


For i = 2 To (openPrice.Count + 1)
    Cells(i, 10).Value = (closePrice(i - 2) - openPrice(i - 2))
Next i


'Percentage Change
For i = 2 To (openPrice.Count + 1)
    Cells(i, 11).Value = FormatPercent((closePrice(i - 2) - openPrice(i - 2)) / openPrice(i - 2))
Next i


'Total Volume (add the values in column g if column A is equal to ticker symbol)
    For i = 2 To tickerCount + 1
        Cells(i, 12).Value = Application.WorksheetFunction.SumIf(Range("A:A"), tickerArray(i - 2), Range("G:G"))
    Next i

'Greatest % Increase
Dim percentIncrease As Double

percentIncrease = 0

    For i = 2 To tickerCount + 1
    
        If percentIncrease < Cells(i, 11).Value Then
        
            percentIncrease = Cells(i, 11).Value
            Range("O2").Value = Cells(i, 9).Value
            Range("P2").Value = FormatPercent(percentIncrease)
        
        End If
    Next i
    
'Greatest & Decrease
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
For i = 2 To tickerCount + 1
    If Cells(i, 10).Value >= 0 Then
        Cells(i, 10).Interior.Color = RGB(124, 252, 0)
    Else
        Cells(i, 10).Interior.Color = RGB(255, 0, 0)
    End If
Next i

'Apply Conditional Formatting to column K
For i = 2 To tickerCount + 1
    If Cells(i, 11).Value >= 0 Then
        Cells(i, 11).Interior.Color = RGB(124, 252, 0)
    Else
        Cells(i, 11).Interior.Color = RGB(255, 0, 0)
    End If
Next i
    
Next ws
Application.ScreenUpdating = True

End Sub
