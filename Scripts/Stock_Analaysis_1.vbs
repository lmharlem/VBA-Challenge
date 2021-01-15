'RUN THIS SCRIPT FIRST

'This script will loop through all the stocks for one year with conditional formatting that highlights
'positive change in green and negative change in red and output the following information
  '1)The ticker symbol.
  '2)Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
  '3)The percent change from opening price at the beginning of a given year to the closing price at the end of that year. 4)The total stock volume of the stock.

Sub Stock_Analysis_1()

'Declare variables
    Dim TickName As String
    Dim LastRow As Long
    Dim TotTickVol As Double
    Dim SumTable As Long
    Dim YearlyOpen As Double
    Dim YearlyClose As Double
    Dim YearlyChange As Double
    Dim PreviousAmount As Long
    Dim PercentChange As Double
    Dim LastRowValue As Long
        
'Loop Through All Worksheets
    For Each ws In Worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
     
'Initialize Variables
    TotTickVol = 0
    SumTable = 2
    PreviousAmount = 2
            
'Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
    For i = 2 To LastRow

' Add To Ticker Total Volume
      TotTickVol = TotTickVol + ws.Cells(i, 7).Value
      
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                TickName = ws.Cells(i, 1).Value
                ws.Range("I" & SumTable).Value = TickName
                ws.Range("L" & SumTable).Value = TotTickVol
                TotTickVol = 0
        
        
' Set Yearly Open, Yearly Close and Yearly Change Name
                YearlyOpen = ws.Range("C" & PreviousAmount)
                YearlyClose = ws.Range("F" & i)
                YearlyChange = YearlyClose - YearlyOpen
                ws.Range("J" & SumTable).Value = YearlyChange

' Determine Percent Change
                If YearlyOpen = 0 Then
                    PercentChange = 0
                Else
                    YearlyOpen = ws.Range("C" & PreviousAmount)
                    PercentChange = YearlyChange / YearlyOpen
                End If
' Format Double To Include % Symbol And Two Decimal Places
                ws.Range("K" & SumTable).NumberFormat = "0.00%"
                ws.Range("K" & SumTable).Value = PercentChange

' Conditional Formatting Highlight Positive (Green) / Negative (Red)
                If ws.Range("J" & SumTable).Value >= 0 Then
                    ws.Range("J" & SumTable).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & SumTable).Interior.ColorIndex = 3
                End If
            
' Add One To The Summary Table Row
                SumTable = SumTable + 1
                PreviousAmount = i + 1
                End If
            Next i
     Next ws
 
End Sub
