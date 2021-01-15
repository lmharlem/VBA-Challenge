'RUN THIS SCRIPT AFTER RUNNING THE STOCK ANALYSIS_1 SCRIPT

'This script will return the stock information with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".

Sub Stock_Analysis_2()

'Declare variables
    Dim GreatestDecrease As Double
    Dim GreatestIncrease As Double
    Dim LastRowValue As Long
    Dim GreatestTotalVolume As Double
        
'Loop Through All Worksheets
    For Each ws In Worksheets
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
'Initialize Variables
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestTotalVolume = 0
     
'Compute Final Results
        LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
            
        For i = 2 To LastRow
            If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
               ws.Range("Q2").Value = ws.Range("K" & i).Value
               ws.Range("P2").Value = ws.Range("I" & i).Value
             End If

                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

         Next i
        
         ws.Range("Q2").NumberFormat = "0.00%"
         ws.Range("Q3").NumberFormat = "0.00%"
         ws.Columns("I:Q").AutoFit

    Next ws

End Sub
