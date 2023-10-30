VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockAnalysis():

    ' Loop through all worksheets
    For Each ws In Worksheets
    
        ' Identify the last row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Assign variables
        Dim Ticker As String
        Dim YearlyChange As Double
        Dim PercentChange As Double
        
        'Start Total Volume variable at 0
        Dim TotalVolume As Double
        TotalVolume = 0

        ' This variable indicates the row where the opening value is calculated for volume change variable
        Dim FirstInGroup As Long
        
        ' Starts FirstInGroup at row 2 (for calculation of yearly change)
        FirstInGroup = 2

        ' Indicate beginning of SummaryTableRow
        Dim SummaryTableRow As Integer
        SummaryTableRow = 2
        
        ' Place headers at the top of the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' AutoFit Columns
        ws.Range("I1:L1").Columns.AutoFit
        
    For i = 2 To lastrow
        
            ' If the value of Column A (ticker) changes perform the following:
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                ' Assign ticker variable
                Ticker = ws.Cells(i, 1).Value
            
                ' Print ticker symbol in designated row in summary table
                ws.Range("I" & SummaryTableRow).Value = Ticker
                
                ' Calculate YearlyChange as the opening value in the first cell in a stock group minus the closing value for the last cell in that group
                YearlyChange = ws.Cells(i, 6).Value - ws.Cells(FirstInGroup, 3).Value
                
                ' Print YearlyChange in designated row in summary table
                ws.Cells(SummaryTableRow, 10).Value = YearlyChange
        
                ' Use conditional formatting to color YearlyChange column as green if its positive and red if its negative or 0
                If YearlyChange > 0 Then
                    ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
                    
                Else
                    ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
                
                End If
                
                ' Calculate PercentChange as YearlyChange over the opening value on the first day of the year
                PercentChange = YearlyChange / ws.Cells(FirstInGroup, 3).Value
                
                 ' Round PercentChange to two decimal places
                PercentChange = WorksheetFunction.Round(PercentChange, 4)
                
                ' Print Percent change in designated row in summary table
                ws.Cells(SummaryTableRow, 11).Value = "%" & (PercentChange * 100)
                
                ' Add value of current cell to TotalVolume variable
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
                ' Print TotalVolume variable in summary table
                ws.Cells(SummaryTableRow, 12) = TotalVolume
                
                ' Move down one row in Summary Table
                SummaryTableRow = SummaryTableRow + 1
                
                ' Set the FirstinGroup variable to be the first row of the ticker change
                FirstInGroup = i + 1
                
                ' Reset TotalVolume to 0 at the last in a stock group
                TotalVolume = 0
                
            ' For all rows where the ticker remains the same perform the following:
            Else
            
                ' Add the current volume to TotalVolume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
            End If
        
        Next i
        
    '-------------------------------------------------------
    ' Make Additional Summary Table
    '-------------------------------------------------------
    
        ' Place headers at the top of the additional summary table
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
    
        ' Place to the left of the second summary table
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
    
        ' AutoFit Column
        ws.Range("O2:O4").Columns.AutoFit
        
        ' Find the max value in the Percentage Change column in the summary table and print in additional summary table
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastrow)) * 100
        
        ' Find the associated ticker value for Greatest PercentChange
        TickerMaxPercent = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
        
        ' Print the ticker symbol for the greatest percentage change in the additional summary table
        ws.Range("P2").Value = ws.Cells((TickerMaxPercent + 1), 9)
    
        'Find the minimum value in the Percentage Change column in the summary table and print in additional summary table
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastrow)) * 100
        
         ' Find the associated ticker value for Smallest PercentChange
        TickerMinPercent = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
        
        ' Print the ticker symbol for the greatest percentage change in the additional summary table
        ws.Range("P3").Value = ws.Cells((TickerMinPercent + 1), 9)
    
        ' Find the maximum total volume in the summary table
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
        
        ' Find the associated ticker value for greatest total volume
        TickerVolume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)
        
        ' Print ticker symbol for the greatest percentage change in the additional summary table
        ws.Range("P4").Value = ws.Cells((TickerVolume + 1), 9)
        
    Next ws
    
End Sub


