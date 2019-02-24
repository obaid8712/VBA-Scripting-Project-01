Attribute VB_Name = "Module1"
Sub MainProg():
    'Declare variables
    Dim i As Integer
    Dim WsCount As Integer
    
    'COUNT NUMBER OF WORKSHEET
    WsCount = Sheets.Count
    'MsgBox (WsCount)
    'CALL EACH SHEET AND CALCULATE DATA
    
    For i = 1 To WsCount
        StockCalculator1 (i)
    Next i
    'StockCalculator1 (2)
End Sub

Sub StockCalculator1(s As Integer):
    'Declare Variables
    Dim Main As Workbook
    Dim LastRow As Long
    Dim i As Double
    Dim j As Integer
    Dim k As Integer
    Dim YearOpenVal As Double
    Dim YearCloseVal As Double
    Dim SumVol As Double
        
    Set Main = ActiveWorkbook
    'Find the last row number of worksheet
    LastRow = Main.Sheets(s).Cells(Rows.Count, 2).End(xlUp).Row
            
    'Lebel the Header
    Worksheets(s).Range("i1").Value = "Ticker"
    Worksheets(s).Range("j1").Value = "Yearly Change"
    Worksheets(s).Range("k1").Value = "Percent Change"
    Worksheets(s).Range("l1").Value = "Total Stock Volume"
    
    'Find the distinct value for Ticker and calculate yearly change and Percent change
    j = 2
    k = 2
    SumVol = 0
    'LastRow = 2000
    For i = 2 To LastRow
        'Store yearly Opening value
        If k = j Then
            YearOpenVal = Worksheets(s).Cells(i, 3).Value
            k = k + 1 'pointer increment
        End If
               
        'Check the Imediate ticker value change
        If Worksheets(s).Cells(i, 1).Value <> Worksheets(s).Cells(i + 1, 1).Value Then
            'Store yearly Closing value
            YearCloseVal = Worksheets(s).Cells(i, 6).Value
            'Output for ticker
            Worksheets(s).Cells(j, 9).Value = Worksheets(s).Cells(i, 1).Value
            'Calculate Yearly changes
            Worksheets(s).Cells(j, 10).Value = YearCloseVal - YearOpenVal
            'Conditional Formatting of Yearly change
            If Worksheets(s).Cells(j, 10).Value > 0 Then
                Worksheets(s).Cells(j, 10).Interior.ColorIndex = 4 'Green
            Else
                Worksheets(s).Cells(j, 10).Interior.ColorIndex = 3  'Red
            End If
            'Check Null value and Calculate Percent Change
            If YearOpenVal <> 0 Then
                Worksheets(s).Cells(j, 11).Value = (YearCloseVal - YearOpenVal) / YearOpenVal
            Else
                Worksheets(s).Cells(j, 11).Value = 0
            End If
            Worksheets(s).Cells(j, 11).NumberFormat = "0.00%"
            'Accumulate Stock volume
            Worksheets(s).Cells(j, 12).Value = SumVol + Worksheets(s).Cells(i, 7).Value
            SumVol = 0
            j = j + 1 'pointer increment
        Else
            'Accumulate Stock volume
            SumVol = SumVol + Worksheets(s).Cells(i, 7).Value
        End If
    Next i
    'CALCULATE GREATEST VALUE % INCREASE, DECREASE AND TOTAL VOL
    Dim TickerInc As String
    Dim TickerDec As String
    Dim TickerTotal As String
    Dim LastRowPercent As Integer
    Dim MaxIncPercent As Double
    Dim MaxDecPercent As Double
    Dim MaxTotalValue As Double
    
    'Lebel Header
    Main.Sheets(s).Cells(1, 15).Value = "Ticker"
    Main.Sheets(s).Cells(1, 16).Value = "Value"
    Main.Sheets(s).Cells(2, 14).Value = "Greatest % increase"
    Main.Sheets(s).Cells(3, 14).Value = "Greatest % Decrease"
    Main.Sheets(s).Cells(4, 14).Value = "Greatest total volume"
    'Find the last row number of worksheet
    LastRowPercent = Main.Sheets(s).Cells(Rows.Count, 11).End(xlUp).Row
    'MsgBox (LastRowPercent)
    'Assinment for initial value
    i = 0
    MaxIncPercent = Worksheets(s).Cells(2, 11).Value
    MaxDecPercent = Worksheets(s).Cells(2, 11).Value
    MaxTotalValue = Main.Sheets(s).Cells(2, 12).Value
    TickerInc = Main.Sheets(s).Cells(2, 9).Value
    TickerDec = Main.Sheets(s).Cells(2, 9).Value
    TickerTotal = Main.Sheets(s).Cells(2, 9).Value
    For i = 3 To LastRowPercent
        'FIND MAXPERCENT INCREASE
        If MaxIncPercent < Main.Sheets(s).Cells(i, 11).Value Then
            MaxIncPercent = Worksheets(s).Cells(i, 11).Value
            TickerInc = Main.Sheets(s).Cells(i, 9).Value
        End If
        'FIND MAXPERCENT DECREASE
        If MaxDecPercent > Main.Sheets(s).Cells(i, 11).Value Then
            MaxDecPercent = Worksheets(s).Cells(i, 11).Value
            TickerDec = Main.Sheets(s).Cells(i, 9).Value
        End If
        
        'FIND MAX TOTAL VOL
        If MaxTotalValue < Main.Sheets(s).Cells(i, 12).Value Then
            MaxTotalValue = Worksheets(s).Cells(i, 12).Value
            TickerTotal = Main.Sheets(s).Cells(i, 9).Value
        End If
    Next i
    'OUTPUT FOR MAX PERCENT INCREASE
    Main.Sheets(s).Cells(2, 16).Value = MaxIncPercent
    Main.Sheets(s).Cells(2, 16).NumberFormat = "0.00%"
    Main.Sheets(s).Cells(2, 15).Value = TickerInc
    
    'OUTPUT FOR MAX PERCENT DECREASE
    Main.Sheets(s).Cells(3, 16).Value = MaxDecPercent
    Main.Sheets(s).Cells(3, 16).NumberFormat = "0.00%"
    Main.Sheets(s).Cells(3, 15).Value = TickerDec
    
    'OUTPUT FOR MAX TOTAL VOLUME
    Main.Sheets(s).Cells(4, 16).Value = MaxTotalValue
    'Main.Sheets(s).Cells(4, 16).NumberFormat = "0.00%"
    Main.Sheets(s).Cells(4, 15).Value = TickerTotal
End Sub

