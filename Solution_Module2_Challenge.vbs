' Solution for Module 2 Challenge  
'  By: Itzel Vázquez Sánchez
' --------------------------------------------------------------------

Sub QuarterlyTicker()

    'Create variables for Loops

    Dim ws As Worksheet
    
    'Loop through all sheets

    For Each ws in Worksheets
        Dim i As Long
        Dim LastRow As Long

        Dim OpenPrice As Double
        Dim ClosePrice As Double

        ' Initial variable for holding stock name
        Dim TickerName As String

        'Initial variable for holding total volume 
        Dim StockVolume As LongLong
        StockVolume = 0

        'Keep track of location for each stock name in summary table
        Dim Stock_Summary_Table As Integer
        Stock_Summary_Table = 2

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Counter for row number
        Dim RowCount As Long
        RowCount = 0

        ' Loop through all ticker names
        For i = 2 To LastRow

            ' Check if we are still within the same credit card brand, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Write Ticker Name
                TickerName = ws.Cells(i, 1).Value

                'Total Stock Volume
                StockVolume = StockVolume + ws.Cells(i, 7).Value                
                
                ' Print in Summary table                
                ws.Range("I" & Stock_Summary_Table).Value = TickerName
                ws.Range("L" & Stock_Summary_Table).Value = StockVolume                
                
                'Calculate quarterly change
                ClosePrice = ws.Cells(i, 6).Value
                OpenPrice = ws.Cells(i - RowCount, 3).Value 

                'Print in Summary table
                ws.Range("J" & Stock_Summary_Table).Value = ClosePrice - OpenPrice 
                ws.Range("K" & Stock_Summary_Table).Value =  (Round((((ClosePrice - OpenPrice)/OpenPrice) * 100) , 2) & "%")

                ' Reset the Stock Volume counter & add one to summary table row
                StockVolume = 0
                Stock_Summary_Table = Stock_Summary_Table + 1
                RowCount = 0
                
            Else

                'If the Ticker is exactly the same, add volume to total volume
                StockVolume = StockVolume + ws.Cells(i, 7).Value 
                RowCount = RowCount + 1

            End If

        Next i

        ' Create variables for change in prices
        Dim QuarterlyChange As Double
        Dim PercentChange As Double


        ' Create four empty columns
        ws.Range("I1").Value = "Ticker" 
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ' Return "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
        Dim MaxIncrease As Double
        Dim MinValue As Double
        Dim GreatestVolume aS LongLong

        MaxIncrease = application.worksheetfunction.max(ws.Range("K2:K" & LastRow))
        MinValue = application.worksheetfunction.min(ws.Range("K2:K" & LastRow))
        GreatestVolume = application.worksheetfunction.max(ws.Range("L2:L" & LastRow))

        ws.Range("O2").Value = "Createst % Increase" 
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("Q2").Value = Round(MaxIncrease * 100, 2) & "%"
        ws.Range("Q3").Value = Round(MinValue * 100, 2) & "%"
        ws.Range("Q4").Value = GreatestVolume
        
        For i = 2 To LastRow
            
            If ws.Cells(i, 11).Value = MaxIncrease Then
               
                ws.Range("P2").Value = ws.Cells(i, 9)

            Elseif ws.Cells(i, 11).Value = MinValue Then
                
            ws.Range("P3").Value = ws.Cells(i, 9)


            End If
                    
            If ws.Cells(i, 12).Value = GreatestVolume Then
                
                ws.Range("P4").Value = ws.Cells(i, 9)
            
            End If


        Next i
        
        ' Conditional formatting    
        For i = 2 To LastRow

            If ws.Cells(i, 10).Value > 0 Then
                
                ws.Cells(i, 10).Interior.ColorIndex = 4

            Elseif ws.Cells(i, 10).Value < 0 Then
                
            ws.Cells(i, 10).Interior.ColorIndex = 3

            End If
        
            If ws.Cells(i, 11).Value > 0 Then
                
                ws.Cells(i, 11).Interior.ColorIndex = 4

            Elseif ws.Cells(i, 11).Value < 0 Then
               
            ws.Cells(i, 11).Interior.ColorIndex = 3

            End If

        Next i

    Next ws

End Sub




