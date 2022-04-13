Sub StockMarket()

    'Loop Through All worksheets
    For Each ws In Worksheets
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

        'Set header of the summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Volume"
        ws.Range("T1").Value = "Year End"
        ws.Range("U1").Value = "Year Beginning"

        'Set an initial variable for holding the ticker name
        Dim Ticker_Name As String

        'Set an initial variables per ticker
        Dim Yearly_Change As Double
        Dim Year_End As Double
        Dim Year_Beginning As Double
        Dim Total_Volume As Double
        Total_Volume = 0
        'Keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        'Print year beginning price for ticker A
        ws.Range("U2").Value = ws.Range("C2").Value

        Dim uniqueStockNum As Integer
        uniqueStockNum = 1
        'Loop through all rows in the combined data worksheet
        For i = 2 To lastRow
            'Check if we are still within the same Ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Increment the uniqueStockNum to know there is new stock ticket
                uniqueStockNum = uniqueStockNum + 1
                'Set the Ticker name
                Ticker_Name = ws.Cells(i, 1).Value
                'Add to the Total Volume
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                'Print the ticker name in the summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                'Set the year end close price per ticker
                Year_End = ws.Cells(i, 6).Value
                'Print the year end close price in the summary table
                ws.Range("T" & Summary_Table_Row).Value = Year_End
                'Print the Total Volume...
                ws.Range("L" & Summary_Table_Row).Value = Total_Volume
                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1

                'Set the year beginning price for next ticker
                Year_Beginning = ws.Cells(i + 1, 3)
                'Print the year beginning price for next ticker
                ws.Range("U" & Summary_Table_Row).Value = Year_Beginning
                'Reset the Year_End and Year_Beginning
                Year_End = 0
                Year_Beginning = 0
                Total_Volume = 0

            'If the cell immediately following a row is the same ticker...
            Else
                'Add to the Total Volume
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value

            End If
            NextIteration:
        Next i

        'looping through all rows within the summary table
        For j = 2 To uniqueStockNum
            'set yearly change value, current year end, current year start
            Dim Current_Year_End As Double
            Dim Current_Year_Start As Double

            Current_Year_End = ws.Cells(j, 20).Value
            Current_Year_Start = ws.Cells(j, 21).Value
            Yearly_Change = Current_Year_End - Current_Year_Start

            'Print yearly change
            ws.Cells(j, 10).Value = Yearly_Change
            'highlight positive and negative yearly_change
            If ws.Cells(j, 10) >= 0 Then
                'Highlight the cell as green
                ws.Cells(j, 10).Interior.ColorIndex = 4
                'Otherwise, highlight the cell as red
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If

            Dim Percent_Change As String
            Dim Percent_Change_Number As Double
            If Yearly_Change <> 0 And Current_Year_Start <> 0 Then
                Percent_Change_Number = Yearly_Change / Current_Year_Start
                Percent_Change = FormatPercent(Percent_Change_Number)
                ws.Cells(j, 11).Value = Percent_Change
            Else
                ws.Cells(j, 11).Value = 0
            End If

        Next j
    Next ws
End Sub
