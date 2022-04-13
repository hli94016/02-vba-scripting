Sub Summary_Table()
    'Loop Through All worksheets
    For Each ws In Worksheets
        'Print out the headers for summary table #2
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        'Set new variables for the summary table #2
        Dim Greatest_Increase As Double
        Dim Greatest_Decrease As Double
        Dim Greatest_Volume As Double
        Dim Percent_Change_Number_1 As Double
        Dim Percent_Change_Number_2 As Double
        Dim Yearly_Change_1 As Double
        Dim Yearly_Change_2 As Double
        Dim Year_Start_1 As Double
        Dim Year_Start_2 As Double
        Dim Year_End_1 As Double
        Dim Year_End_2 As Double
        Dim Ticker_Greatest_Increase As String

        'Set up start value of variables
        Greatest_Increase = 0
        Greatest_Decrease = 0
        Greatest_Volume = 0

        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        Dim uniqueStockNum As Integer
        uniqueStockNum = 1
        'Loop through all rows in the combined data worksheet
        For i = 2 To lastRow
            'Check if we are still within the same Ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Increment the uniqueStockNum to know there is new stock ticket
                uniqueStockNum = uniqueStockNum + 1
            End If
            NextIteration:
        Next i

        For k = 2 To uniqueStockNum
            'Set up this ticker comparables
            Year_End_1 = ws.Cells(k, 20).Value
            Year_Start_1 = ws.Cells(k, 21).Value
            Yearly_Change_1 = Year_End_1 - Year_Start_1
            'Set up next ticker comparables
            Year_End_2 = ws.Cells(k + 1, 20).Value
            Year_Start_2 = ws.Cells(k + 1, 21).Value
            Yearly_Change_2 = Year_End_2 - Year_Start_2
            'If current ticker's yearly change <> 0, then
            If Yearly_Change_1 <> 0 And Yearly_Change_2 <> 0 And Year_Start_1 <> 0 And Year_Start_2 <> 0 Then
                Percent_Change_Number_1 = Yearly_Change_1 / Year_Start_1
                Percent_Change_Number_2 = Yearly_Change_2 / Year_Start_2
            Else
                Percent_Change_Number_1 = 0
                Percent_Change_Number_2 = 0
            End If

            If Percent_Change_Number_1 > Percent_Change_Number_2 And Percent_Change_Number_1 > Greatest_Increase Then
                Greatest_Increase = Percent_Change_Number_1
                Ticker_Greatest_Increase = ws.Cells(k, 9).Value
            ElseIf Percent_Change_Number_1 > Percent_Change_Number_2 And Percent_Change_Number_1 < Greatest_Increase Then
                Greatest_Increase = Greatest_Increase
                Ticker_Greatest_Increase = Ticker_Greatest_Increase
            ElseIf Percent_Change_Number_1 < Percent_Change_Number_2 And Percent_Change_Number_2 > Greatest_Increase Then
                Greatest_Increase = Percent_Change_Number_2
                Ticker_Greatest_Increase = ws.Cells(k + 1, 9).Value
            Else
                Greatest_Increase = Greatest_Increase
                Ticker_Greatest_Increase = Ticker_Greatest_Increase
            End If

            Dim Ticker_Greatest_Decrease As String
            If Percent_Change_Number_1 < Percent_Change_Number_2 And Percent_Change_1 < Greatest_Decrease Then
                Greatest_Decrease = Percent_Change_Number_1
                Ticker_Greatest_Decrease = ws.Cells(k, 9).Value
            ElseIf Percent_Change_Number_1 < Percent_Change_Number_2 And Percent_Change_Number_1 > Greatest_Decrease Then
                Greatest_Decrease = Greatest_Decrease
                Ticker_Greatest_Decrease = Ticker_Greatest_Decrease
            ElseIf Percent_Change_Number_1 > Percent_Change_Number_2 And Percent_Change_Number_2 < Greatest_Decrease Then
                Greatest_Decrease = Percent_Change_Number_2
                Ticker_Greatest_Decrease = ws.Cells(k + 1, 9).Value
            Else
                Greatest_Decrease = Greatest_Decrease
                Ticker_Greatest_Decrease = Ticker_Greatest_Decrease
            End If

            Dim Ticker_Greatest_Volume As String
            Dim Total_Volume_1 As Double
            Dim Total_Volume_2 As Double
            Total_Volume_1 = ws.Cells(k, 12).Value
            Total_Volume_2 = ws.Cells(k + 1, 12).Value
            If Total_Volume_1 > Total_Volume_2 And Total_Volume_1 > Greatest_Volume Then
                Greatest_Volume = Total_Volume_1
                Ticker_Greatest_Volume = ws.Cells(k, 9).Value
            ElseIf Total_Volume_1 > Total_Volume_2 And Total_Volume_1 < Greatest_Volume Then
                Greatest_Volume = Greatest_Volume
                Ticker_Greatest_Volume = Ticker_Greatest_Volume
            ElseIf Total_Volume_1 < Total_Volume_2 And Total_Volume_2 > Greatest_Volume Then
                Greatest_Volume = Total_Volume_2
                Ticker_Greatest_Volume = ws.Cells(k + 1, 9).Value
            Else
                Greatest_Volume = Greatest_Volume
                Ticker_Greatest_Volume = Ticker_Greatest_Volume
            End If

        Next k

        ws.Range("Q2").Value = FormatPercent(Greatest_Increase)
        ws.Range("P2").Value = Ticker_Greatest_Increase
        ws.Range("Q3").Value = FormatPercent(Greatest_Decrease)
        ws.Range("P3").Value = Ticker_Greatest_Decrease
        ws.Range("Q4").Value = Greatest_Volume
        ws.Range("P4").Value = Ticker_Greatest_Volume
    Next ws


End Sub

