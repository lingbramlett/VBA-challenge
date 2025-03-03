Attribute VB_Name = "VBA"
Sub Original_code()
    Dim ws As Worksheet
    Dim Ticker As String
    Dim I As Long
    Dim lastRow As Long
    Dim j As Integer
    Dim open_start As Double
    Dim close_last As Double
    Dim rng As Range
    Dim jcell As Range
    Dim TotalVol As Double
    Dim Greatest_increase As Variant
    Dim Greatest_decrease As Variant
    Dim Greatest_total As Double
    Dim start_row As Long
    Dim quarter_change As Double
    Dim quarter_percent As Double
    Dim Ticker_inc As String
    Dim Ticker_dec As String

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For Each ws In Worksheets
        j = 0
        TotalVol = 0
        start_row = 2
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For I = 2 To lastRow
            TotalVol = TotalVol + ws.Cells(I, 7).Value
            Ticker = ws.Cells(I, 1).Value

            If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then
                ' Setting ticker and quarterly change
                ws.Cells(2 + j, 9).Value = Ticker
                open_start = ws.Cells(start_row, 3).Value
                start_row = I + 1
                close_last = ws.Cells(I, 6).Value
                quarter_change = close_last - open_start
                If open_start <> 0 Then
                    quarter_percent = quarter_change / open_start
                Else
                    quarter_percent = 0
                End If

                ws.Cells(2 + j, 10).Value = quarter_change
                ws.Cells(2 + j, 11).Value = quarter_percent
                ws.Cells(2 + j, 11).NumberFormat = "0.00%"

                j = j + 1
            End If
        Next I

        lastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

        ' Find greatest increase
        Greatest_increase = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
        ws.Range("Q2").Value = ws.Cells(Greatest_increase + 1, 11).Value
        ws.Range("Q2").NumberFormat = "0.00%"
        Ticker_inc = ws.Cells(Greatest_increase + 1, 9).Value
        ws.Range("P2").Value = Ticker_inc

        ' Find greatest decrease
        Greatest_decrease = Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
        ws.Range("Q3").Value = ws.Cells(Greatest_decrease + 1, 11).Value
        ws.Range("Q3").NumberFormat = "0.00%"
        Ticker_dec = ws.Cells(Greatest_decrease + 1, 9).Value
        ws.Range("P3").Value = Ticker_dec

        ' Find greatest total volume
        lastRow = ws.Cells(Rows.Count, 12).End(xlUp).Row
        Greatest_total = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
        ws.Range("Q4").Value = Greatest_total
        ws.Range("P4").Value = ws.Cells(Application.WorksheetFunction.Match(Greatest_total, ws.Range("L2:L" & lastRow), 0) + 1, 9).Value
        ws.Range("Q4").NumberFormat = "0"

        ' Color formatting
        j = 2
        Set rng = ws.Range(ws.Cells(j, 10), ws.Cells(lastRow, 10))

        For Each jcell In rng
            If jcell.Value < 0 Then
                jcell.Interior.ColorIndex = 3
            ElseIf jcell.Value > 0 Then
                jcell.Interior.ColorIndex = 4
            Else
                jcell.Interior.ColorIndex = xlNone
            End If
        Next jcell
    Next ws

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
                jcell.Interior.ColorIndex = 4
            End If
            
            If jcell.Value = 0 Then
                jcell.Interior.ColorIndex = xlnocolorindex
            End If
        Next jcell
    Next ws

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


