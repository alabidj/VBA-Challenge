VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Assignment_V2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Data()

Dim ws As Worksheet
Dim i As Long
Dim Ticker As String
Dim Total_Stock_Volume As Double ' Changed to Double
Dim Quarterly_Change As Double
Dim Percent_Change As Double
Dim FirstRow As Long
Dim LastRow As Long
Dim OpenFirst As Double
Dim CloseLast As Double

For Each ws In Worksheets
    If ws.Name = "Q1" Or ws.Name = "Q2" Or ws.Name = "Q3" Or ws.Name = "Q4" Then ' Process specific worksheets only
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Reference the specific worksheet
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        ' Add column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ' Add text labels
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest % Volume"

        For i = 2 To LastRow
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then ' Reference the specific worksheet

                Ticker = ws.Cells(i, 1).Value
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

                ' Identify the first row of the current ticker
                FirstRow = i
                Do While ws.Cells(FirstRow - 1, 1).Value = Ticker And FirstRow > 1
                    FirstRow = FirstRow - 1
                Loop

                OpenFirst = ws.Cells(FirstRow, 3).Value ' Reference the specific worksheet
                CloseLast = ws.Cells(i, 6).Value ' Reference the specific worksheet

                Quarterly_Change = CloseLast - OpenFirst
                Percent_Change = (CloseLast / OpenFirst) - 1

                ws.Range("I" & Summary_Table_Row).Value = Ticker ' Reference the specific worksheet
                ws.Range("J" & Summary_Table_Row).Value = Quarterly_Change ' Reference the specific worksheet
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change ' Reference the specific worksheet
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume ' Reference the specific worksheet

                Summary_Table_Row = Summary_Table_Row + 1
                Total_Stock_Volume = 0

            Else
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value ' Reference the specific worksheet
            End If
        Next i

        ' Find the greatest % increase, decrease, and volume
        Dim MaxIncrease As Double
        Dim MaxDecrease As Double
        Dim MaxVolume As Double
        Dim MaxIncreaseTicker As String
        Dim MaxDecreaseTicker As String
        Dim MaxVolumeTicker As String

        MaxIncrease = WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table_Row - 1))
        MaxDecrease = WorksheetFunction.Min(ws.Range("K2:K" & Summary_Table_Row - 1))
        MaxVolume = WorksheetFunction.Max(ws.Range("L2:L" & Summary_Table_Row - 1))

        MaxIncreaseTicker = ws.Cells(WorksheetFunction.Match(MaxIncrease, ws.Range("K2:K" & Summary_Table_Row - 1), 0) + 1, 9).Value
        MaxDecreaseTicker = ws.Cells(WorksheetFunction.Match(MaxDecrease, ws.Range("K2:K" & Summary_Table_Row - 1), 0) + 1, 9).Value
        MaxVolumeTicker = ws.Cells(WorksheetFunction.Match(MaxVolume, ws.Range("L2:L" & Summary_Table_Row - 1), 0) + 1, 9).Value

        ws.Range("O2").Value = MaxIncreaseTicker
        ws.Range("P2").Value = MaxIncrease

        ws.Range("O3").Value = MaxDecreaseTicker
        ws.Range("P3").Value = MaxDecrease

        ws.Range("O4").Value = MaxVolumeTicker
        ws.Range("P4").Value = MaxVolume

    End If
Next ws

End Sub

