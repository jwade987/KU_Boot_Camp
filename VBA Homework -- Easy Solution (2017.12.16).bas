Attribute VB_Name = "Module1"
Sub ticker()

Range("I1").Value = "Ticker"
Range("J1").Value = "Total Stock Volume"
Range("C:F").NumberFormat = "#,##0.00"
Range("G:G").NumberFormat = "#,##0"
Range("J:J").NumberFormat = "#,##0"

Range("C1:G1").HorizontalAlignment = xlRight

Range("J1").HorizontalAlignment = xlRight

  Dim Ticker_Name As String

  Dim Ticker_Total As Double
  Ticker_Total = 0

  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  For i = 2 To 797712

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      Ticker_Name = Cells(i, 1).Value
      Ticker_Total = Ticker_Total + Cells(i, 7).Value
      Range("I" & Summary_Table_Row).Value = Ticker_Name
      Range("J" & Summary_Table_Row).Value = Ticker_Total
      Summary_Table_Row = Summary_Table_Row + 1
      Ticker_Total = 0

    Else

      Ticker_Total = Ticker_Total + Cells(i, 7).Value

    End If

  Next i

End Sub
