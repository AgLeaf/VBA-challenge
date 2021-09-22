Attribute VB_Name = "Module1"
Sub oneyear():
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
  Dim Ticker As String
  
  Dim yearlyChange As Double
  yearlyChange = 0
  
  ' dim percentChange as double
  
  Dim volume As Variant
      
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  Dim ws As Worksheet

  For Each ws In Worksheets
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  Next ws

  For i = 2 To lastrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
      Ticker = Cells(i, 1).Value
      yearlyChange = yearlyChange + Cells(i, 6).Value - Cells(i, 3)
      Range("I" & Summary_Table_Row).Value = Ticker
      Range("J" & Summary_Table_Row).Value = yearlyChange
      Summary_Table_Row = Summary_Table_Row + 1
    End If
      
      yearlyChange = 0
      
  Next i

If Range("J2").Value > 0 Then
        Range("J2").Interior.ColorIndex = 4
    ElseIf Range("J2").Value < 0 Then
        Range("J2").Interior.ColorIndex = 3
End If

End Sub



