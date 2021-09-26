Attribute VB_Name = "Module2"
Sub greatestFind()
        
        Dim sheet1 As Worksheet
        Dim sheet2 As Worksheet
        Dim sheet3 As Worksheet
        
        Set sheet1 = Worksheets("2016")
        Set sheet2 = Worksheets("2015")
        Set sheet3 = Worksheets("2014")
        
        Dim maxIncrease, maxDecrease As Double
        Dim maxVolume As LongLong
        
        maxIncrease = WorksheetFunction.Max(sheet1.Range("K2:K10000"))
        sheet1.Range("Q2").Value = maxIncrease
        maxDecrease = WorksheetFunction.Min(sheet1.Range("K2:K10000"))
        sheet1.Range("Q3").Value = maxDecrease
        maxVolume = WorksheetFunction.Max(sheet1.Range("L2:L10000"))
        sheet1.Range("Q4").Value = maxVolume
        
        Columns("I:Q").AutoFit
        
        maxIncrease = WorksheetFunction.Max(sheet2.Range("K2:K10000"))
        sheet2.Range("Q2").Value = maxIncrease
        maxDecrease = WorksheetFunction.Min(sheet2.Range("K2:K10000"))
        sheet2.Range("Q3").Value = maxDecrease
        maxVolume = WorksheetFunction.Max(sheet2.Range("L2:L10000"))
        sheet2.Range("Q4").Value = maxVolume
        
        Columns("I:Q").AutoFit
        
        maxIncrease = WorksheetFunction.Max(sheet3.Range("K2:K10000"))
        sheet3.Range("Q2").Value = maxIncrease
        maxDecrease = WorksheetFunction.Min(sheet3.Range("K2:K10000"))
        sheet3.Range("Q3").Value = maxDecrease
        maxVolume = WorksheetFunction.Max(sheet3.Range("L2:L10000"))
        sheet3.Range("Q4").Value = maxVolume
        
        Columns("I:Q").AutoFit
                
        tickerRow = 0
        
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
        For tickerRow = 2 To RowCount
        
            If sheet1.Cells(tickerRow, "K").Value = sheet1.Cells(2, "Q").Value Then
                 sheet1.Cells(2, "P").Value = sheet1.Cells(tickerRow, "I").Value
            End If
            
            If sheet2.Cells(tickerRow, "K").Value = sheet2.Cells(2, "Q").Value Then
                 sheet2.Cells(2, "P").Value = sheet2.Cells(tickerRow, "I").Value
            Else
            End If
            
            If sheet3.Cells(tickerRow, "K").Value = sheet3.Cells(2, "Q").Value Then
                 sheet3.Cells(2, "P").Value = sheet3.Cells(tickerRow, "I").Value
            Else
            End If
        
        Next tickerRow
        
        tickerRow = 0
        
        For tickerRow = 2 To RowCount
        
            If sheet1.Cells(tickerRow, "K").Value = sheet1.Cells(3, "Q").Value Then
                 sheet1.Cells(3, "P").Value = sheet1.Cells(tickerRow, "I").Value
            End If
            
            If sheet2.Cells(tickerRow, "K").Value = sheet2.Cells(3, "Q").Value Then
                 sheet2.Cells(3, "P").Value = sheet2.Cells(tickerRow, "I").Value
            Else
            End If
            
            If sheet3.Cells(tickerRow, "K").Value = sheet3.Cells(3, "Q").Value Then
                 sheet3.Cells(3, "P").Value = sheet3.Cells(tickerRow, "I").Value
            Else
            End If
        
        Next tickerRow
        
        tickerRow = 0
        
        For tickerRow = 2 To RowCount
        
            If sheet1.Cells(tickerRow, "L").Value = sheet1.Cells(4, "Q").Value Then
                 sheet1.Cells(4, "P").Value = sheet1.Cells(tickerRow, "I").Value
            End If
            
            If sheet2.Cells(tickerRow, "L").Value = sheet2.Cells(4, "Q").Value Then
                 sheet2.Cells(4, "P").Value = sheet2.Cells(tickerRow, "I").Value
            Else
            End If
            
            If sheet3.Cells(tickerRow, "L").Value = sheet3.Cells(4, "Q").Value Then
                 sheet3.Cells(4, "P").Value = sheet3.Cells(tickerRow, "I").Value
            Else
            End If
        
        Next tickerRow
        
sheet1.Range("Q2").Value = FormatPercent(sheet1.Range("Q2"))
sheet1.Range("Q3").Value = FormatPercent(sheet1.Range("Q3"))
sheet2.Range("Q2").Value = FormatPercent(sheet2.Range("Q2"))
sheet2.Range("Q3").Value = FormatPercent(sheet2.Range("Q3"))
sheet3.Range("Q2").Value = FormatPercent(sheet3.Range("Q2"))
sheet3.Range("Q3").Value = FormatPercent(sheet3.Range("Q3"))

End Sub
