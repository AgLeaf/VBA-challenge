Attribute VB_Name = "Module1"
Sub stock_market_analysis()

    ' create a loop through all the worksheets
    For Each ws In Worksheets
        ws.Activate
        
        ' output results will be populated in the range of column I to column L
        ' Ticker in column I
        ' Price Change in column J (last closing price in column F - first opening price in column C for a ticker)
        ' Percent Change in column K (Price Change / first opening price)
        ' Total Stock Volume in column L (the total of column G for a ticker)
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Price Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        ' set cell names for additional output
        Cells(2, "O").Value = "Greatest % Increase"
        Cells(3, "O").Value = "Greatest % Decrease"
        Cells(4, "O").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Columns("I:Q").AutoFit
        
        ' get the row count of the activated worksheet
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
        ' initialize variables
        Dim Ticker As String
        Ticker = ""
  
        Dim yearlyChange As Double
        yearlyChange = 0
  
        Dim openingPrice As Double
        openingPrice = 0
  
        Dim closingPrice As Double
        closingPrice = 0
        
        Dim volume As Variant
        totalVolume = 0
        
        ' when it is ready to populate the output results, the starting row is 2 in column I
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        ' loop through the rows
        For currentRow = 2 To RowCount
        
            ' if the ticker name of currentRow is different from the previous row, openingPrice is found in column C
            If Cells(currentRow, "A").Value <> Cells(currentRow - 1, "A").Value Then
                openingPrice = Cells(currentRow, "C").Value
                
                If (openingPrice = 0) Then
          
                    For pointer = currentRow To RowCount
                        If Cells(pointer + 1, "C").Value And (Cells(currentRow, "A").Value = Cells(currentRow + 1, "A").Value) Then
                            openingPrice = Cells(pointer + 1, "C").Value
                            pointer = RowCount + 100
                        End If
                        
                    Next pointer
                    
                End If

            End If
        
            ' if the ticker name of currentRow is the same ticker name in the next row, add the volume of column G to totalVolume
            If Cells(currentRow, "A").Value = Cells(currentRow + 1, "A").Value Then
                totalVolume = totalVolume + Cells(currentRow, "G").Value
                
            ' Otherwise, these two ticket names are different, closingPrice is found in column F for the ticker of currentRow
            Else
                closingPrice = Cells(currentRow, "F").Value
                 
                totalVolume = totalVolume + Cells(currentRow, "G").Value
                 
                yearlyChange = closingPrice - openingPrice
                 
                percentChange = Round((yearlyChange / openingPrice * 100), 2)
                
                If totalVolume = 0 Then
                    Cells(Summary_Table_Row, "I").Value = Cells(currentRow, "A").Value
                    Cells(Summary_Table_Row, "J").Value = 0
                    Cells(Summary_Table_Row, "K").Value = "0%"
                    Cells(Summary_Table_Row, "L").Value = 0
                Else
                    Cells(Summary_Table_Row, "I").Value = Cells(currentRow, "A").Value
                    Cells(Summary_Table_Row, "J").Value = yearlyChange
                    Cells(Summary_Table_Row, "K").Value = "%" & percentChange
                    Cells(Summary_Table_Row, "L").Value = totalVolume
                End If
                
                If yearlyChange > 0 Then
                    Cells(Summary_Table_Row, "J").Interior.ColorIndex = 4
                ElseIf yearlyChange < 0 Then
                    Cells(Summary_Table_Row, "J").Interior.ColorIndex = 3
                Else
                    Cells(Summary_Table_Row, "J").Interior.ColorIndex = 2
                End If
                             
                ' Reset variables for new stock ticker
                totalVolume = 0
                openingPrice = 0
                closingPrice = 0
                yearlyChange = 0
                Summary_Table_Row = Summary_Table_Row + 1
            End If
        
        Next currentRow
       
    Next ws
    
    MsgBox ("Done")
    
End Sub


