Sub stock_data_2014()

  Dim ticker As String
  
  Dim Total_stock_volume As Double
  Total_stock_volume = 0
  
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  For i = 2 To 760192

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ticker = Cells(i, 1).Value

      Total_stock_volume = Total_stock_volume + Cells(i, 7).Value

      Range("I" & Summary_Table_Row).Value = ticker

      Range("J" & Summary_Table_Row).Value = Total_stock_volume

      Summary_Table_Row = Summary_Table_Row + 1
      
      Total_stock_volume = 0

    Else

      Total_stock_volume = Total_stock_volume + Cells(i, 7).Value

    End If

  Next i

End Sub
