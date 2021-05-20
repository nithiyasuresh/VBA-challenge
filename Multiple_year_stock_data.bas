Attribute VB_Name = "Module1"
Sub WorksheetLoop_YearlyChange()
         Dim WS_Count As Integer
         Dim I As Integer
         Dim j As Long
         Dim number_tickers As Integer
         Dim yearly_change As Double
         Dim percentage As Double
         Dim opening_price As Double
         Dim closing_price As Double
         Dim stock_volume As Double
         
         WS_Count = ActiveWorkbook.Worksheets.Count
         
         For I = 1 To WS_Count
                ActiveWorkbook.Worksheets(I).Activate
                ticker_row = 2
                yearly_change = 0
                percentage = 0
                opening_price = 0
                number_tickers = 0
                stock_volume = 0
                Cells(1, "J").Value = "ticker_list"
                Cells(1, "K").Value = "yearly_change_test"
                Cells(1, "L").Value = "percentage"
                Cells(1, "M").Value = "stock_volume"
                
                For j = 2 To ActiveWorkbook.Worksheets(I).UsedRange.Rows.Count
                        If opening_price = 0 Then
                        opening_price = Cells(j, 3).Value
                        End If
                        
                        stock_volume = stock_volume + Cells(j, 7).Value
                
                        If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
                        ticker = Cells(j, 1).Value
                        Cells(ticker_row, "J").Value = ticker
                        number_tickers = number_tickers + 1
                        Cells(number_tickers + 1, 10) = Cells(j, 1).Value
                        closing_price = Cells(j, 6)
                        yearly_change = closing_price - opening_price
                        Cells(number_tickers + 1, 11).Value = yearly_change
                        
            If yearly_change > 0 Then
                Cells(number_tickers + 1, 11).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                Cells(number_tickers + 1, 11).Interior.ColorIndex = 3
            Else
                Cells(number_tickers + 1, 11).Interior.ColorIndex = 6
            End If
                        
                        If opening_price = 0 Then
                        percentage = 0
                        Else
                        percentage = (yearly_change / opening_price)
                        End If
                        Cells(number_tickers + 1, 12).Value = Format(percentage, "Percent")
                        
                        Cells(number_tickers + 1, 13).Value = stock_volume
                        End If
                                   
            If percentage > 0 Then
           Cells(number_tickers + 1, 12).Interior.ColorIndex = 4
            ElseIf percentage < 0 Then
            Cells(number_tickers + 1, 12).Interior.ColorIndex = 3
            Else
            Cells(number_tickers + 1, 12).Interior.ColorIndex = 6
            End If

                Next j

         Next I
         
      End Sub


