Sub Wall_Street():

    'Loop through worksheets

    For Each ws in worksheets

            ' Label  the new columns
            Range("I1").Value = "Ticker"
            Range("J1").Value = "Total Stock Volume"

            'Put the ticker symbol in the I column
            dim ticker_sym as String
            dim ticker_sym_row as integer
            ticker_sym_row = 2

            'Put the total stock volume in the J column
            dim total_stock_volume as double
            total_stock_volume = 0

            For i = 2 to 70926

                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                ticker_sym = Cells(i, 1).Value

                total_stock_volume = total_stock_volume + Cells(i, 7).Value

                Range("I" & ticker_sym_row).Value = ticker_sym

                Range("J" & ticker_sym_row).Value = total_stock_volume

                ticker_sym_row = ticker_sym_row + 1
                total_stock_volume = 0

                Else

                total_stock_volume = total_stock_volume + Cells(i, 7).Value

                End If
            Next i
    Next ws
    
End Sub