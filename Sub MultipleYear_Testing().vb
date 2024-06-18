Sub MultipleYear_Testing()

'Create Variables
Dim ticker_symbol As String
Dim open_price As Double
Dim close_price As Double
Dim Qchange As Double
Dim Pchange As Double
Dim volume As Double
Dim Total_stock_volume As Double
Dim ws As Worksheet
Dim LastRow As Long
Dim table_column As Integer
Dim table_row As Integer

    'Loop through all sheets
    For Each ws In Worksheets

        volume = 0
        table_row = 2
        table_column = 9
        newticker = 2
    
        'Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Add the Column Headers
        ws.Cells(1, 9).Value = "Ticker"
        
        ws.Cells(1, 10).Value = "Quarterly Change"
        
        ws.Cells(1, 11).Value = "Percent Change"
        
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'Loop through each row to output ticket symbol
        For i = 2 To LastRow

            'Set the ticker, open price, and close price
            ticker_symbol = ws.Cells(i, 1).Value
            open_price = ws.Cells(newticker, 3).Value
            close_price = ws.Cells(i, 6).Value

            'Check if we are still within the same ticker and if we are not
            If ws.Cells(i + 1, 1).Value <> ticker_symbol Then

                'Find new ticker and print
                ws.Cells(table_row, table_column).Value = ticker_symbol

                'Find quarterly change
                Qchange = close_price - open_price

                'Use conditional formatting to highlight positive vs negative quarterly change
                If Qchange < 0 Then
                    ws.Cells(table_row, 10).Interior.ColorIndex = 3

                ElseIf Qchange > 0 Then
                    ws.Cells(table_row, 10).Interior.ColorIndex = 4
                End If

                'Find the percent change
                 If open_price <> 0 Then
                    Pchange = (Qchange / open_price)
                Else
                    Pchange = 0
                End If
                
                'Print quarterly change
                ws.Cells(table_row, 10).Value = Qchange
                
                'Print the percent change
                ws.Cells(table_row, 11).Value = Format(Pchange, "Percent")

                'Add to the Total Stock Volume using an inbuilt VBA function
                Total_stock_volume = WorksheetFunction.Sum(Range(ws.Cells(newticker, 7), ws.Cells(i, 7)))

                'Print the Total Stock Volume
                ws.Cells(table_row, 12).Value = Total_stock_volume

                Total_stock_volume = 0

                table_row = table_row + 1

                newticker = i + 1

            'If we are still within the same ticker
            Else
                Total_stock_volume = volume + ws.Cells(i, 7).Value

            End If
        
        'Iterate through all the rows starting from 2 to the last row
        Next i 

        'Add Column Headers

        ws.Cells(1, 16).Value = "Ticker"

        ws.Cells(1, 17).Value = "Value"

        ws.Cells(2, 15).Value = "Greatest % Increase"

        ws.Cells(3, 15).Value = "Greatest % Decrease"

        ws.Cells(4, 15).Value = "Greatest Total Volume"
            
        'Create Variables
        Dim greatestInc As Double
        Dim greatestDec As Double
        Dim greatestVol As Double

        'Initialize Variables
        greatestInc = ws.Cells(2,11).Value
        greatestDec = ws.Cells(2,11).Value
        greatestVol = ws.Cells(2,12).Value


        LastRowGreat = ws.Cells(Rows.Count, 9).End(xlUp).Row

        For i = 2 to LastRowGreat

            Pchange = ws.Cells(i,11).Value 

            If Pchange > greatestInc Then
                greatestInc = Pchange
                ws.Cells(2, 17).Value = Format(greatestInc, "Percent")
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            Else 
                greatestInc = greatestInc
            End If

            If Pchange < greatestDec Then
                greatestDec = Pchange
                ws.Cells(3, 17).Value = Format(greatestDec, "Percent")
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            Else
                greatestDec = greatestDec
            End If

            If ws.Cells(i,12).Value > greatestVol Then
                greatestVol = ws.Cells(i,12).Value
                ws.Cells(4, 17).Value = greatestVol
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            Else
                greatestVol= greatestVol
            End If
    

        'Iterate through all the rows
        Next i
    
        'Autofit to display data
        ws.Columns("A:Q").AutoFit
        
    'Iterate through all the worksheets
    Next ws


End Sub
