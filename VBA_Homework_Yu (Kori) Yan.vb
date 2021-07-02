Sub Stock()
    
    'Apply to all the worksheet
    Dim ws As Worksheet

    'Create the For loop
    For Each ws In Worksheets

        'Declare variables
        Dim symbol As String

        Dim total_vol As Double
        total_vol = 0

        Dim stock_open As Double
        stock_open = 0

        Dim stock_close As Double
        stock_close = 0
        
        Dim stock_change As Double
        stock_change = 0

        Dim percent_change As Double
        percent_change = 0
        
        Dim rowcount As Long
        rowcount = 2

        Dim total_rows As Long
        total_rows = ws.Cells(1, 1).End(xlDown).Row

        'Name the header using cells
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'Conditional
        For i = 2 To total_rows
            
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

                stock_open = ws.Cells(i, 3).Value

            End If

            'Calculate the total stock volume
            total_vol = total_vol + ws.Cells(i, 7)

            'Copy the ticker symbol
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value

                'Copy the total stock volume
                ws.Cells(rowcount, 12).Value = total_vol

                'Calculate the price change
                stock_close = ws.Cells(i, 6).Value

                stock_change = stock_close - stock_open
                ws.Cells(rowcount, 10).Value = stock_change

                If stock_change >= 0 Then
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 3
                End If

                'Calculate the percent change
                If stock_open = 0 And stock_close = 0 Then
                    percent_change = 0
                    ws.Cells(rowcount, 11).Value = percent_change
                    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                ElseIf stock_open = 0 Then
                    ws.Cells(rowcount, 11).Value = percent_change
                Else
                    percent_change = stock_change / stock_open
                    ws.Cells(rowcount, 11).Value = percent_change
                    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                End If

                'Move it to the next row
                rowcount = rowcount + 1

                'Reset the values
                total_vol = 0
                stock_open = 0
                stock_close = 0
                stock_change = 0
                percent_change = 0
                
            End If
        Next i
    Next ws
End Sub
