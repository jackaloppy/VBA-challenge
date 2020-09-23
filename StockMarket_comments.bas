Attribute VB_Name = "Module1"
Sub StockMarket()

    Dim ws As Worksheet
    Dim rowi As Long
    Dim rowlen As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim ouput_row As Long
    Dim volume As Long
    Dim openzeroi As Long
    Dim closezeroi As Long
    Dim maxpercent As Double
    Dim minpercent As Double
    Dim maxv As Double
    
    For Each ws In Worksheets:
        'Add header to each worksheet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        output_row = 2
        rowlen = ws.Cells(Rows.Count, 1).End(xlUp).Row
        maxpercent = 0
        minpercent = 0
        maxv = 0
        
        For rowi = 1 To rowlen
            'Find the begining and the end of one stock based on ticker name
            If ws.Cells(rowi, 1).Value <> ws.Cells(rowi + 1, 1) Then
                'Empty cell false to exclude the blank row in the end for each sheet
                If IsEmpty(ws.Cells(rowi + 1, 1)) = False Then
                    'Output ticker name
                    ws.Cells(output_row, 9) = ws.Cells(rowi + 1, 1)
                    'Set the first volume for each sotck
                    ws.Cells(output_row, 12).Value = ws.Cells(rowi + 1, 7).Value
                End If
                
                'Intentionally skip the first iteration for retrieving close price,
                'so that it retrieves open price first, then in the second iteration,
                'it retrieves the close price of the previous stock.
                If rowi <> 1 Then
                    'To find out data containing 0 value
                    If ws.Cells(rowi, 6).Value <> 0 Then
                        close_price = ws.Cells(rowi, 6).Value
                    Else
                        'If the stock has 0 as close price, then it iterate backwards to
                        'find the possible non-zero close price at earily time point within
                        'the same stock. If all 0, then close price is 0.
                        For closezeroi = rowi To 1 Step -1
                            If ws.Cells(closezeroi, 1).Value = ws.Cells(rowi, 1).Value And ws.Cells(closezeroi, 6).Value <> 0 Then
                                close_price = ws.Cells(closezeroi, 6).Value
                            Else
                                close_price = 0
                            End If
                        Next closezeroi
                    End If
                    
                    'Calculation with meaningful data (non-zero)
                    If open_price <> 0 And close_price <> 0 Then
                        
                        ws.Cells(output_row - 1, 10).Value = close_price - open_price
                    
                        If ws.Cells(output_row - 1, 10).Value >= 0 Then
                            ws.Cells(output_row - 1, 10).Interior.ColorIndex = 4
                        Else
                            ws.Cells(output_row - 1, 10).Interior.ColorIndex = 3
                        End If
                        
                        ws.Cells(output_row - 1, 11).Value = (close_price - open_price) / open_price
                        ws.Cells(output_row - 1, 11).NumberFormat = "0.00%"
                        'The following three if statements are used to generate the maximum
                        'percentage, the minimum percentage, and the maximum total volume
                        'note that total v should be dimed as double because long won't
                        'hold up this large value.
                        If maxpercent <= ws.Cells(output_row - 1, 11).Value Then
                            maxpercent = ws.Cells(output_row - 1, 11).Value
                            ws.Cells(2, 17).Value = maxpercent
                            ws.Cells(2, 17).NumberFormat = "0.00%"
                            ws.Cells(2, 16).Value = ws.Cells(output_row - 1, 9).Value
                        End If
                        
                        If minpercent >= ws.Cells(output_row - 1, 11).Value Then
                            minpercent = ws.Cells(output_row - 1, 11).Value
                            ws.Cells(3, 17).Value = minpercent
                            ws.Cells(3, 17).NumberFormat = "0.00%"
                            ws.Cells(3, 16).Value = ws.Cells(output_row - 1, 9).Value
                        End If
                        
                        If maxv <= ws.Cells(output_row - 1, 12).Value Then
                            maxv = ws.Cells(output_row - 1, 12).Value
                            ws.Cells(4, 17).Value = maxv
                            ws.Cells(4, 16).Value = ws.Cells(output_row - 1, 9).Value
                        End If
                            
                    'If open price or close price is 0, then output NA
                    Else
                        ws.Cells(output_row - 1, 10).Value = "NA"
                        ws.Cells(output_row - 1, 11).Value = "NA"
                        ws.Cells(output_row - 1, 12).Value = "NA"
                    End If
                    
                End If
                    
                'Retrieving non-zero open price here
                If ws.Cells(rowi + 1, 3).Value <> 0 Then
                    open_price = ws.Cells(rowi + 1, 3).Value
                Else
                    'If 0, to find possible non-zero open price at later time point within
                    'the same stock. If no result, then output open price as 0
                    For openzeroi = rowi + 1 To rowlen
                        If ws.Cells(openzeroi, 1).Value = ws.Cells(rowi + 1, 1).Value And ws.Cells(openzeroi, 3).Value <> 0 Then
                            open_price = ws.Cells(openzeroi, 3).Value
                        Else
                            open_price = 0
                        End If
                    Next openzeroi
                End If
                'Counting up for output row
                output_row = output_row + 1
            'If two rows are identical, then adding the value to total stock volume
            Else
               volume = ws.Cells(rowi + 1, 7).Value
               ws.Cells(output_row - 1, 12).Value = ws.Cells(output_row - 1, 12).Value + volume
                        
            End If
        Next rowi
    Next
            
End Sub
