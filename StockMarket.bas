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
    

    For Each ws In Worksheets:
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        output_row = 2
        rowlen = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For rowi = 1 To rowlen
            If ws.Cells(rowi, 1).Value <> ws.Cells(rowi + 1, 1) Then
                If IsEmpty(ws.Cells(rowi + 1, 1)) = False Then
                    ws.Cells(output_row, 9) = ws.Cells(rowi + 1, 1)
                    ws.Cells(output_row, 12).Value = ws.Cells(rowi + 1, 7).Value
                End If
                
                If rowi <> 1 Then
                    If ws.Cells(rowi, 6).Value <> 0 Then
                        close_price = ws.Cells(rowi, 6).Value
                    Else
                        For closezeroi = rowi To 1 Step -1
                            If ws.Cells(closezeroi, 1).Value = ws.Cells(rowi, 1).Value And ws.Cells(closezeroi, 6).Value <> 0 Then
                                close_price = ws.Cells(closezeroi, 6).Value
                            Else
                                close_price = 0
                            End If
                        Next closezeroi
                    End If
                    
                    If open_price <> 0 And close_price <> 0 Then
                        
                        ws.Cells(output_row - 1, 10).Value = close_price - open_price
                    
                        If ws.Cells(output_row - 1, 10).Value >= 0 Then
                            ws.Cells(output_row - 1, 10).Interior.ColorIndex = 4
                        Else
                            ws.Cells(output_row - 1, 10).Interior.ColorIndex = 3
                        End If
                        
                        ws.Cells(output_row - 1, 11).Value = (close_price - open_price) / open_price
                        ws.Cells(output_row - 1, 11).NumberFormat = "0.00%"
                    Else
                        ws.Cells(output_row - 1, 10).Value = "NA"
                        ws.Cells(output_row - 1, 11).Value = "NA"
                        ws.Cells(output_row - 1, 12).Value = "NA"
                    End If
                    
                End If
                    
                If ws.Cells(rowi + 1, 3).Value <> 0 Then
                    open_price = ws.Cells(rowi + 1, 3).Value
                Else
                    For openzeroi = rowi + 1 To rowlen
                        If ws.Cells(openzeroi, 1).Value = ws.Cells(rowi + 1, 1).Value And ws.Cells(openzeroi, 3).Value <> 0 Then
                            open_price = ws.Cells(openzeroi, 3).Value
                        Else
                            open_price = 0
                        End If
                    Next openzeroi
                End If
        
                output_row = output_row + 1
            Else
               volume = ws.Cells(rowi + 1, 7).Value
               ws.Cells(output_row - 1, 12).Value = ws.Cells(output_row - 1, 12).Value + volume
                        
            End If
        Next rowi
    Next
            
End Sub

