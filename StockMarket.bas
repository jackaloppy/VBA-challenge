Attribute VB_Name = "Module1"
Sub StockMarket()
    Dim ws As Worksheet
    Dim rowi As Long
    Dim rowlen As Long
    Dim combined_row As Long
    Dim open_price As Double
    Dim close_price As Double
    
    Sheets.Add(bEFORE:=Sheets(1)).Name = "Combined_Data"
    Set combined_sheet = Worksheets("Combined_Data")
    combined_sheet.Range("A1").Value = "Ticker"
    combined_sheet.Range("B1").Value = "Yearly Change"
    combined_sheet.Range("C1").Value = "Percent Change"
    combined_sheet.Range("D1").Value = "Total Stock Volume"
    combined_sheet.Range("H1").Value = "Ticker"
    combined_sheet.Range("I1").Value = "Value"
    combined_row = 2
     
    For Each ws In Worksheets:
        If ws.Name <> "Combined_Data" Then
            rowlen = ws.Cells(rows.Count, 1).End(xlUp).Row
            For rowi = 1 To rowlen
                If ws.Cells(rowi, 1).Value <> ws.Cells(rowi + 1, 1) Then
                    If IsEmpty(ws.Cells(rowi + 1, 1)) = False Then
                        combined_sheet.Cells(combined_row, 1) = ws.Cells(rowi + 1, 1)
                    End If
                    
                    If Right(ws.Cells(rowi, 2).Value, 4) = "1230" Then
                        close_price = ws.Cells(rowi, 6).Value
                        combined_sheet.Cells(combined_row - 1, 2) = close_price - open_price
                    End If
                    
                    If Right(ws.Cells(rowi + 1, 2).Value, 4) = "0101" Then
                        open_price = ws.Cells(rowi + 1, 3).Value
                        combined_row = combined_row + 1
                    End If
    
                End If
            Next rowi
        End If
    Next
            
End Sub
