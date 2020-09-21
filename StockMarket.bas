Attribute VB_Name = "Module1"
Sub StockMarket()
    Dim ws As Worksheet
    Dim rowi As Long
    Dim rowlen As Long
    
    Sheets.Add(bEFORE:=Sheets(1)).Name = "Combined_Data"
    
    For Each ws In Worksheets:
        If ws.Name <> "Combined_Data" Then
            rowlen = ws.Cells(rows.Count, 1).End(xlUp).Row
            MsgBox (rowlen)
            
        End If
        
    Next ws
        
End Sub
