Attribute VB_Name = "Module1"
Sub StockMarket()
    Dim ws As Worksheet
    
    Sheets.Add(bEFORE:=Sheets(1)).Name = "Combined_Data"
    
    For Each ws In Worksheets:
        If ws.Name <> "Combined_Data" Then
            MsgBox (ws.Cells(2, 1).Value)
        End If
        
    Next ws
        
End Sub
