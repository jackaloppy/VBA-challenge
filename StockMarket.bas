Attribute VB_Name = "Module1"
Sub StockMarket()

    For Each ws In Worksheets:
        MsgBox (ws.Cells(2, 1).Value)
    Next ws
        
End Sub
