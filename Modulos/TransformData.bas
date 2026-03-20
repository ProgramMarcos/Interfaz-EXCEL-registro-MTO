Attribute VB_Name = "TransformData"
'------------------------------------------------------------------
'------------------------------------------------------------------
'---    MARCOS LÓPEZ LÓPEZ
'---    2024/2025
'---    MODULO TransformData
'------------------------------------------------------------------
'------------------------------------------------------------------

Sub TransformData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Resultados")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = lastRow To 2 Step -1
        If ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value And ws.Cells(i, 2).Value = ws.Cells(i - 1, 2).Value _
        And ws.Cells(i, 3).Value = ws.Cells(i - 1, 3).Value Then
            ws.Cells(i, 2).ClearContents
            ws.Cells(i, 3).ClearContents
            ws.Cells(i, 4).ClearContents
            ws.Cells(i, 5).ClearContents
        End If
    Next i
End Sub

