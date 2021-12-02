Sub clearLinksToOtherSheets()
'Run this macro with the sheet you wish to clean selected
    Dim aCell As Range, thisWS As Worksheet, ws As Worksheet
    Set thisWS = ActiveSheet
    
    Application.EnableEvents = False
    
    For Each aCell In thisWS.UsedRange.Cells
        If Not Application.WorksheetFunction.IsFormula(aCell) Then
            'skip
        ElseIf InStr(1, aCell.Formula, "!", vbTextCompare) > 0 Then
            For Each ws In Workbooks(thisWS.Parent.Name).Worksheets
                If InStr(1, aCell.Formula, ws.Name, vbTextCompare) > 0 And ws.Name <> thisWS.Name Then
                    aCell.Value = aCell.Value
                    Exit For
                End If
            Next ws
        End If
    Next aCell
    Application.EnableEvents = True
End Sub
