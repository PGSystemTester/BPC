'Run this macro with the sheet you wish to clean selected
Sub clearLinksToOtherSheets()
    Const turnOffEvents As Boolean = False 'set this to true if process crashes.

    Dim aCell As Range, thisWS As Worksheet, ws As Worksheet
    Set thisWS = ActiveSheet
    
    If turnOffEvents Then
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
    End If
    
    
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
    
    If turnOffEvents Then
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
    End If
    
End Sub
