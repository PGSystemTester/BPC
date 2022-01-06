Sub LabelDataRanges()
    Dim theWkbk As Workbook, BPCObject As Object
 
        Set theWkbk = ThisWorkbook

    Dim WS As Worksheet

    For Each WS In theWkbk.Worksheets
        Call setRangesOnSheet(WS, BPCObject)
    Next WS
End Sub


Private Sub setRangesOnSheet(WS As Worksheet, BPCObject As Object)
    Dim allReports() As String
    allReports = BPCObject.GetAllReportNames(WS)

    Dim i As Long

    For i = LBound(allReports) To UBound(allReports)
        Call setReportRanges(allReports(i), WS, BPCObject)
    Next i
End Sub

Private Sub setReportRanges(RPT_ID As String, WS As Worksheet, BPCObject As Object)
'Adds named ranges to all sheets of report data after a refresh.
'Must include BPCCom when calling LabelDataRanges!!
Const theComment = " data range."
Const preFixofRangeName = "RPT_"

    With WS
    Dim theRange As Range, theAddress As String, theName As String
            Set theRange = Range(.Range(BPCObject.GetDataBottomRightCell(WS, RPT_ID)), _
                .Range(BPCObject.GetDataTopLeftCell(WS, RPT_ID)))

            theAddress = "='" & WS.Name & "'!" & theRange.Address(True, True, xlR1C1)
            theName = preFixofRangeName & RPT_ID

        .Names.Add Name:=theName, RefersToR1C1:=theAddress

        If theComment <> "" Then .Names(theName).Comment = theName & theComment
    End With
End Sub
