'Builds named ranges by worksheet for report range

Sub LabelDataRanges()
Dim theWkbk     As Workbook: Set theWkbk = ActiveWorkbook
Dim BPCObject   As Object:   Set BPCObject = buildBPCObject

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
Const preFixofRangeName = "dataRange"

With WS
    Dim theRange As Range
        Set theRange = Range(.Range(BPCObject.GetDataBottomRightCell(WS, RPT_ID)), _
            .Range(BPCObject.GetDataTopLeftCell(WS, RPT_ID)))
        
    Dim theAddress As String
        theAddress = "='" & WS.Name & "'!" & theRange.Address(True, True, xlR1C1)

    .Names.Add Name:=preFixofRangeName & RPT_ID, RefersToR1C1:=theAddress

End With


End Sub



Private Function buildBPCObject() As Object
    Const NoConnectMessage As String = "No Connection Found"
    Dim aoComAdd As Object, successConnection As Boolean, ObjAddOn As COMAddIn
 
        For Each ObjAddOn In Application.COMAddIns
            If ObjAddOn.progID = "FPMXLClient.Connect" Then
                'EPM/BPC
                Set buildBPCObject = ObjAddOn.Object
                successConnection = True
                Exit For
            ElseIf ObjAddOn.progID = "SapExcelAddIn" Then
                'Analysis for Office Version
                Set aoComAdd = ObjAddOn.Object
                Set buildBPCObject = aoComAdd.GetPlugin("com.sap.epm.FPMXLClient")
                successConnection = True
                Exit For
            End If
        Next ObjAddOn
     
        If Not successConnection Then
            MsgBox NoConnectMessage
            End
        End If
    
End Function
