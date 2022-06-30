Private Const delayCheckInSeconds As Byte = 12 'Option to change delay to force full refresh
Dim lastRefresh As Date

Private Function getBPC_API() As Object

    Set getBPC_API = BPC_API 'or link to a different object calling API object

End Function


Private Function Before_Refresh() As Boolean
    Dim tempNow As Date
        tempNow = Now
    
    If tempNow - lastRefresh > TimeSerial(0, 0, delayCheckInSeconds) Then
        lastRefresh = tempNow
        Before_Refresh = False
        getBPC_API.RefreshActiveWorkbook
    End If
    
End Function

Private Function Before_Save() As Boolean
    Dim tempNow As Date
        tempNow = Now

    If tempNow - lastRefresh > TimeSerial(0, 0, delayCheckInSeconds) Then
        lastRefresh = tempNow
        Before_Save = False
        getBPC_API.SaveAndRefreshWorkbookData
    End If
    
End Function

'API could exist somewhere else, but listed here for illustration and to allow
'full copy paste
Private Function BPC_API() As Object
    Const NoConnectMessage As String = "No Connection Found"
    Dim aoComAdd As Object, successConnection As Boolean, ObjAddOn As COMAddIn
 
        For Each ObjAddOn In Application.COMAddIns
            If ObjAddOn.progID = "FPMXLClient.Connect" Then
                'EPM/BPC
                Set BPC_API = ObjAddOn.Object
                successConnection = True
                Exit For
            ElseIf ObjAddOn.progID = "SapExcelAddIn" Then
                'Analysis for Office Version
                Set aoComAdd = ObjAddOn.Object
                Set BPC_API = aoComAdd.GetPlugin("com.sap.epm.FPMXLClient")
                successConnection = True
                Exit For
            End If
        Next ObjAddOn
     
        If Not successConnection Then
            MsgBox NoConnectMessage
            End
        End If
End Function
