Attribute VB_Name = "BPC_API_MODULE"

Function BPC_API() As Object
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
