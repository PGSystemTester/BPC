Function bpcAPI() As Object
	Const NoConnectMessage As String = "No Connection Found"
	
	Dim aoComAdd As Object, isConnected As Boolean, ObjAddOn As COMAddIn
		For Each ObjAddOn In Application.COMAddIns
			If ObjAddOn.progID = "FPMXLClient.Connect" Then
				'EPM/BPC
				Set bpcAPI = ObjAddOn.Object
				isConnected = True
				Exit For
			ElseIf ObjAddOn.progID = "SapExcelAddIn" Then
				'Analysis for Office Version
				Set aoComAdd = ObjAddOn.Object
				Set bpcAPI = aoComAdd.GetPlugin("com.sap.epm.FPMXLClient")
				isConnected = True
				Exit For
			End If
		Next ObjAddOn
	 
		If Not isConnected Then
			MsgBox NoConnectMessage
			End
		End If
End Function
