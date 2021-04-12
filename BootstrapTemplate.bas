'Standard Routines I use in Most Templates

Private lastRefresh As Double
Private beginTimeOfRefresh As Date
Private stopTimer As Boolean

Private Function After_ContextChange()

If Now - lastRefresh > TimeSerial(0, 0, 5) Then
   'this stops duplicative refreshes that happen with context on the as a formula
    stopTimer = True
    BPC_API_COM.RefreshActiveSheet
    lastRefresh = Now
    stopTimer = False
End If
End Function

Private Function Before_ContextChange()
  'infrequently used

End Function

Private Function before_Refresh()
  ''to cancel refresh

  'timer
  If Not stopTimer Then
      beginTimeOfRefresh = Now
      Application.StatusBar = "Refesh Started At " & beginTimeOfRefresh
  End If

End Function


Private Function BEFORE_SAVE

   'to cancel
   'Before_Save = false
End Function
   
   
Private Function after_Refresh()
  'after refresh code

  If Not stopTimer Then
      Dim timeInSeconds As Long
     timeInSeconds = Round((Now - beginTimeOfRefresh) * 24 * 3600, 2)

      If timeInSeconds <= 2 Then
          Application.StatusBar = "Completed refresh in less than two seconds."
      Else
          Application.StatusBar = "Completed time in " & timeInSeconds & " seconds."
      End If

  End If
End Function

Sub refreshWorksheet()
    BPC_API_COM.RefreshActiveWorkbook
End Sub


Sub refreshWorkbook()
    BPC_Code.RefreshActiveWorkbook
End Sub

Private Function BPC_API_COM() As Object
Const NoConnectMessage As String = "No Connection Found"
      
    Dim aoComAdd As Object, successConnection As Boolean, ObjAddOn As COMAddIn
        For Each ObjAddOn In Application.COMAddIns
            If ObjAddOn.progID = "FPMXLClient.Connect" Then
                'EPM/BPC
               Set BPC_API_COM = ObjAddOn.Object
                successConnection = True
                Exit For
            ElseIf ObjAddOn.progID = "SapExcelAddIn" Then
                'Analysis for Office Version
               Set aoComAdd = ObjAddOn.Object
                Set BPC_API_COM = aoComAdd.GetPlugin("com.sap.epm.FPMXLClient")
                successConnection = True
                Exit For
            End If
        Next ObjAddOn
        If Not successConnection Then
            MsgBox NoConnectMessage
            End
        End If

End Function


Private Sub after_workbook_Open()
    'runs when workbook opens from EPM menu
    Application.CalculateFull'useful for updating context cells
    'BPC_API_COM.RefreshActiveWorkbook
    
End Sub
