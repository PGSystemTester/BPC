'this procedure will use a named range to store the last user to refresh. Typically should be used after a BPC Refresh
'BPC Routines

Private Sub setUserRefresh()
    Const nRefreshUser As String = "RefreshUser"
    On Error GoTo firstTime
    ThisWorkbook.Names(nRefreshUser).RefersTo = Evaluate("=EPMUSER()")
    On Error GoTo 0

Exit Sub
firstTime:
    Const showNamedRange As Boolean = False
    ThisWorkbook.Names.Add nRefreshUser, Evaluate("=EPMUSER()"), showNamedRange

End Sub
