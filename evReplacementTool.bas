'This should replace all legacy BPC "ev" formulas and replace them with EPM functions.
'Do not use this unless you know what you're doing.

Public Sub replace_All_EV_Formulas_In_All_Open_Workbooks()
Dim theWorkbook As Workbook
    
For Each theWorkbook In Application.Workbooks
    Call wkbkSheetReplace(theWorkbook)
Next theWorkbook
    
End Sub


Public Sub replace_All_EV_Formulas_In_ActiveWorkbook()

    Call wkbkSheetReplace(ActiveWorkbook)

End Sub

Private Sub wkbkSheetReplace(theWorkbook As Workbook)
Dim WS As Worksheet

For Each WS In theWorkbook.Worksheets
    Call FindReplaceEv_With_EP(WS)
Next WS

End Sub



Private Sub FindReplaceEv_With_EP(WS As Worksheet)

With WS.Cells


'EVDES
    .Replace What:="evdes(", Replacement:="EPMmemberdesc("
        
'EVPRO
    .Replace What:="evpro(", Replacement:="EPMmemberproperty("

'EVTIM
    .Replace What:="evtim(", Replacement:="EPMmemberoffset("
        
'EVCOM
    .Replace What:="evcom(", Replacement:="EPMSaveComment("

'EVRNG(
    .Replace What:="evrng(", Replacement:="EPMCellRanges("

'EVCVW
    .Replace What:="evcvw(", Replacement:="EPMContextMember("

'EVUSR
    .Replace What:="evusr(", Replacement:="EPMUser("
        
'EVBET
    .Replace What:="evbet(", Replacement:="EPMComparison("
        
'EVGET
    .Replace What:="evget(", Replacement:="EPMRetrieveData("
        
'EVSND
    .Replace What:="evsnd(", Replacement:="EPMSaveData("

'EVGTS
    .Replace What:="evgts(", Replacement:="EPMScaleData("
    
'EVSVR
    .Replace What:="evsvr(", Replacement:="EPMServer("
    
'EVAPD
    .Replace What:="evapd(", Replacement:="EPMModelCubeDesc("

'EVAPP
    .Replace What:="evapp(", Replacement:="EPMModelCubeID("

'EVMBR
    .Replace What:="evmbr(", Replacement:="EPMSelectMember("

'EVAST
    .Replace What:="evAST(", Replacement:="EPMeNVdATABASEID("

'EVASD
    .Replace What:="EVASD(", Replacement:="EPMeNVdATABASEDesc("

'EVCGT
    .Replace What:="EVCGT(", Replacement:="EPMCommentFullContext("

'EVDIM
    .Replace What:="EVDIM(", Replacement:="EPMDimensionType("
    
'EVRTI(
    .Replace What:="EVRTI(", Replacement:="EPMRefreshTime("

End With

End Sub
