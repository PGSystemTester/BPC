Const questionforReduction As String = "The row axis has dimensions with a single member in it. Would you like to move these members to the column axis?"
Const sName As String = "NumberHunter_"
Const showGridelines As Boolean = False
Const noFormatSheet As String = "No Format Sheet"
Const setFreezePanes As Boolean = True
Const showZeroIntersections As Boolean = True


Dim EA As Object
Dim RPT_ID As String, ws As Worksheet, StopEverything As Boolean, LocalMemberFound As Range
Dim RowMembers() As String, ColMembers() As String, PageMembers() As String
Dim AllDimensions() As i_DIMENSION

Type i_DIMENSION
    i_DimName As String
    i_DimID As String
    i_isCalc As Boolean
    i_Type As String * 1
    i_inREPORT As Boolean
    i_DisplayType As displyTYPE

End Type

Enum displyTYPE
    ID_ONly = 0
    Desc_Only = 1
    Both_ID_DESC = 2
End Enum


Sub NumberHunter()
Set EA = Nothing
Call Full_Number_Hunter(ActiveCell)
Set EA = Nothing

End Sub

Sub addRemoveBotton(adIsTrue As Boolean)
Const zCaption = "Number Hunter"
Const zMacro = "NumberHunter"
Dim iconExist As Boolean

'must have this code in workbook module
    'Private Sub Workbook_Activate()
    '    Call addRemoveBotton(True)
    'End Sub
    '
    'Private Sub Workbook_Deactivate()
    '    Call addRemoveBotton(False)
    'End Sub


iconExist = testIfIconExists(zCaption)

If adIsTrue Then
    If testIfIconExists("EPM") And Not iconExist Then
    
        Dim cmdBtn As CommandBarButton
       Set cmdBtn = Application.CommandBars("Cell").Controls.Add(Temporary:=True)
        
        With cmdBtn
            .Caption = zCaption
           .Style = msoButtonCaption
           .OnAction = zMacro
        End With
    End If
ElseIf iconExist Then

    Call clearButtons(zCaption)
    
End If

     
End Sub


Private Function testIfIconExists(theName) As Boolean

For i = 1 To Application.CommandBars("Cell").Controls.Count
    
    With Application.CommandBars("Cell").Controls.Item(i)
    
        If .Caption = theName Then
            testIfIconExists = .Visible
            Exit Function
        End If
    End With
Next i

End Function




Private Sub clearButtons(theName As String)

Dim i As Long

For i = Application.CommandBars("Cell").Controls.Count To 1 Step -1
    
    With Application.CommandBars("Cell").Controls.Item(i)
    
        If .Caption = theName Then
            .Delete
           
        End If
    End With
    
Next i
End Sub





Private Sub Full_Number_Hunter(rCell As Range)
Set ws = rCell.Worksheet

'get report ID
    RPT_ID = getReportID(rCell)

'Test if valid ID
    If StopEverything Then GoTo endTHIS

'Find Row key members
    RowMembers = Split(getROWmembers(rCell), ",")
    If StopEverything Then GoTo endTHIS
    
'Find Column key members
    ColMembers = Split(getCOLmembers(rCell), ",")
    If StopEverything Then GoTo endTHIS
    
'Find Page Axis Members
    PageMembers = EA.GetPageAxisMembers(ws, RPT_ID)

'Put dimensions into collection of dims array
    Dim pullDIMS() As String
    pullDIMS = EA.GetDimensionList(EA.GetActiveConnection(ws))
    
    ReDim AllDimensions(UBound(pullDIMS)) As i_DIMENSION
    
    Dim d As Long
    For d = 0 To UBound(AllDimensions)
        With AllDimensions(d)
            .i_DimName = pullDIMS(d)
            .i_DimID = igetMemberfromDIM(.i_DimName)
            .i_isCalc = Evaluate("=EPMMEMBERPROPERTY(,""" & .i_DimID & """,""CALC"")") = "Y"
            .i_Type = TypeOFDim(.i_DimName)
            
        End With
        
    Next d

    Dim nws As Worksheet
    ActiveWorkbook.Sheets.Add
        Set nws = ActiveSheet
        
        'renames sheet
        Call nameNewSheet(nws)
        
        ActiveWindow.DisplayGridlines = showGridelines

        EA.CreateReport ActiveSheet, EA.GetActiveConnection(ws), "000", _
            ReturnDimMember("R"), 1, _
            ReturnDimMember("A"), 1, Range("a1")
            DoEvents
            

        For d = 0 To UBound(AllDimensions)
            If AllDimensions(d).i_Type = "R" Then
                'skip
            ElseIf AllDimensions(d).i_Type = "A" Then
                EA.AddMemberToRowAxis nws, "000", AllDimensions(d).i_DimID, 6
                EA.RemoveMemberFromRowAxis nws, "000", AllDimensions(d).i_DimID, 1
                
            ElseIf AllDimensions(d).i_DimName = "MEASURES" Then
                EA.AddMemberToColumAxis nws, "000", ReturnDimMember("MEASURES"), 1
            
            ElseIf AllDimensions(d).i_isCalc = False Then
                'put in Column axis
                EA.AddMemberToColumAxis nws, "000", AllDimensions(d).i_DimID, 1
                
            Else
                'put in rows
                EA.AddMemberToRowAxis nws, "000", AllDimensions(d).i_DimID, 6
            End If
            DoEvents
        Next d
                
       If showZeroIntersections Then
                    EA.SetSheetOption nws, 7, True
                Else
                    EA.SetSheetOption nws, 7, True
        End If
                    
        
        EA.SetSheetOption nws, 111, True 'Clear formatting
        
        
        If getFormattingSheetName <> noFormatSheet Then
            EA.SetSheetOption nws, 110, getFormattingSheetName() 'Apply Formatting
        End If
        

        EA.SetSheetOption nws, 100, True 'Show row header
        EA.RefreshActiveSheet
        
        'check for single member rows.
        Call ClearSingleDimRows
        
        Call ReducereportSize(nws)


        Range(EA.GetDataTopLeftCell(nws, "000")).Select
        
        'set freeze panes
        If setFreezePanes Then ActiveWindow.FreezePanes = True


endTHIS:
If Not LocalMemberFound Is Nothing Then
    Call RejectionOfLocalMember
End If
RPT_ID = ""
Set ws = Nothing
Set EA = Nothing
StopEverything = False


End Sub

Private Function getReportID(rCell As Range) As String
 If EA Is Nothing Then Set EA = ntt_BPC_API
    Dim rptNAmes() As String, r As Long
    
    'captures all reports on sheet
    rptNAmes = EA.GetAllReportNames(rCell.Worksheet)
    
    'tests if intersection exists in any reports
    For r = 0 To UBound(rptNAmes)
        
        If Not Intersect(rCell, Range(Range(EA.GetDataTopLeftCell(ws, rptNAmes(r))), _
            Range(EA.GetDataBottomRightCell(ws, rptNAmes(r))))) Is Nothing Then
                
                'intersection found
                getReportID = rptNAmes(r)
                Exit Function
        
        End If
    Next r
    
    StopEverything = True


End Function




Private Function getROWmembers(rCell As Range) As String
Dim c As Long, allRowMembers() As String, theDim As Range, memberID As String, axis_RPT_ID As String

If EA Is Nothing Then Set EA = ntt_BPC_API

axis_RPT_ID = EA.GetRowAxisOwner(ws, RPT_ID)
If axis_RPT_ID = "" Then axis_RPT_ID = RPT_ID


allRowMembers = EA.GetRowAxisMembers(ws, RPT_ID)


    Dim columnX As Long: columnX = FirstRowAxisNumber(allRowMembers(0))


    If StopEverything Then Exit Function


'Loop through all members in Row Axis
    For c = columnX To columnX + EA.GetRowAxisDimensionCount(rCell.Worksheet, RPT_ID) - 1
        Set theDim = ws.Cells(rCell.Row, c)
        

        Do Until Not IsEmpty(theDim)
            Set theDim = theDim.Offset(-1, 0)
        Loop
        
        'capture memberID
        memberID = Evaluate("=EPMMemberID(" & theDim.Address & ")")
       
        
        'Test if LOCAL MEMBER or something else
        If memberID = Evaluate("=EPMMemberID(xfa999999)") Then
            Set LocalMemberFound = theDim
            StopEverything = True
            getROWmembers = "end"
            Exit For
        
        
        Else
            'defense code
           ' StopEverything = Evaluate("=Iserror(Colerror)")

        End If
        
        'String together results
        getROWmembers = getROWmembers & memberID & ","
    
    Next c

'gets rid of last comma
getROWmembers = Mid(getROWmembers, 1, Len(getROWmembers) - 1)
End Function


Private Function FirstRowAxisNumber(memberID As String) As Long
If EA Is Nothing Then Set EA = ntt_BPC_API

Dim axisRPT_ID As String

'confirm the row axis is not shared
axisRPT_ID = EA.GetRowAxisOwner(ws, RPT_ID)

If axisRPT_ID = "" Then axisRPT_ID = RPT_ID

'find upper left cell
Dim uCell As Range
Set uCell = ws.Range(EA.GetDataTopLeftCell(ws, axisRPT_ID)).Offset(0, -1)

'find first value that is closet to left
Dim goThisWay As Long, totalMembers As Long
goThisWay = -1

lookForRow:

Do Until InStr(2, uCell.Formula, "EPMlocal", vbTextCompare) > 0 Or InStr(2, uCell.Formula, "EPMOlapMemberO", vbTextCompare) > 0 And _
    InStr(2, uCell.Formula, axisRPT_ID, vbTextCompare) > 0

Set uCell = uCell.Offset(0, goThisWay)

If uCell.Column = 1 Then
    goThisWay = 1
    Set uCell = Cells(uCell.Row, Range(EA.GetDataBottomRightCell(ws, axisRPT_ID)).Column).Offset(0, goThisWay)
    GoTo lookForRow
End If
    
Loop

totalMembers = EA.GetRowAxisDimensionCount(ws, axisRPT_ID) - 1

'if on the left...
    If goThisWay = -1 Then Set uCell = uCell.Offset(0, totalMembers * -1)
    

FirstRowAxisNumber = uCell.Column
End Function


Private Sub RejectionOfLocalMember()
Const theTitle As String = "Not a valid intersection"

MsgBox "Your intersection includes a local member in cell " & LocalMemberFound.Address & " """ & LocalMemberFound.Value & """", vbCritical, theTitle

Set LocalMemberFound = Nothing

End Sub


Private Function getCOLmembers(rCell As Range) As String
Dim c As Long, allCOLMembers() As String, theDim As Range, memberID As String

If EA Is Nothing Then Set EA = ntt_BPC_API

allCOLMembers = EA.GetColumnAxisMembers(ws, RPT_ID)

'Find Starting row
    Dim rowX As Long: rowX = FirstColAxisNumber(allCOLMembers(0))
    If StopEverything Then Exit Function


'Loop through all members in Column Axis
    For c = rowX To rowX + EA.GetColumnAxisDimensionCount(rCell.Worksheet, RPT_ID) - 1
        Set theDim = ws.Cells(c, rCell.Column)

        Do Until Not IsEmpty(theDim)
            Set theDim = theDim.Offset(0, -1)
        Loop
        
        'capture memberID
        memberID = Evaluate("=EPMMemberID(" & theDim.Address & ")")
        
        'Test if LOCAL MEMBER or something else
        If memberID = Evaluate("=EPMMemberID(xfa999999)") Then
            Set LocalMemberFound = theDim
            StopEverything = True
            getCOLmembers = "end"
            Exit For
        End If
        
        
        'String together results
        getCOLmembers = getCOLmembers & memberID & ","
    
    Next c

'gets rid of last comma
getCOLmembers = Mid(getCOLmembers, 1, Len(getCOLmembers) - 1)

End Function

Private Function FirstColAxisNumber(memberID As String) As Long
If EA Is Nothing Then Set EA = ntt_BPC_API

Dim axisRPT_ID As String

'confirm the row axis is not shared
axisRPT_ID = EA.GetColumnAxisOwner(ws, RPT_ID)

If axisRPT_ID = "" Then axisRPT_ID = RPT_ID

'find first top cell
Dim uCell As Range
Set uCell = ws.Range(EA.GetDataTopLeftCell(ws, axisRPT_ID)).Offset(-1, 0)

'find first column axis member
Dim totalMembers As Long

Do Until InStr(2, uCell.Formula, "EPMlocal", vbTextCompare) > 0 Or InStr(2, uCell.Formula, "EPMOlapMemberO", vbTextCompare) > 0 And _
    InStr(2, uCell.Formula, axisRPT_ID, vbTextCompare) > 0

    Set uCell = uCell.Offset(-1, 0)

Loop

totalMembers = EA.GetColumnAxisDimensionCount(ws, axisRPT_ID) - 1

FirstColAxisNumber = uCell.Offset(totalMembers * -1, 0).Row
    
End Function

Private Function igetMemberfromDIM(theDim As String) As String
If EA Is Nothing Then Set EA = ntt_BPC_API
Dim d As Long

'Check Row Access
For d = 0 To UBound(RowMembers)
    If EA.GetMemberDimension(EA.GetActiveConnection(ws), RowMembers(d)) = theDim Then
        igetMemberfromDIM = RowMembers(d)
        Exit Function
    End If
Next d

For d = 0 To UBound(ColMembers)
    If EA.GetMemberDimension(EA.GetActiveConnection(ws), ColMembers(d)) = theDim Then
        igetMemberfromDIM = ColMembers(d)
        Exit Function
    End If
Next d

For d = 0 To UBound(PageMembers)
    If EA.GetMemberDimension(EA.GetActiveConnection(ws), PageMembers(d)) = theDim Then
        igetMemberfromDIM = PageMembers(d)
        Exit Function
    End If
Next d

igetMemberfromDIM = Evaluate("=EPMCONTEXTMEMBER(,""" & theDim & """)")


End Function


Private Function TypeOFDim(theDim As String) As String
Dim c As Long
    For c = 65 To 90
    
        If Evaluate("=EPMDimensionType(,""" & Chr(c) & """)") = theDim Then
            TypeOFDim = Chr(c)
            Exit Function
        End If
        
    Next c
    
    TypeOFDim = "U"

End Function


Private Function ReturnDimMember(dimNAME As String) As String
dimNAME = UCase(dimNAME)
Dim d As Long
For d = 0 To UBound(AllDimensions)
    If UCase(AllDimensions(d).i_DimName) = dimNAME Then
        ReturnDimMember = AllDimensions(d).i_DimID
        Exit Function
    End If
Next d


'check for dimenion type (Allows users to just put in "A"
For d = 0 To UBound(AllDimensions)
    If UCase(AllDimensions(d).i_Type) = dimNAME Then
        ReturnDimMember = AllDimensions(d).i_DimID
        Exit Function
    End If
Next d



End Function

Private Function getFormattingSheetName() As String
Dim xws As Worksheet
For Each xws In ActiveWorkbook.Worksheets

    If xws.Range("B5").Value = "Hierarchy Level Formatting" And xws.Range("B1").Value = "EPM Formatting Sheet" Then
        getFormattingSheetName = xws.Name
        Exit Function
    End If
Next xws

getFormattingSheetName = noFormatSheet


End Function

Private Sub ReducereportSize(aWS As Worksheet)
Dim d As Long

        'delete unused rows
        With aWS
        d = .Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlNext).Row
            If d > 1 Then Range(.Rows(1), .Rows(d - 1)).Delete
           
        'delete unused Columns
        d = .Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
            If d > 1 Then Range(.Columns(1), .Columns(d - 1)).Delete

        End With


End Sub

Private Sub ClearSingleDimRows()
If EA Is Nothing Then Set EA = ntt_BPC_API
Dim c As Long, fullDimText() As String, tangoFIND As Range, UserWantsReduction As Boolean

'does not run on single member
If EA.GetRowAxisDimensionCount(ActiveSheet, "000") = 1 Then Exit Sub


fullDimText = EA.GetRowAxisMembers(ActiveSheet, "000")

'eliminates members not needed
ReDim Preserve fullDimText(EA.GetRowAxisDimensionCount(ActiveSheet, "000") - 1)

'find first instance
    Set tangoFIND = ActiveSheet.Cells.Find(fullDimText(0), _
        LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)


'Loop through first row of dims
Dim cRng As Range, g As Long
For c = LBound(fullDimText) To UBound(fullDimText)
    
    If g + 2 = ActiveSheet.UsedRange.Columns.Count Then Exit For


    Set cRng = Intersect(Range(tangoFIND, Cells(Rows.Count, tangoFIND.Column)), ActiveSheet.UsedRange)
    
    If Application.WorksheetFunction.CountA(cRng) = Application.WorksheetFunction.CountIf(cRng, tangoFIND.Value2) Then
        g = g + 1
        
        UserWantsReduction = True 'assume this always true unless someone asks....
        If UserWantsReduction = False Then
            Dim answer As Long
            
            answer = MsgBox(questionforReduction, vbQuestion + vbYesNo)
            If answer = vbYes Then
                UserWantsReduction = True
            Else
                Exit Sub
            End If
            
        End If
        
        'remove member from row and insert to column
            Dim dimToRemove As String
            dimToRemove = EA.GetMemberDimension(EA.GetActiveConnection(ActiveSheet), Evaluate("=EPMMEMBERID(" & tangoFIND.Address & ")"))
            dimToRemove = ReturnDimMember(dimToRemove)
            
            
            On Error GoTo nodata
            EA.RemoveMemberFromRowAxis ActiveSheet, "000", dimToRemove, 6
            EA.AddMemberToColumnAxis ActiveSheet, "000", Evaluate("=epmMemberID(" & tangoFIND.Address & ")"), 1
    
    
    End If
    Set tangoFIND = tangoFIND.Offset(0, 1)
Next c

    If UserWantsReduction Then EA.RefreshActiveSheet
nodata:

End Sub

Private Sub nameNewSheet(theWS As Worksheet)
Dim sCount As Long
Dim newName As String
sCount = 0
  
  
runCheck:
newName = sName & Application.WorksheetFunction.Base(sCount, 10, 3)

Dim aSh As Worksheet
For Each aSh In ActiveWorkbook.Worksheets
    If aSh.Name = newName Then
        sCount = sCount + 1
        GoTo runCheck
    End If
Next aSh

theWS.Name = newName

End Sub

Private Function ntt_BPC_API() As Object
    Const NoConnectMessage As String = "No Connection Found"
    Dim aoComAdd As Object, successConnection As Boolean, ObjAddOn As COMAddIn
 
        For Each ObjAddOn In Application.COMAddIns
            If ObjAddOn.progID = "FPMXLClient.Connect" Then
                'EPM/BPC
               Set ntt_BPC_API = ObjAddOn.Object
                successConnection = True
                Exit For
            ElseIf ObjAddOn.progID = "SapExcelAddIn" Then
                'Analysis for Office Version
               Set aoComAdd = ObjAddOn.Object
                Set ntt_BPC_API = aoComAdd.GetPlugin("com.sap.epm.FPMXLClient")
                successConnection = True
                Exit For
            End If
        Next ObjAddOn
     
        If Not successConnection Then
            MsgBox NoConnectMessage
            End
        End If
   
End Function
