option explicit
'Code modified from
'https://blogs.sap.com/2019/06/24/bpc-nw-10-vba-function-to-get-bassomeparent-dimension-members-list/

Public Function getBaseMembers(theDim As String, theParent As String)
  Const strSplitter As String = ";:"
  Dim tempArray() As String
      tempArray = GetBAS(theDim, theParent, ActiveSheet)
      
  If testForSpill Then
      getBaseMembers = tempArray
  Else
      getBaseMembers = Join(tempArray, strSplitter)
  End If


End Function

Private Function GetBAS(strDim As String, strParentMember As String, Optional WS As Worksheet) As String()
    Dim objAddIn As COMAddIn, epm As Object, aoComAdd As Object
    
    Dim strDims() As String, strProps() As String, strMem() As String, _
        strMemIDParent() As String, strMemBAS() As String
    
    Dim blnExistFlag As Boolean, blnFormulaExistFlag As Boolean, _
        blnIsNotFormulaFlag As Boolean, blnEPMInstalled As Boolean, blnIsCalcFlag As Boolean
        
    Dim strCalcProp As String, strParentMemberDim As String, _
        strConn As String, strParentHierarchy As String, hierarchyName As String

    Dim lngTemp As Long, lngBASCount As Long
    
    If WS Is Nothing Then Set WS = ActiveSheet
    
On Error GoTo Err
    'Universal code to get FPMXLClient for standalone EPM or AO
    For Each objAddIn In Application.COMAddIns
        If objAddIn.progID = "FPMXLClient.Connect" Then
            Set epm = objAddIn.Object
            blnEPMInstalled = True
            Exit For
        ElseIf objAddIn.progID = "SapExcelAddIn" Then
            Set aoComAdd = objAddIn.Object
            Set epm = aoComAdd.GetPlugin("com.sap.epm.FPMXLClient")
            blnEPMInstalled = True
            Exit For
        End If
    Next objAddIn
    
    If Not blnEPMInstalled Then
        ReDim strMemBAS(0 To 1)
        strMemBAS(0) = ""
        strMemBAS(1) = "NO_EPM"
        GetBAS = strMemBAS
        Exit Function
    End If
    
    strConn = epm.getactiveconnection(WS)
    
    'Check if Dimension strDim exists
    strDims = epm.GetDimensionList(strConn)
    For lngTemp = 0 To UBound(strDims)
        If UCase(strDims(lngTemp)) = UCase(strDim) Then
            blnExistFlag = True
            strDim = strDims(lngTemp)
        End If
    Next lngTemp
    
    Erase strDims
    
    If Not blnExistFlag Then
        ReDim strMemBAS(0 To 1)
        strMemBAS(0) = ""
        strMemBAS(1) = "NO_DIMENSION"
        GetBAS = strMemBAS
        Exit Function
    End If
    
    'Check if Dimension strDim has one or more hierarchies and contain FORMULA property
    blnExistFlag = False
    strProps = epm.GetPropertyList(strConn, strDim)
    
        hierarchyName = IIf(testBpcIsMicrosoft(strConn), "H1", "PARENTH1")
    
    For lngTemp = 0 To UBound(strProps)
        If strProps(lngTemp) = hierarchyName Then
            blnExistFlag = True
            If blnFormulaExistFlag Then Exit For
        ElseIf strProps(lngTemp) = "FORMULA" Then
            blnFormulaExistFlag = True
            If blnExistFlag Then Exit For
        End If
    Next lngTemp
    
    Erase strProps
    
    If Not blnExistFlag Then
        'No hierarchy
        'Check that member exists in dimension
        strMem = epm.GetHierarchyMembers(strConn, "", strDim)
        For lngTemp = 0 To UBound(strMem)
            If UCase(Application.Run("EPMMemberProperty", "", strMem(lngTemp), "ID")) = UCase(strParentMember) Then
                strParentMember = Application.Run("EPMMemberProperty", "", strMem(lngTemp), "ID")
                GoTo MEMBER_ITSELF
            End If
            
        Next lngTemp
        GoTo NO_MEMBER
    End If
    
    strParentMemberDim = strDim & ":" & strParentMember
    
    strCalcProp = Application.Run("EPMMemberProperty", "", strParentMemberDim, "CALC")
    If strCalcProp = "Y" Then
        blnIsCalcFlag = True
    ElseIf strCalcProp Like "[#]Error - Invalid Member Name:*" Then GoTo NO_MEMBER
    End If
    
    
    'If we have dimension member formulas - check for formula of strParentMember
    blnIsNotFormulaFlag = True
    If blnFormulaExistFlag Then
        If Application.Run("EPMMemberProperty", "", strParentMemberDim, "FORMULA") <> "" Then
            blnIsNotFormulaFlag = False
        End If
    End If

    If blnIsCalcFlag And blnIsNotFormulaFlag Then
        strParentHierarchy = epm.GetMemberHierarchy(strConn, strParentMemberDim)
        strMem = epm.GetHierarchyMembers(strConn, strParentHierarchy, strDim)
        ReDim strMemIDParent(0 To 1, 0 To UBound(strMem))
        ReDim strMemBAS(0 To UBound(strMem))
        blnExistFlag = False
        For lngTemp = 0 To UBound(strMem)
            strMemIDParent(0, lngTemp) = Application.Run("EPMMemberProperty", "", strMem(lngTemp), "ID")
            strMemIDParent(1, lngTemp) = Application.Run("EPMMemberProperty", "", strMem(lngTemp), strParentHierarchy)
            If strMemIDParent(0, lngTemp) = strParentMember Then
                blnExistFlag = True
            End If
        Next lngTemp
        If Not blnExistFlag Then GoTo NO_MEMBER
        GetChildren strParentMember, strMemIDParent, strMemBAS, lngBASCount
        ReDim Preserve strMemBAS(0 To lngBASCount - 1)
        GetBAS = strMemBAS
    Else
        'Check that member exists in dimension
        strMem = epm.GetHierarchyMembers(strConn, "", strDim)
        For lngTemp = 0 To UBound(strMem)
            If UCase(Application.Run("EPMMemberProperty", "", strMem(lngTemp), "ID")) = UCase(strParentMember) Then
                strParentMember = Application.Run("EPMMemberProperty", "", strMem(lngTemp), "ID")
                GoTo MEMBER_ITSELF
            End If
            
        Next lngTemp
        GoTo NO_MEMBER
    End If
    
    Exit Function

MEMBER_ITSELF:
    'Member found
    ReDim strMemBAS(0 To 0)
    strMemBAS(0) = strParentMember
    GetBAS = strMemBAS
    Exit Function

NO_MEMBER:
    'Member not found
    ReDim strMemBAS(0 To 1)
    strMemBAS(0) = ""
    strMemBAS(1) = "NO_MEMBER"
    GetBAS = strMemBAS
    Exit Function

Err:
    ReDim strMemBAS(0 To 1)
    strMemBAS(0) = ""
    If Err.Number = -1073479167 Then
        strMemBAS(1) = "NO_CONNECTION"
    Else
        strMemBAS(1) = "OTHER_ERROR"
    End If
    GetBAS = strMem

End Function

Private Sub GetChildren(strParent As String, ByRef strMemIDParent() As String, _
    ByRef strMemBAS() As String, ByRef lngBASCount As Long)
    
    Dim lngTemp As Long
    Dim blnParent As Boolean
    
    For lngTemp = 0 To UBound(strMemIDParent, 2)
        If strMemIDParent(1, lngTemp) = strParent Then
            blnParent = True
            GetChildren strMemIDParent(0, lngTemp), strMemIDParent, strMemBAS, lngBASCount
        End If
    Next lngTemp
    If Not blnParent Then
        strMemBAS(lngBASCount) = strParent
        lngBASCount = lngBASCount + 1
    End If
End Sub

Private Function testForSpill() As Boolean
    'tests for excel Spill range functionality
    On Error GoTo nopE
        testForSpill = IsArray(Application.WorksheetFunction.Unique(Array("a")))
    On Error GoTo 0
    Exit Function
nopE:
End Function

Private Function testBpcIsMicrosoft(theConnection As String) As Boolean
    If InStr(1, Left(theConnction, InStr(1, theConnction, "[", vbTextCompare)), "FPM_BPCMS", vbTextCompare) > 0 Then
        testBpcIsMicrosoft = True
    End If

End Function
