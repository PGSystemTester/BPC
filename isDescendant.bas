Public Function isDescendant(memberBeingTested As String, parentMember As String, Optional hierarchyName As String) As Boolean
    Dim zDim As String, hLevel As Long, zHierachy As String
    
    zHierachy = IIf(hierarchyName = "", "PARENTH1", hierarchyName)
    
    zDim = memberBeingTested
    For hLevel = Evaluate("=EPMMEMBERPROPERTY(,""" & zDim & """,""HLEVEL"")") To 1 Step -1
        If UCase(zDim) = UCase(parentMember) Then
            isDescendant = True
            Exit For
        Else
            zDim = Evaluate("=EPMMEMBERPROPERTY(,""" & zDim & """,""" & zHierachy & """)")
        End If
        
    Next hLevel
End Function
