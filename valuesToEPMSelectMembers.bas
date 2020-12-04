'This will not work in Non-English versions of BPC.
'It is intended to turn a list of accounts/dates/companies into EPMSelectMembers
'to make it easier for customers/clients to pick which dimension should be where 
'in a report. Use at your own risk.



Sub turnToSelections()
Const theFormulaToReplace = "=EPMSelectMember(""zzTheModelzz"",""zztheMemberzz"")"
Dim theErrorText As String


Dim theModel_ID As String
    theModel_ID = InputBox("Enter the MODEL ID to be used", "EPMSelectMember Conversion")
    
    If theModel_ID = "" Then Exit Sub

Dim theRange As Range
On Error GoTo endOfsub
    Set theRange = Application.InputBox( _
      Title:="Turn Cells To EPMSelections", _
      Prompt:="Select a range of cells to convert to EPMSelections." & _
        Chr(10) & "This Macro will skip any values that are not recognized members.", _
      Type:=8)
      On Error GoTo 0
      
Dim badValue As String
    badValue = vbaDesc("?stevenRider?")

Dim theRay(), i As Long
    theRay = theRange.Value
    For i = LBound(theRay) To UBound(theRay)
        For j = LBound(theRay, 2) To UBound(theRay, 2)
        If Len(theRay(i, j)) > 0 Then
            theRay(i, j) = Replace(Replace(theFormulaToReplace, "zzTheModelzz", theModel_ID), "zztheMemberzz", theRay(i, j))
        End If
        Next j
    Next i
    
    theRange.Formula = theRay

endOfsub:

End Sub

Private Function vbaDesc(theValue As String, Optional theModelID As String) As String
  Evaluate ("=EPMMEMERDESC(""" & theValue & """" & IIf(theModelID = "", ")", ",""" & theModelID & """)"))
End Function
