Option Explicit

Sub find_regex_in_string()

Dim lngRow, lngRowMax As Long
Dim lngColToSearch As Long
Dim strContent As String
Dim wksSheet as Worksheet

Set wksSheet = Sheet1
lngColToSearch = 1

With Sheet1

  lngRowMax = .Cells(.Rows.Count, lngColToSearch).End(xlUp).Row
  For lngRow = lngRowMax To 2 Step -1
    strContent = .Cells(lngRow, lngColToSearch).Value
    If isRegEx(strContent) = False Then
      .Rows(lngRow).Delete
    End If
  Next lngRow
  
End With

End Sub

' Regular Expressions are usually defined in a function:

Function isRegEx(strContent As String) As Boolean

Dim objRegEx, objMatch As Object

Set objRegEx = CreateObject("vbscript.regexp")

With objRegEx
  .Global = True
  .Pattern = "(A)[0-9][0-9]"
  Set objMatch = .Execute(strContent)
End With

If objMatch.Count = 1 Then
  isRegEx = True
Else
  isRegEx = False
End If

End Function
