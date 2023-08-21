Option Explicit

' _____________________________________________________________
' Purpose: Find a substring within a given string matching to a
'          corresponding regex pattern, e.g. ABC123, AB1234
' Input:   Sheet1 contains cells with strings in column 1.
           The term and the title are separated by a comma.
' _____________________________________________________________

Sub split_id_title()

Dim lngRow As Long, lngRowMax As Long
Dim strLine As String, strTerm As String, strTitle As String
Dim varArray As Variant

With Sheet1
  lngRowMax = .UsedRange.Rows.Count
  For lngRow = 2 To lngRowMax
    strLine = .Cells(lngRow, 1).Value
    strTerm = get_term(strLine)
    varArray = Split(strLine, ",")
    strTitle = varArray(1)
    .Cells(lngRow, 2).Value = strTerm
    .Cells(lngRow, 3).Value = strTitle
  Next lngRow
End With

End Sub

' _____________________________________________________________
Function get_term(ByVal strLine As String) As String

Dim objRegEx As Object
Dim objMatch As Object

Set objRegEx = CreateObject("vbscript.regexp")

With objRegEx
  .Global = True
  ' Pattern like: ABC123 | AB1234
  .Pattern = "(A)[A-Z]{2}[0-9]{3}|(A)[A-Z][0-9]{4}"
  Set objMatch = .Execute(strLine)
End With

If objMatch.Count = 1 Then
  get_term = objMatch(0)
ElseIf objMatch.Count = 2 Then ' Pattern occurs twice in a line.
  get_term = objMatch(1)
Else
  get_term = "N/A"
End If

End Function
