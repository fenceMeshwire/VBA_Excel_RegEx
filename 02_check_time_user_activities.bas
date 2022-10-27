Option Explicit

' Existing structure:
' user | time | count
' Added column:
' user | time | count | hour
' The added column countains the hour of time from the column "time"

' ============================================================
Sub add_column_hour()

Dim lngRow, lngRowMax As Long
Dim strHour, strTime As String
Dim wksSheet As Worksheet

Set wksSheet = Sheet1

With wksSheet
  lngRowMax = .UsedRange.Rows.Count
  For lngRow = 2 To lngRowMax
    strTime = .Cells(lngRow, 2).Value
    strHour = Left(strTime, 2)
    strHour = check_hour(strHour)
    .Cells(lngRow, 4).Value = strHour
  Next lngRow
End With

End Sub

' ============================================================
Function check_hour(strHour) As String

Dim objRegEx, objMatch As Object

Set objRegEx = CreateObject("vbscript.regexp")

With objRegEx
  .Global = True
  .Pattern = "0[0-9]:[0-5][0-9]"
  Set objMatch = .Execute(strHour)
End With

If objMatch.Count = 1 Then check_hour = Right(strHour, 1) Else check_hour = strHour

End Function

' ============================================================
Sub cumsum_hour_usage()

Dim i As Integer
Dim intHour As Integer
Dim lngRow, lngRowMax As Long
Dim strHour, strTime As String
Dim lngUsage As Long
Dim wksSheet As Worksheet
Dim varHours As Variant

Set wksSheet = Sheet1

With wksSheet
            
  ' Create hours array with no duplicates
  ReDim varHours(i)
  lngRowMax = .UsedRange.Rows.Count
  For lngRow = 2 To lngRowMax
    strHour = .Cells(lngRow, 4).Value
    If Not IsNumeric(Application.Match(strHour, varHours, 0)) Then ' No duplicates
      varHours(i) = strHour
      i = i + 1
      ReDim Preserve varHours(i)
    End If
  Next lngRow
  ReDim Preserve varHours(UBound(varHours) - 1)
  
  ' Accumulate user activities
  For i = LBound(varHours) To UBound(varHours)
    intHour = varHours(i)
    lngUsage = 0
    For lngRow = 2 To lngRowMax
      If .Cells(lngRow, 4).Value = intHour Then
        lngUsage = lngUsage + .Cells(lngRow, 3).Value
        If .Cells(lngRow + 1, 4).Value <> intHour Then
          .Cells(lngRow, 5).Value = lngUsage
        End If
      End If
    Next lngRow
  Next i
  
  ' Delete rows with no accumulated user activities for summary
  For lngRow = lngRowMax To 2 Step -1
    If .Cells(lngRow, 5).Value = "" Then .Rows(lngRow).Delete
  Next lngRow
  
End With

End Sub
