<%
'******************************************************************
Function IIf( expr, truepart, falsepart )
   IIf = falsepart
   If expr Then IIf = truepart
End Function


' Returns Null for blank/whitespace; otherwise returns a string
Function NullIfEmpty(ByVal s)
  s = Trim("" & s)
  If Len(s) = 0 Then
    NullIfEmpty = Null
  Else
    NullIfEmpty = CStr(s)
  End If
End Function

' Returns Null for blank; otherwise returns a Double.
' Strips $ and commas first.
Function NullIfBlankNumber(ByVal s)
  s = Trim("" & s)
  s = Replace(s, "$", "")
  s = Replace(s, ",", "")
  If Len(s) = 0 Then
    NullIfBlankNumber = Null
  ElseIf IsNumeric(s) Then
    NullIfBlankNumber = CDbl(s)
  Else
    ' choose one:
    ' 1) return Null silently:
    NullIfBlankNumber = Null
    ' 2) OR raise an error:
    ' Err.Raise vbObjectError + 1001, "NullIfBlankNumber", "Invalid number: " & s
  End If
End Function

' Handy for adLongVarWChar input params: ADO wants a real Size (not -1)
Function SizeOrOne(ByVal s)
  s = "" & s
  If Len(s) = 0 Then SizeOrOne = 1 Else SizeOrOne = Len(s)
End Function



%>