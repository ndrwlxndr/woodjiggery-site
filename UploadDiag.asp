<%
Response.Buffer = True
Dim total, bin
total = Request.TotalBytes
If total > 0 Then
    On Error Resume Next
    bin = Request.BinaryRead(total)
    If Err.Number <> 0 Then
        Response.Write "<p style='color:red'>BinaryRead FAILED: " & Err.Number & " - " & Err.Description & "</p>"
        Response.Write "<p>Content-Type: " & Request.ServerVariables("CONTENT_TYPE") & "</p>"
        Response.Write "<p>Request Method: " & Request.ServerVariables("REQUEST_METHOD") & "</p>"
    Else
        Response.Write "<p style='color:green'>BinaryRead OK - got " & LenB(bin) & " bytes</p>"
    End If
    On Error GoTo 0
Else
    Response.Write "<p>No bytes received (GET request or empty POST)</p>"
End If
%>
<form method="post" enctype="multipart/form-data">
  <input type="file" name="f" />
  <input type="submit" value="Test BinaryRead" />
</form>
