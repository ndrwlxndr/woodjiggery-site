<%
Dim path : path = LCase(Request.ServerVariables("SCRIPT_NAME"))

If InStr(path, "/admin/login.asp") = 0 Then
  If Session("IsAdmin") <> True Then
    Dim ret : ret = Server.URLEncode(Request.ServerVariables("URL"))
    If Len(Request.ServerVariables("QUERY_STRING")) > 0 Then
      ret = Server.URLEncode(Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING"))
    End If
    Response.Redirect "login.asp?return=" & ret
  End If
End If
%>