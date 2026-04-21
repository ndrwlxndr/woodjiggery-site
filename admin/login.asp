
<%
'Option Explicit
Response.Buffer = True

ADMIN_USER = "admin"
ADMIN_PASS = "temp" ' CHANGE THIS!


Dim errMsg : errMsg = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
  Dim u, p
  u = Trim(Request.Form("username"))
  p = Trim(Request.Form("password"))

  If LCase(u) = LCase(ADMIN_USER) And p = ADMIN_PASS Then
    Session("IsAdmin") = True

    Dim ret : ret = Trim(Request.QueryString("return"))
    If Len(ret) > 0 Then
      Response.Redirect ret
    Else
      Response.Redirect "default.asp"
    End If
  Else
    errMsg = "Invalid username/password."
  End If
End If
%>
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>Admin Login</title>
  <style>
    body{font-family:Segoe UI,Arial; padding:24px;}
    .card{max-width:420px; border:1px solid #ddd; border-radius:10px; padding:18px;}
    label{display:block; margin-top:10px;}
    input{width:100%; padding:10px; box-sizing:border-box;}
    button{margin-top:14px; padding:10px 14px;}
    .err{color:#b00020; margin-top:10px;}
  </style>
</head>
<body>
  <div class="card">
    <h2>Admin Login</h2>
    <% If Len(errMsg)>0 Then %><div class="err"><%=Server.HTMLEncode(errMsg)%></div><% End If %>
    <form method="post" action="login.asp?<%=Server.HTMLEncode(Request.QueryString)%>">
      <label>Username</label>
      <input name="username" autocomplete="username" />
      <label>Password</label>
      <input name="password" type="password" autocomplete="current-password" />
      <button type="submit">Log in</button>
    </form>
  </div>
</body>
</html>