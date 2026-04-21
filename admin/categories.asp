<!--#include file="_auth.asp"-->
<!--#include virtual="includes/adovbs.asp"-->
<!--#include virtual="includes/datacon.asp"-->
<%
Response.Buffer = True

Function Slugify(ByVal s)
  Dim i, ch, out
  s = LCase(Trim(s))
  out = ""
  For i = 1 To Len(s)
    ch = Mid(s, i, 1)
    If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
      out = out & ch
    Else
      out = out & "-"
    End If
  Next
  Do While InStr(out, "--") > 0
    out = Replace(out, "--", "-")
  Loop
  out = Trim(out)
  If Left(out,1) = "-" Then out = Mid(out,2)
  If Right(out,1) = "-" Then out = Left(out, Len(out)-1)
  Slugify = out
End Function

On error resume next

Dim conn : Set conn = OpenConn()
OpenConn()



Dim action : action = LCase(Trim(Request("action")))
Dim msg : msg = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

  If action = "add" Then
    Dim name, slug, sortOrder, isActive
    name = Trim(Request.Form("Name"))
    slug = Trim(Request.Form("Slug"))
    sortOrder = Trim(Request.Form("SortOrder"))
    isActive = (Request.Form("IsActive") = "1")
 
    If Len(name) = 0 Then
      msg = "Name is required."
    Else
      If Len(slug) = 0 Then slug = Slugify(name)
      If Len(sortOrder) = 0 Then sortOrder = 0

      Dim cmdA : Set cmdA = Server.CreateObject("ADODB.Command")
      Set cmdA.ActiveConnection = conn
      cmdA.CommandType = adCmdText
      cmdA.CommandText = "INSERT INTO dbo.Categories (Name, Slug, SortOrder, IsActive) VALUES (?,?,?,?)"
      cmdA.Parameters.Append cmdA.CreateParameter("@Name", adVarWChar, adParamInput, 100, name)
      cmdA.Parameters.Append cmdA.CreateParameter("@Slug", adVarWChar, adParamInput, 100, slug)
      cmdA.Parameters.Append cmdA.CreateParameter("@Sort", adInteger,  adParamInput, , CInt(sortOrder))
      cmdA.Parameters.Append cmdA.CreateParameter("@Active", adBoolean, adParamInput, , CBool(isActive))
      cmdA.Execute
      Set cmdA = Nothing

      msg = "Category added."
    End If


ElseIf action = "update" Then
    Dim idU, nameU, slugU, sortU, activeU
    idU = CLng(Request.Form("CategoryID"))
    nameU = Trim(Request.Form("Name"))
    slugU = Trim(Request.Form("Slug"))
    sortU = Trim(Request.Form("SortOrder"))
    activeU = (Request.Form("IsActive") = "1")

    If Len(nameU) = 0 Then
      msg = "Name is required."
    Else
      If Len(slugU) = 0 Then slugU = Slugify(nameU)
      If Len(sortU) = 0 Then sortU = 0

      Dim cmdU : Set cmdU = Server.CreateObject("ADODB.Command")
      Set cmdU.ActiveConnection = conn
      cmdU.CommandType = adCmdText
      cmdU.CommandText = "UPDATE dbo.Categories SET Name=?, Slug=?, SortOrder=?, IsActive=? WHERE CategoryID=?"
      cmdU.Parameters.Append cmdU.CreateParameter("@Name", adVarWChar, adParamInput, 100, nameU)
      cmdU.Parameters.Append cmdU.CreateParameter("@Slug", adVarWChar, adParamInput, 100, slugU)
      cmdU.Parameters.Append cmdU.CreateParameter("@Sort", adInteger,  adParamInput, , CInt(sortU))
      cmdU.Parameters.Append cmdU.CreateParameter("@Active", adBoolean, adParamInput, , CBool(activeU))
      cmdU.Parameters.Append cmdU.CreateParameter("@ID", adInteger,   adParamInput, , CInt(idU))
      cmdU.Execute
      Set cmdU = Nothing

      msg = "Category updated."
    End If

  ElseIf action = "delete" Then
    Dim idD : idD = CLng(Request.Form("CategoryID"))

    Dim cmdD : Set cmdD = Server.CreateObject("ADODB.Command")
    Set cmdD.ActiveConnection = conn
    cmdD.CommandType = adCmdText
    cmdD.CommandText = "UPDATE dbo.Categories SET IsActive=0 WHERE CategoryID=?"
    cmdD.Parameters.Append cmdD.CreateParameter("@ID", adInteger, adParamInput, , CInt(idD))
    cmdD.Execute
    Set cmdD = Nothing

    msg = "Category deactivated."
  End If

End If

Dim editId : editId = 0

If Len(Request.QueryString("edit")) > 0 Then editId = CLng(Request.QueryString("edit"))

Dim editName, editSlug, editSort, editActive
editName = "" : editSlug = "" : editSort = 0 : editActive = True

If editId > 0 Then
  Dim cmdE : Set cmdE = Server.CreateObject("ADODB.Command")
  Set cmdE.ActiveConnection = conn
  cmdE.CommandType = adCmdText
  cmdE.CommandText = "SELECT CategoryID, Name, Slug, SortOrder, IsActive FROM dbo.Categories WHERE CategoryID=?"
  cmdE.Parameters.Append cmdE.CreateParameter("@ID", adInteger, adParamInput, , CInt(editId))

  Dim rsE : Set rsE = cmdE.Execute
  If Not rsE.EOF Then
    editName = rsE("Name")
    editSlug = rsE("Slug")
    editSort = rsE("SortOrder")
    editActive = CBool(rsE("IsActive"))
  End If
  rsE.Close : Set rsE = Nothing
  Set cmdE = Nothing
End If

' list categories

Dim rs : Set rs = conn.Execute("SELECT CategoryID, Name, Slug, SortOrder, IsActive FROM dbo.Categories ORDER BY SortOrder, Name")

%>
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>Admin - Categories</title>
  <style>
    body{font-family:Segoe UI,Arial; padding:24px;}
    table{border-collapse:collapse; width:100%; margin-top:14px;}
    th,td{border:1px solid #ddd; padding:8px;}
    th{background:#f5f5f5; text-align:left;}
    .row{display:flex; gap:18px; align-items:flex-start;}
    .card{border:1px solid #ddd; border-radius:10px; padding:14px; min-width:360px;}
    input{padding:8px; width:100%; box-sizing:border-box;}
    .btn{padding:8px 12px; cursor:pointer;}
    .msg{margin:10px 0; color:#0b6; font-weight:600;}
    .muted{color:#666;}
    .actions form{display:inline;}
  </style>
</head>
<body>
  <div class="row">
    <div style="flex:1">
      <h2>Categories</h2>
      <div class="muted"><a href="default.asp">Admin Home</a> | <a href="logout.asp">Logout</a></div>
      <% If Len(msg)>0 Then %><div class="msg"><%=Server.HTMLEncode(msg)%></div><% End If %>

      <table>
        <thead>
          <tr>
            <th>ID</th><th>Name</th><th>Slug</th><th>Sort</th><th>Active</th><th>Actions</th>
          </tr>
        </thead>
        <tbody>
        <% Do While Not rs.EOF %>
          <tr>
            <td><%=rs("CategoryID")%></td>
            <td><%=Server.HTMLEncode(rs("Name"))%></td>
            <td><%=Server.HTMLEncode(rs("Slug"))%></td>
            <td><%=rs("SortOrder")%></td>
            <td><%=IIf(CBool(rs("IsActive")), "Yes", "No")%></td>
            <td class="actions">
              <a href="categories.asp?edit=<%=rs("CategoryID")%>">Edit</a>
              &nbsp;|&nbsp;
              <form method="post" action="categories.asp" onsubmit="return confirm('Deactivate this category?');">
                <input type="hidden" name="action" value="delete" />
                <input type="hidden" name="CategoryID" value="<%=rs("CategoryID")%>" />
                <button class="btn" type="submit">Deactivate</button>
              </form>
            </td>
          </tr>
        <% rs.MoveNext : Loop %>
        </tbody>
      </table>
    </div>

    <div class="card">
      <% If editId > 0 Then %>
        <h3>Edit Category</h3>
        <form method="post" action="categories.asp">
          <input type="hidden" name="action" value="update" />
          <input type="hidden" name="CategoryID" value="<%=editId%>" />

          <label>Name</label>
          <input name="Name" value="<%=Server.HTMLEncode(editName)%>" />

          <label>Slug</label>
          <input name="Slug" value="<%=Server.HTMLEncode(editSlug)%>" />

          <label>Sort Order</label>
          <input name="SortOrder" value="<%=editSort%>" />

          <label>Active</label>
          <select name="IsActive">
            <option value="1" <%=IIf(editActive,"selected","")%>>Yes</option>
            <option value="0" <%=IIf(Not editActive,"selected","")%>>No</option>
          </select>

          <div style="margin-top:12px;">
            <button class="btn" type="submit">Save</button>
            <a class="btn" href="categories.asp" style="text-decoration:none;">Cancel</a>
          </div>
        </form>
      <% Else %>
        <h3>Add Category</h3>
        <form method="post" action="categories.asp">
          <input type="hidden" name="action" value="add" />

          <label>Name</label>
          <input name="Name" />

          <label>Slug (optional)</label>
          <input name="Slug" />

          <label>Sort Order</label>
          <input name="SortOrder" value="0" />

          <label>Active</label>
          <select name="IsActive">
            <option value="1" selected>Yes</option>
            <option value="0">No</option>
          </select>

          <div style="margin-top:12px;">
            <button class="btn" type="submit">Add</button>
          </div>
        </form>
      <% End If %>
    </div>
  </div>
</body>
</html>
<%
rs.Close : Set rs = Nothing
conn.Close : Set conn = Nothing
%>