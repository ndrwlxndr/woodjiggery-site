<%
'Option Explicit
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
%>
<!--#include virtual="/admin/_auth.asp"-->
<!--#include virtual="/includes/adovbs.asp"-->
<!--#include virtual="/includes/datacon.asp"-->
<%

' Debug helper: prints VB + ADO errors and stops
Sub DieWithAdo(label, conn)
  Dim e
  Response.Write "<hr><b>" & Server.HTMLEncode(label) & "</b><br>"
  Response.Write "VB Err: " & Err.Number & " - " & Server.HTMLEncode(Err.Description) & "<br>"

  If Not (conn Is Nothing) Then
    If conn.Errors.Count > 0 Then
      For Each e In conn.Errors
        Response.Write "ADO: " & e.Number & " - " & Server.HTMLEncode(e.Description) & "<br>"
      Next
    Else
      Response.Write "ADO: (no connection errors)<br>"
    End If
  End If
  Response.End
End Sub

ON error resume next
Dim conn : Set conn = OpenConn()

Dim productId : productId = 0
If Len(Request.QueryString("id")) > 0 Then productId = CLng(Request.QueryString("id"))

' Load categories for dropdown
Dim rsCats : Set rsCats = conn.Execute("SELECT CategoryID, Name FROM dbo.Categories WHERE IsActive=1 ORDER BY SortOrder, Name;")

' Defaults
Dim CategoryID, Name, Slug, ShortDesc, LongDesc, Price, SortOrder, IsActive
CategoryID = 0
Name = ""
Slug = ""
ShortDesc = ""
LongDesc = ""
Price = ""
SortOrder = 0
IsActive = True

Dim msg : msg = ""
Dim errMsg : errMsg = ""

' If editing, load existing product
If productId > 0 Then
  Dim cmdL : Set cmdL = Server.CreateObject("ADODB.Command")
  Set cmdL.ActiveConnection = conn
  cmdL.CommandType = adCmdText
  cmdL.CommandText = "SELECT ProductID, CategoryID, Name, Slug, ShortDesc, LongDesc, Price, SortOrder, IsActive " & _
                     "FROM dbo.Products WHERE ProductID=?;"
  cmdL.Parameters.Append cmdL.CreateParameter("@ID", adInteger, adParamInput, , CInt(productId))

  Dim rsL : Set rsL = cmdL.Execute
  If Not rsL.EOF Then
    CategoryID = rsL("CategoryID")
    Name = rsL("Name")
    Slug = rsL("Slug")
    If Not IsNull(rsL("ShortDesc")) Then ShortDesc = rsL("ShortDesc")
    If Not IsNull(rsL("LongDesc")) Then LongDesc = rsL("LongDesc")
    If Not IsNull(rsL("Price")) Then Price = CStr(rsL("Price"))
    SortOrder = rsL("SortOrder")
    IsActive = CBool(rsL("IsActive"))
  Else
    errMsg = "Product not found."
  End If
  rsL.Close : Set rsL = Nothing
  Set cmdL = Nothing
End If

' Handle POST (save)
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
  Dim fCat, fName, fSlug, fShort, fLong, fPrice, fSort, fActive
  fCat = Trim(Request.Form("CategoryID"))
  fName = Trim(Request.Form("Name"))
  fSlug = Trim(Request.Form("Slug"))
  fShort = Trim(Request.Form("ShortDesc"))
  fLong = Trim(Request.Form("LongDesc"))
  fPrice = Trim(Request.Form("Price"))
  fSort = Trim(Request.Form("SortOrder"))
  fActive = (Request.Form("IsActive") = "1")

  If Len(fName) = 0 Then
    errMsg = "Name is required."
  ElseIf Len(fCat) = 0 Or CLng(fCat) = 0 Then
    errMsg = "Category is required."
  Else
    If Len(fSlug) = 0 Then fSlug = Slugify(fName)
    If Len(fSort) = 0 Then fSort = 0

    ' Update local vars so form re-renders with user input
    CategoryID = CLng(fCat)
    Name = fName
    Slug = fSlug
    ShortDesc = fShort
    LongDesc = fLong
    Price = fPrice
    SortOrder = CLng(fSort)
    IsActive = fActive

    If productId = 0 Then


Response.Buffer = True
On Error Resume Next
Err.Clear


' INSERT
' --- IMPORTANT: localize error handling so we SEE failures ---
On Error GoTo 0

' INSERT
Dim cmdI : Set cmdI = Server.CreateObject("ADODB.Command")
Set cmdI.ActiveConnection = conn
cmdI.CommandType = adCmdText
cmdI.CommandText = "SET NOCOUNT ON; " & _
                   "INSERT INTO dbo.Products (CategoryID, Name, Slug, ShortDesc, LongDesc, Price, SortOrder, IsActive) " & _
                   "VALUES (?,?,?,?,?,?,?,?); " & _
                   "SELECT SCOPE_IDENTITY() AS NewID;"

On Error Resume Next
Err.Clear
conn.Errors.Clear

cmdI.Parameters.Append cmdI.CreateParameter("@CategoryID", adInteger, adParamInput, , CLng(CategoryID))
cmdI.Parameters.Append cmdI.CreateParameter("@Name", adVarWChar, adParamInput, 150, CStr(Name))
cmdI.Parameters.Append cmdI.CreateParameter("@Slug", adVarWChar, adParamInput, 150, CStr(Slug))

Dim p, longSize

Set p = cmdI.CreateParameter("@ShortDesc", adVarWChar, adParamInput, 300)
If Len(ShortDesc) = 0 Then p.Value = Null Else p.Value = CStr(ShortDesc)
cmdI.Parameters.Append p
Set p = Nothing

longSize = 1
If Len(LongDesc) > 0 Then longSize = Len(LongDesc)
Set p = cmdI.CreateParameter("@LongDesc", adLongVarWChar, adParamInput, longSize)
If Len(LongDesc) = 0 Then p.Value = Null Else p.Value = CStr(LongDesc)
cmdI.Parameters.Append p
Set p = Nothing

Set p = cmdI.CreateParameter("@Price", adCurrency, adParamInput)
If Len(Trim(Price)) = 0 Then
  p.Value = Null
Else
  p.Value = CDbl(Replace(Replace(Price, "$", ""), ",", ""))
End If
cmdI.Parameters.Append p
Set p = Nothing

cmdI.Parameters.Append cmdI.CreateParameter("@SortOrder", adInteger, adParamInput, , CLng(SortOrder))
cmdI.Parameters.Append cmdI.CreateParameter("@IsActive", adBoolean, adParamInput, , CBool(IsActive))

If Err.Number <> 0 Then DieWithAdo "Parameter build failed (INSERT)", conn

Dim rsNew : Set rsNew = cmdI.Execute
If Err.Number <> 0 Then DieWithAdo "cmdI.Execute failed (INSERT)", conn

productId = CLng(rsNew("NewID"))
rsNew.Close : Set rsNew = Nothing
Set cmdI = Nothing
msg = "Product created."

On Error GoTo 0
    Else
      ' UPDATE
      Dim cmdU : Set cmdU = Server.CreateObject("ADODB.Command")
      Set cmdU.ActiveConnection = conn
      cmdU.CommandType = adCmdText
      cmdU.CommandText = "UPDATE dbo.Products " & _
                         "SET CategoryID=?, Name=?, Slug=?, ShortDesc=?, LongDesc=?, Price=?, SortOrder=?, IsActive=? " & _
                         "WHERE ProductID=?;"

      cmdU.Parameters.Append cmdU.CreateParameter("@CategoryID", adInteger,  adParamInput, , CInt(CategoryID))
      cmdU.Parameters.Append cmdU.CreateParameter("@Name",       adVarWChar, adParamInput, 150, Name)
      cmdU.Parameters.Append cmdU.CreateParameter("@Slug",       adVarWChar, adParamInput, 150, Slug)
      cmdU.Parameters.Append cmdU.CreateParameter("@ShortDesc",  adVarWChar, adParamInput, 300, IIf(Len(ShortDesc)=0, Null, ShortDesc))
      cmdU.Parameters.Append cmdU.CreateParameter("@LongDesc",   adLongVarWChar, adParamInput, -1, IIf(Len(LongDesc)=0, Null, LongDesc))

      If Len(Price) = 0 Then
        cmdU.Parameters.Append cmdU.CreateParameter("@Price", adCurrency, adParamInput, , Null)
      Else
        cmdU.Parameters.Append cmdU.CreateParameter("@Price", adCurrency, adParamInput, , CDbl(Price))
      End If

      cmdU.Parameters.Append cmdU.CreateParameter("@SortOrder", adInteger,  adParamInput, , CInt(SortOrder))
      cmdU.Parameters.Append cmdU.CreateParameter("@IsActive",  adBoolean,  adParamInput, , CBool(IsActive))
      cmdU.Parameters.Append cmdU.CreateParameter("@ProductID", adInteger,  adParamInput, , CInt(productId))

      cmdU.Execute
      Set cmdU = Nothing

      msg = "Product updated."
    End If
  End If
End If
%>
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>Admin - Product</title>
  <style>
    body{font-family:Segoe UI,Arial; padding:24px; max-width:900px;}
    label{display:block; margin-top:12px; font-weight:600;}
    input,select,textarea{width:100%; padding:10px; box-sizing:border-box;}
    textarea{min-height:140px;}
    .row{display:flex; gap:12px;}
    .row > div{flex:1;}
    .msg{margin:10px 0; color:#0b6; font-weight:600;}
    .err{margin:10px 0; color:#b00020; font-weight:600;}
    .btn{display:inline-block; padding:8px 12px; border:1px solid #333; border-radius:8px; text-decoration:none; color:#111; background:#fff; cursor:pointer;}
    .muted{color:#666;}
  </style>
</head>
<body>
  <h2><% If productId=0 Then %>Add Product<% Else %>Edit Product #<%=productId%><% End If %></h2>
  <div class="muted"><a href="products.asp">Back to Products</a> | <a href="logout.asp">Logout</a></div>

  <% If Len(msg)>0 Then %><div class="msg"><%=Server.HTMLEncode(msg)%></div><% End If %>
  <% If Len(errMsg)>0 Then %><div class="err"><%=Server.HTMLEncode(errMsg)%></div><% End If %>

  <form method="post" action="product-edit.asp<% If productId>0 Then Response.Write("?id=" & productId) %>">
    <label>Category</label>
    <select name="CategoryID">
      <option value="0">-- select --</option>
      <%
        rsCats.MoveFirst
        Do While Not rsCats.EOF
          Dim cid : cid = CLng(rsCats("CategoryID"))
      %>
          <option value="<%=cid%>" <% If CLng(CategoryID)=cid Then Response.Write("selected") %>>
            <%=Server.HTMLEncode(rsCats("Name"))%>
          </option>
      <%
          rsCats.MoveNext
        Loop
      %>
    </select>

    <label>Name</label>
    <input name="Name" value="<%=Server.HTMLEncode(Name)%>" />

    <div class="row">
      <div>
        <label>Slug (optional)</label>
        <input name="Slug" value="<%=Server.HTMLEncode(Slug)%>" />
      </div>
      <div>
        <label>Price (optional)</label>
        <input name="Price" value="<%=Server.HTMLEncode(Price)%>" placeholder="e.g. 49.95" />
      </div>
    </div>

    <div class="row">
      <div>
        <label>Sort Order</label>
        <input name="SortOrder" value="<%=SortOrder%>" />
      </div>
      <div>
        <label>Active</label>
        <select name="IsActive">
          <option value="1" <% If IsActive Then Response.Write("selected") %>>Yes</option>
          <option value="0" <% If Not IsActive Then Response.Write("selected") %>>No</option>
        </select>
      </div>
    </div>

    <label>Short Description</label>
    <input name="ShortDesc" value="<%=Server.HTMLEncode(ShortDesc)%>" />

    <label>Long Description</label>
    <textarea name="LongDesc"><%=Server.HTMLEncode(LongDesc)%></textarea>

    <div style="margin-top:16px;">
      <button class="btn" type="submit">Save</button>
      <a class="btn" href="products.asp">Cancel</a>
      <% If productId>0 Then %>
        <a class="btn" href="product-images.asp?productid=<%=productId%>">Images…</a>
      <% End If %>
    </div>
  </form>
</body>
</html>
<%
rsCats.Close : Set rsCats = Nothing
conn.Close : Set conn = Nothing
%>