<%
Option Explicit
Response.Buffer = True
%>
<!--#include virtual="/admin/_auth.asp"-->
<!--#include virtual="/includes/adovbs.asp"-->
<!--#include virtual="/includes/datacon.asp"-->
<!--#include virtual="/includes/upload.asp"-->
<!--#include virtual="/includes/code.asp"-->
<%
Dim conn : Set conn = OpenConn()

Dim productId : productId = 0
If Len(Request.QueryString("productid")) > 0 Then productId = CLng(Request.QueryString("productid"))
If productId = 0 Then
  Response.Write "Missing productid"
  Response.End
End If

Dim msg : msg = ""
Dim errMsg : errMsg = ""

On Error GoTo 0

'--- Helpers ---
Function SafeExt(ByVal fn)
  Dim p : p = InStrRev(fn, ".")
  If p = 0 Then SafeExt = "" : Exit Function
  SafeExt = LCase(Mid(fn, p+1))
End Function

Function SafeFileBase(ByVal s)
  Dim i, ch, out
  s = Trim(s)
  out = ""
  For i = 1 To Len(s)
    ch = Mid(s, i, 1)
    If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
      out = out & ch
    ElseIf ch = "-" Or ch = "_" Then
      out = out & ch
    Else
      out = out & "_"
    End If
  Next
  If Len(out) = 0 Then out = "img"
  SafeFileBase = out
End Function

Sub EnsureFolder(ByVal physPath)
  Dim fso : Set fso = Server.CreateObject("Scripting.FileSystemObject")
  If Not fso.FolderExists(physPath) Then
    fso.CreateFolder physPath
  End If
  Set fso = Nothing
End Sub

Sub ClearPrimaries(ByVal pid)
  Dim cmdClr : Set cmdClr = Server.CreateObject("ADODB.Command")
  Set cmdClr.ActiveConnection = conn
  cmdClr.CommandType = adCmdText
  cmdClr.CommandText = "UPDATE dbo.ProductImages SET IsPrimary=0 WHERE ProductID=?;"
  cmdClr.Parameters.Append cmdClr.CreateParameter("@PID", adInteger, adParamInput, , CInt(pid))
  cmdClr.Execute
  Set cmdClr = Nothing
End Sub

Sub InsertImageRow(ByVal pid, ByVal imgPath, ByVal altText, ByVal sortOrder, ByVal isPrimary)
  Dim cmdA : Set cmdA = Server.CreateObject("ADODB.Command")
  Set cmdA.ActiveConnection = conn
  cmdA.CommandType = adCmdText
  cmdA.CommandText = "INSERT INTO dbo.ProductImages (ProductID, ImagePath, AltText, SortOrder, IsPrimary) VALUES (?,?,?,?,?);"
  cmdA.Parameters.Append cmdA.CreateParameter("@PID",  adInteger,  adParamInput, , CInt(pid))
  cmdA.Parameters.Append cmdA.CreateParameter("@Path", adVarWChar, adParamInput, 260, CStr(imgPath))

  Dim pAlt : Set pAlt = cmdA.CreateParameter("@Alt", adVarWChar, adParamInput, 200)
  If Len(Trim(altText)) = 0 Then
    pAlt.Value = Null
  Else
    pAlt.Value = CStr(altText)
  End If
  cmdA.Parameters.Append pAlt
  Set pAlt = Nothing

  cmdA.Parameters.Append cmdA.CreateParameter("@Sort", adInteger, adParamInput, , CLng(sortOrder))
  cmdA.Parameters.Append cmdA.CreateParameter("@Prim", adBoolean, adParamInput, , CBool(isPrimary))
  cmdA.Execute
  Set cmdA = Nothing
End Sub

Function GetImageCount(ByVal pid)
  Dim rsCnt
  Set rsCnt = conn.Execute("SELECT COUNT(*) AS Cnt FROM dbo.ProductImages WHERE ProductID=" & CLng(pid))
  GetImageCount = CLng(rsCnt("Cnt"))
  rsCnt.Close : Set rsCnt = Nothing
End Function

' Get product name
Dim prodName : prodName = ""
Dim cmdP : Set cmdP = Server.CreateObject("ADODB.Command")
Set cmdP.ActiveConnection = conn
cmdP.CommandType = adCmdText
cmdP.CommandText = "SELECT Name FROM dbo.Products WHERE ProductID=?;"
cmdP.Parameters.Append cmdP.CreateParameter("@ID", adInteger, adParamInput, , CInt(productId))
Dim rsP : Set rsP = cmdP.Execute
If Not rsP.EOF Then prodName = rsP("Name")
rsP.Close : Set rsP = Nothing
Set cmdP = Nothing

' ------------------------------
' POST handling (uses Upload class for file uploads, but also handles non-file actions like setting primary and deleting)
' ------------------------------
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

  Dim up : Set up = New Upload
  Dim actionPost : actionPost = LCase(Trim("" & up.Form("action")))

  Select Case actionPost

    Case "upload"
      Dim f : Set f = Nothing
      If up.Files.Exists("imgFile") Then Set f = up.Files("imgFile")

      If f Is Nothing Then
        errMsg = "No file selected."
      Else
        Dim altTextU, sortOrderU, isPrimaryU
        altTextU = Trim("" & up.Form("AltText"))
        sortOrderU = Trim("" & up.Form("SortOrder"))
        If Len(sortOrderU) = 0 Then sortOrderU = 0
        isPrimaryU = ("" & up.Form("IsPrimary") = "1")

        Dim imgCountU : imgCountU = GetImageCount(productId)

        ' If this is the first image ever for this product, force it to primary
        If imgCountU = 0 Then
            isPrimaryU = True
        End If

        Dim ext : ext = SafeExt(f.FileName)
        If ext <> "jpg" And ext <> "jpeg" And ext <> "png" And ext <> "webp" And ext <> "gif" Then
          errMsg = "Only jpg/jpeg/png/webp/gif allowed."
        Else
          ' Ensure /images/products and /images/products/<id>
          Dim physBase, physFolder, relFolder
          relFolder = "/images/products/" & productId
          physBase = Server.MapPath("/images/products")
          physFolder = Server.MapPath(relFolder)

          EnsureFolder physBase
          EnsureFolder physFolder

          ' Safe filename
          Dim baseName, newName
          baseName = SafeFileBase(prodName)
          newName = baseName & "_" & _
                    Year(Now()) & Right("0" & Month(Now()),2) & Right("0" & Day(Now()),2) & "_" & _
                    Right("0" & Hour(Now()),2) & Right("0" & Minute(Now()),2) & Right("0" & Second(Now()),2) & _
                    "." & ext

          On Error Resume Next
          Err.Clear
          f.SaveToAs physFolder, newName   ' <-- uses the new method
          If Err.Number <> 0 Then
            errMsg = "Upload failed: " & Err.Description
            Err.Clear
          Else
            Dim webPath : webPath = relFolder & "/" & newName

            If isPrimaryU Then ClearPrimaries productId
            InsertImageRow productId, webPath, altTextU, CLng(sortOrderU), isPrimaryU
            msg = "Image uploaded."
          End If
          On Error GoTo 0
        End If
      End If

    Case "addpath"
      Dim imagePath, altTextP, sortOrderP, isPrimaryP
      imagePath = Trim("" & up.Form("ImagePath"))
      altTextP = Trim("" & up.Form("AltText"))
      sortOrderP = Trim("" & up.Form("SortOrder"))
      If Len(sortOrderP) = 0 Then sortOrderP = 0
      isPrimaryP = ("" & up.Form("IsPrimary") = "1")

      Dim imgCountP : imgCountP = GetImageCount(productId)

        ' If this is the first image ever for this product, force it to primary
        If imgCountP = 0 Then
        isPrimaryP = True
        End If

      If Len(imagePath) = 0 Then
        errMsg = "ImagePath is required."
      Else
        If isPrimaryP Then ClearPrimaries productId
        InsertImageRow productId, imagePath, altTextP, CLng(sortOrderP), isPrimaryP
        msg = "Image added by path."
      End If

    Case "primary"
      Dim imageIdP : imageIdP = CLng("" & up.Form("ImageID"))
      ClearPrimaries productId

      Dim cmdPri : Set cmdPri = Server.CreateObject("ADODB.Command")
      Set cmdPri.ActiveConnection = conn
      cmdPri.CommandType = adCmdText
      cmdPri.CommandText = "UPDATE dbo.ProductImages SET IsPrimary=1 WHERE ProductID=? AND ImageID=?;"
      cmdPri.Parameters.Append cmdPri.CreateParameter("@PID", adInteger, adParamInput, , CInt(productId))
      cmdPri.Parameters.Append cmdPri.CreateParameter("@IID", adInteger, adParamInput, , CInt(imageIdP))
      cmdPri.Execute
      Set cmdPri = Nothing

      msg = "Primary image updated."

    Case "delete"
      Dim imageIdD : imageIdD = CLng("" & up.Form("ImageID"))

      Dim cmdD : Set cmdD = Server.CreateObject("ADODB.Command")
      Set cmdD.ActiveConnection = conn
      cmdD.CommandType = adCmdText
      cmdD.CommandText = "DELETE FROM dbo.ProductImages WHERE ProductID=? AND ImageID=?;"
      cmdD.Parameters.Append cmdD.CreateParameter("@PID", adInteger, adParamInput, , CInt(productId))
      cmdD.Parameters.Append cmdD.CreateParameter("@IID", adInteger, adParamInput, , CInt(imageIdD))
      cmdD.Execute
      Set cmdD = Nothing

      msg = "Image removed."
  End Select

  Set up = Nothing
End If

Dim sql
sql = "SELECT ImageID, ImagePath, AltText, SortOrder, IsPrimary " & _
      "FROM dbo.ProductImages WHERE ProductID=" & productId & " " & _
      "ORDER BY IsPrimary DESC, SortOrder, ImageID;"

Dim rs : Set rs = conn.Execute(sql)
%>
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>Admin - Product Images</title>
  <style>
    body{font-family:Segoe UI,Arial; padding:24px; max-width:1100px;}
    table{border-collapse:collapse; width:100%; margin-top:14px;}
    th,td{border:1px solid #ddd; padding:8px; vertical-align:top;}
    th{background:#f5f5f5; text-align:left;}
    .msg{margin:10px 0; color:#0b6; font-weight:600;}
    .err{margin:10px 0; color:#b00020; font-weight:600;}
    .btn{padding:6px 10px; border:1px solid #333; border-radius:8px; background:#fff; cursor:pointer;}
    input{width:100%; padding:8px; box-sizing:border-box;}
    .grid{display:grid; grid-template-columns: 1fr 1fr; gap:16px; margin-top:12px;}
    img{max-width:240px; max-height:160px; object-fit:cover; border:1px solid #ddd; border-radius:8px;}
    .card{border:1px solid #ddd; border-radius:10px; padding:12px; background:#fff;}
    code{background:#f6f6f6; padding:2px 6px; border-radius:6px;}
  </style>
</head>
<body>
  <h2>Images for: <%=Server.HTMLEncode(prodName)%> (#<%=productId%>)</h2>
  <div><a href="product-edit.asp?id=<%=productId%>">Back to Product</a> | <a href="products.asp">Products</a></div>

  <% If Len(msg)>0 Then %><div class="msg"><%=Server.HTMLEncode(msg)%></div><% End If %>
  <% If Len(errMsg)>0 Then %><div class="err"><%=Server.HTMLEncode(errMsg)%></div><% End If %>

  <div class="grid">
    <div class="card">
      <h3 style="margin-top:0;">Upload image</h3>
      <form method="post" enctype="multipart/form-data" action="product-images.asp?productid=<%=productId%>">
        <input type="hidden" name="action" value="upload" />

        <label>Image file</label>
        <input type="file" name="imgFile" accept=".jpg,.jpeg,.png,.webp,.gif" />

        <label>AltText (optional)</label>
        <input name="AltText" />

        <label>SortOrder</label>
        <input name="SortOrder" value="0" />

        <label><input type="checkbox" name="IsPrimary" value="1" /> Set as primary (first image is automatically primary)</label>

        <div style="margin-top:10px;">
          <button class="btn" type="submit">Upload</button>
        </div>
      </form>
      <p style="color:#666;">Uploads to: <code>/images/products/<%=productId%>/</code></p>
    </div>

    <div class="card">
      <h3 style="margin-top:0;">Add by path (optional)</h3>
      <p style="color:#666;margin-top:-6px;">Fallback if permissions block uploads.</p>
      <form method="post" action="product-images.asp?productid=<%=productId%>">
        <input type="hidden" name="action" value="addpath" />
        <label>ImagePath</label>
        <input name="ImagePath" placeholder="/images/products/<%=productId%>/hero.jpg" />
        <label>AltText (optional)</label>
        <input name="AltText" />
        <label>SortOrder</label>
        <input name="SortOrder" value="0" />
        <label><input type="checkbox" name="IsPrimary" value="1" /> Set as primary</label>
        <div style="margin-top:10px;">
          <button class="btn" type="submit">Add</button>
        </div>
      </form>
    </div>
  </div>

  <h3>Current images</h3>
  <table>
    <thead>
      <tr><th>Preview</th><th>Path</th><th>Alt</th><th>Sort</th><th>Primary</th><th>Actions</th></tr>
    </thead>
    <tbody>
      <% If rs.EOF Then %>
        <tr><td colspan="6" style="color:#666;">No images yet.</td></tr>
      <% End If %>

      <% Do While Not rs.EOF %>
        <tr>
          <td><img src="<%=Server.HTMLEncode(rs("ImagePath"))%>" alt=""></td>
          <td><%=Server.HTMLEncode(rs("ImagePath"))%></td>
          <td><%=Server.HTMLEncode("" & rs("AltText"))%></td>
          <td><%=rs("SortOrder")%></td>
          <td><%=IIf(CBool(rs("IsPrimary")), "Yes", "No")%></td>
          <td>
            <form method="post" action="product-images.asp?productid=<%=productId%>" style="display:inline;">
              <input type="hidden" name="action" value="primary" />
              <input type="hidden" name="ImageID" value="<%=rs("ImageID")%>" />
              <button class="btn" type="submit">Make Primary</button>
            </form>
            <form method="post" action="product-images.asp?productid=<%=productId%>" style="display:inline;" onsubmit="return confirm('Delete this image record?');">
              <input type="hidden" name="action" value="delete" />
              <input type="hidden" name="ImageID" value="<%=rs("ImageID")%>" />
              <button class="btn" type="submit">Delete</button>
            </form>
          </td>
        </tr>
      <% rs.MoveNext : Loop %>
    </tbody>
  </table>
</body>
</html>
<%
rs.Close : Set rs = Nothing
conn.Close : Set conn = Nothing
%>