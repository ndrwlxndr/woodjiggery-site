<%
Option Explicit
Response.Buffer = True
%>
<!--#include virtual="/includes/adovbs.asp"-->
<!--#include virtual="/includes/datacon.asp"-->
<%
Function H(s) : H = Server.HTMLEncode("" & s) : End Function

Dim conn : Set conn = OpenConn()

Dim sql
sql = ""
sql = sql & "SELECT c.CategoryID, c.Name, c.Slug, "
sql = sql & "       img.ImagePath "
sql = sql & "FROM dbo.Categories c "
sql = sql & "OUTER APPLY ("
sql = sql & "  SELECT TOP 1 i.ImagePath "
sql = sql & "  FROM dbo.Products p "
sql = sql & "  INNER JOIN dbo.ProductImages i ON i.ProductID = p.ProductID "
sql = sql & "  WHERE p.IsActive=1 AND p.CategoryID = c.CategoryID "
sql = sql & "  ORDER BY i.IsPrimary DESC, i.SortOrder, i.ImageID "
sql = sql & ") img "
sql = sql & "WHERE c.IsActive=1 "
sql = sql & "ORDER BY c.SortOrder, c.Name;"

Dim rs : Set rs = conn.Execute(sql)
%>
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>Creations - WoodJiggery</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <style>
    body{font-family:Segoe UI,Arial; margin:0; padding:24px; background:#fafafa; color:#111;}
    .wrap{max-width:1100px; margin:0 auto;}
    .grid{display:grid; grid-template-columns:repeat(auto-fill, minmax(240px, 1fr)); gap:16px; margin-top:16px;}
    .card{background:#fff; border:1px solid #e2e2e2; border-radius:14px; overflow:hidden; text-decoration:none; color:inherit;}
    .img{height:170px; background:#f0f0f0; display:flex; align-items:center; justify-content:center;}
    .img img{width:100%; height:100%; object-fit:cover; display:block;}
    .pad{padding:12px 12px 14px;}
    .name{font-weight:800; margin:0;}
    .muted{color:#666; font-size:14px;}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="muted"><a href="/index.asp">Home</a> &nbsp;›&nbsp; Creations</div>
    <h1 style="margin:10px 0 0;">Creations</h1>
    <div class="muted">Pick a category to browse products.</div>

    <div class="grid">
      <% Do While Not rs.EOF %>
        <%
          Dim catName, catSlug, imgPath
          catName = rs("Name")
          catSlug = rs("Slug")
          imgPath = ""
          If Not IsNull(rs("ImagePath")) Then imgPath = rs("ImagePath")
        %>
        <a class="card" href="/category.asp?slug=<%=Server.URLEncode(catSlug)%>">
          <div class="img">
            <% If Len(imgPath)>0 Then %>
              <img src="<%=H(imgPath)%>" alt="<%=H(catName)%>">
            <% Else %>
              <span class="muted">No image yet</span>
            <% End If %>
          </div>
          <div class="pad">
            <p class="name"><%=H(catName)%></p>
          </div>
        </a>
      <% rs.MoveNext : Loop %>
    </div>
  </div>
</body>
</html>
<%
rs.Close : Set rs = Nothing
conn.Close : Set conn = Nothing
%>