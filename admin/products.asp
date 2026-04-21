<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="_auth.asp"-->
<!--#include virtual="/includes/adovbs.asp"-->
<!--#include virtual="/includes/datacon.asp"-->
<!--#include virtual="/includes/code.asp"-->

<%
Dim conn : Set conn = OpenConn()

Dim sql
sql = ""
sql = sql & "SELECT p.ProductID, p.Name, p.Slug, p.Price, p.SortOrder, p.IsActive, "
sql = sql & "       c.Name AS CategoryName "
sql = sql & "FROM dbo.Products p "
sql = sql & "INNER JOIN dbo.Categories c ON c.CategoryID = p.CategoryID "
sql = sql & "ORDER BY c.SortOrder, c.Name, p.SortOrder, p.Name;"

Dim rs : Set rs = conn.Execute(sql)
%>
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>Admin - Products</title>
  <style>
    body{font-family:Segoe UI,Arial; padding:24px;}
    table{border-collapse:collapse; width:100%; margin-top:14px;}
    th,td{border:1px solid #ddd; padding:8px;}
    th{background:#f5f5f5; text-align:left;}
    .topbar{display:flex; justify-content:space-between; align-items:center; gap:12px;}
    .btn{display:inline-block; padding:8px 12px; border:1px solid #333; border-radius:8px; text-decoration:none; color:#111;}
    .muted{color:#666;}
  </style>
</head>
<body>
  <div class="topbar">
    <div>
      <h2 style="margin:0;">Products</h2>
      <div class="muted"><a href="default.asp">Admin Home</a> | <a href="categories.asp">Categories</a> | <a href="logout.asp">Logout</a></div>
    </div>
    <div>
      <a class="btn" href="product-edit.asp">+ Add Product</a>
    </div>
  </div>

  <table>
    <thead>
      <tr>
        <th>ID</th>
        <th>Category</th>
        <th>Name</th>
        <th>Slug</th>
        <th>Price</th>
        <th>Sort</th>
        <th>Active</th>
        <th></th>
      </tr>
    </thead>
    <tbody>
      <% Do While Not rs.EOF %>
        <tr>
          <td><%=rs("ProductID")%></td>
          <td><%=Server.HTMLEncode(rs("CategoryName"))%></td>
          <td><%=Server.HTMLEncode(rs("Name"))%></td>
          <td><%=Server.HTMLEncode(rs("Slug"))%></td>
          <td>
            <% If IsNull(rs("Price")) Then %>
              -
            <% Else %>
              $<%=FormatNumber(rs("Price"), 2, -1, 0, -1)%>
            <% End If %>
          </td>
          <td><%=rs("SortOrder")%></td>
          <td><%=IIf(CBool(rs("IsActive")), "Yes", "No")%></td>
          <td><a href="product-edit.asp?id=<%=rs("ProductID")%>">Edit</a></td>
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