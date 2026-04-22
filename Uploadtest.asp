<%
Response.Buffer = True
Dim result : result = ""
If Request.QueryString("result") <> "" Then result = Server.HTMLEncode(Request.QueryString("result"))
%>
<!DOCTYPE html>
<html>
<head><title>Upload Test</title></head>
<body>
<h2>Upload Test</h2>
<% If result <> "" Then Response.Write "<p><strong>" & result & "</strong></p>" End If %>
<form method="post" enctype="multipart/form-data" action="/upload.ashx">
  <input type="file" name="file1" /><br /><br />
  <input type="submit" value="Upload" />
</form>
<p><small>Mapped path: <%=Server.MapPath(".")%></small></p>
</body>
</html>
