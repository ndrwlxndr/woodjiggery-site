<%
Option Explicit
Response.Buffer = True

Function H(s) : H = Server.HTMLEncode("" & s) : End Function

%>
<!--#include virtual="/includes/adovbs.asp"-->
<!--#include virtual="/includes/datacon.asp"-->
<%
Dim conn : Set conn = OpenConn()

Dim sql, rs
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

Set rs = conn.Execute(sql)
%>
<!DOCTYPE html>
<html style="font-size: 16px;" lang="en">
<head>
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta charset="utf-8">
  <title>Creations</title>

  <!-- Keep Nicepage base styling + behavior -->
  <link rel="stylesheet" href="nicepage.css" media="screen">
  <script class="u-script" type="text/javascript" src="jquery.js" defer=""></script>
  <script class="u-script" type="text/javascript" src="nicepage.js" defer=""></script>

  <link id="u-page-google-font" rel="stylesheet"
        href="https://fonts.googleapis.com/css2?display=swap&family=Roboto:ital,wght@0,100;0,200;0,300;0,400;0,500;0,600;0,700;0,800;0,900;1,100;1,200;1,300;1,400;1,500;1,600;1,700;1,800;1,900&family=Open+Sans:ital,wght@0,300;0,400;0,500;0,600;0,700;0,800;1,300;1,400;1,500;1,600;1,700;1,800">

  <script type="application/ld+json">
  {
    "@context": "http://schema.org",
    "@type": "Organization",
    "name": "WoodJiggery",
    "url": "/",
    "logo": "images/Woodjiggerygray.svg"
  }
  </script>

  <style>
    /* Minimal catalog styling, scoped so it doesn't fight Nicepage */
    .wj-crumb { color: #cfcfcf; margin: 10px 0 0; font-size: 14px; }
    .wj-title { margin: 10px 0 0; }
    .wj-sub   { color: #cfcfcf; margin: 6px 0 16px; }

    .wj-grid{
      display:grid;
      grid-template-columns:repeat(auto-fill, minmax(240px, 1fr));
      gap:16px;
      margin-top:16px;
    }
    .wj-card{
      display:block;
      text-decoration:none;
      color: inherit;
      background: rgba(255,255,255,.06);
      border: 1px solid rgba(255,255,255,.14);
      border-radius: 14px;
      overflow:hidden;
      transition: transform .12s ease, background .12s ease;
    }
    .wj-card:hover{ transform: translateY(-2px); background: rgba(255,255,255,.09); }

    .wj-img{
      height: 170px;
      background: rgba(255,255,255,.05);
      display:flex; align-items:center; justify-content:center;
    }
    .wj-img img{ width:100%; height:100%; object-fit:cover; display:block; }
    .wj-pad{ padding: 12px 12px 14px; }
    .wj-name{ margin:0; font-weight:800; }
    .wj-empty{
      margin-top:16px;
      padding:14px;
      border:1px dashed rgba(255,255,255,.25);
      border-radius:12px;
      color:#cfcfcf;
      background: rgba(255,255,255,.04);
    }
  </style>
</head>

<body data-path-to-root="./" data-include-products="false" class="u-body u-clearfix u-grey-80 u-xl-mode" data-lang="en">

<header class="u-clearfix u-header u-header" id="header">
  <div class="u-clearfix u-sheet u-sheet-1">
    <a href="./" class="u-image u-logo u-image-1" data-image-width="210" data-image-height="92">
      <img src="images/Woodjiggerygray.svg" class="u-logo-image u-logo-image-1" alt="WoodJiggery">
    </a>

    <nav class="u-menu u-menu-one-level u-offcanvas u-menu-1" role="navigation" aria-label="Menu navigation">
      <div class="menu-collapse" style="font-size: 1rem; letter-spacing: 0px;">
        <a class="u-button-style u-custom-left-right-menu-spacing u-custom-padding-bottom u-custom-top-bottom-menu-spacing u-hamburger-link u-nav-link u-text-active-palette-1-base u-text-hover-palette-2-base"
           href="#" tabindex="-1" aria-label="Open menu" aria-controls="42b1">
          <svg class="u-svg-link" viewBox="0 0 24 24"><use xlink:href="#menu-hamburger"></use></svg>
          <svg class="u-svg-content" version="1.1" id="menu-hamburger" viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg">
            <g><rect y="1" width="16" height="2"></rect><rect y="7" width="16" height="2"></rect><rect y="13" width="16" height="2"></rect></g>
          </svg>
        </a>
      </div>

      <div class="u-custom-menu u-nav-container">
        <ul class="u-nav u-unstyled u-nav-1" role="menubar">
          <li role="none" class="u-nav-item"><a role="menuitem" class="u-button-style u-nav-link u-text-active-palette-1-base u-text-hover-palette-2-base" href="./" style="padding: 10px 20px;">Home</a></li>
          <li role="none" class="u-nav-item"><a role="menuitem" class="u-button-style u-nav-link u-text-active-palette-1-base u-text-hover-palette-2-base" href="Creations.asp" style="padding: 10px 20px;">Creations</a></li>
          <li role="none" class="u-nav-item"><a role="menuitem" class="u-button-style u-nav-link u-text-active-palette-1-base u-text-hover-palette-2-base" href="About.html" style="padding: 10px 20px;">About</a></li>
          <li role="none" class="u-nav-item"><a role="menuitem" class="u-button-style u-nav-link u-text-active-palette-1-base u-text-hover-palette-2-base" href="Contact.html" style="padding: 10px 20px;">Contact</a></li>
        </ul>
      </div>

      <div class="u-custom-menu u-nav-container-collapse" id="42b1" role="region" aria-label="Menu panel">
        <div class="u-black u-container-style u-inner-container-layout u-opacity u-opacity-95 u-sidenav">
          <div class="u-inner-container-layout u-sidenav-overflow">
            <div class="u-menu-close" tabindex="-1" aria-label="Close menu"></div>
            <ul class="u-align-center u-nav u-popupmenu-items u-unstyled u-nav-2" role="menubar">
              <li role="none" class="u-nav-item"><a role="menuitem" class="u-button-style u-nav-link" href="./">Home</a></li>
              <!-- changed from Creations.html to Creations.asp -->
              <li role="none" class="u-nav-item"><a role="menuitem" class="u-button-style u-nav-link" href="Creations.asp">Creations</a></li>
              <li role="none" class="u-nav-item"><a role="menuitem" class="u-button-style u-nav-link" href="About.html">About</a></li>
              <li role="none" class="u-nav-item"><a role="menuitem" class="u-button-style u-nav-link" href="Contact.html">Contact</a></li>
            </ul>
          </div>
        </div>
        <div class="u-black u-menu-overlay u-opacity u-opacity-70"></div>
      </div>
    </nav>
  </div>
</header>

<section class="u-clearfix u-section-1" id="creations">
  <div class="u-clearfix u-sheet u-sheet-1">
    <div class="wj-crumb"><a href="./" style="color:#cfcfcf;">Home</a> &nbsp;›&nbsp; Creations</div>
    <h1 class="u-text u-text-default wj-title">Creations</h1>
    <div class="wj-sub">Pick a category to browse products.</div>

    <% If rs.EOF Then %>
      <div class="wj-empty">No categories yet.</div>
    <% Else %>
      <div class="wj-grid">
        <% Do While Not rs.EOF %>
          <%
            Dim catName, catSlug, imgPath
            catName = rs("Name")
            catSlug = rs("Slug")
            imgPath = ""
            If Not IsNull(rs("ImagePath")) Then imgPath = rs("ImagePath")
          %>
          <a class="wj-card" href="category.asp?slug=<%=Server.URLEncode(catSlug)%>">
            <div class="wj-img">
              <% If Len(imgPath) > 0 Then %>
                <img src="<%=H(imgPath)%>" alt="<%=H(catName)%>">
              <% Else %>
                <span style="color:#cfcfcf;font-size:13px;">No image yet</span>
              <% End If %>
            </div>
            <div class="wj-pad">
              <p class="wj-name"><%=H(catName)%></p>
            </div>
          </a>
          <%
            rs.MoveNext
          Loop
          %>
      </div>
    <% End If %>
  </div>
</section>

<footer class="u-align-center u-clearfix u-container-align-center u-footer u-grey-80 u-footer" id="footer">
  <div class="u-clearfix u-sheet u-sheet-1">
    <div class="u-list u-list-1">
      <div class="u-repeater u-repeater-1">
        <div class="u-container-style u-list-item u-repeater-item">
          <div class="u-container-layout u-similar-container u-container-layout-1">
            <a href="./" class="u-border-1 u-border-active-palette-2-base u-border-hover-palette-1-base u-border-no-left u-border-no-right u-border-no-top u-btn u-button-style u-custom-item u-none u-text-palette-1-base u-btn-1" title="Home">Home</a>
          </div>
        </div>
        <div class="u-container-style u-list-item u-repeater-item">
          <div class="u-container-layout u-similar-container u-container-layout-2">
            <!-- changed from Creations.html to Creations.asp -->
            <a href="Creations.asp" class="u-active-none u-border-2 u-border-no-left u-border-no-right u-border-no-top u-border-palette-1-base u-btn u-button-style u-custom-item u-hover-none u-none u-btn-2" title="Creations">Creations</a>
          </div>
        </div>
        <div class="u-container-style u-list-item u-repeater-item">
          <div class="u-container-layout u-similar-container u-container-layout-3">
            <a href="About.html" class="u-active-none u-border-2 u-border-no-left u-border-no-right u-border-no-top u-border-palette-1-base u-btn u-button-style u-custom-item u-hover-none u-none u-btn-3" title="About">About</a>
          </div>
        </div>
        <div class="u-container-style u-list-item u-repeater-item">
          <div class="u-container-layout u-similar-container u-container-layout-4">
            <a href="Contact.html" class="u-active-none u-border-2 u-border-no-left u-border-no-right u-border-no-top u-border-palette-1-base u-btn u-button-style u-custom-item u-hover-none u-none u-btn-4" title="Contact">Contact</a>
          </div>
        </div>
      </div>
    </div>
  </div>
</footer>

</body>
</html>
<%
rs.Close : Set rs = Nothing
conn.Close : Set conn = Nothing
%>