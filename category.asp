<%
Option Explicit
Response.Buffer = True

Function H(s) : H = Server.HTMLEncode("" & s) : End Function

Function JsonEsc(ByVal s)
  s = "" & s
  s = Replace(s, "\", "\\")
  s = Replace(s, Chr(34), "\" & Chr(34))
  s = Replace(s, vbCrLf, "\n")
  s = Replace(s, vbCr, "\n")
  s = Replace(s, vbLf, "\n")
  JsonEsc = s
End Function
%>
<!--#include virtual="/includes/adovbs.asp"-->
<!--#include virtual="/includes/datacon.asp"-->
<%
Dim conn : Set conn = OpenConn()

Dim slug : slug = Trim(Request.QueryString("slug"))
If Len(slug) = 0 Then
  Response.Write "Missing category slug."
  Response.End
End If

'--- Load category ---
Dim catId, catName

Dim cmdC : Set cmdC = Server.CreateObject("ADODB.Command")
Set cmdC.ActiveConnection = conn
cmdC.CommandType = adCmdText
cmdC.CommandText = "SELECT CategoryID, Name FROM dbo.Categories WHERE IsActive=1 AND Slug=?;"
cmdC.Parameters.Append cmdC.CreateParameter("@Slug", adVarWChar, adParamInput, 100, slug)

Dim rsC : Set rsC = cmdC.Execute
If rsC.EOF Then
  rsC.Close : Set rsC = Nothing
  Set cmdC = Nothing
  conn.Close : Set conn = Nothing
  Response.Write "Category not found."
  Response.End
End If

catId = CLng(rsC("CategoryID"))
catName = rsC("Name")

rsC.Close : Set rsC = Nothing
Set cmdC = Nothing

'--- Load products + primary image ---
Dim cmdP : Set cmdP = Server.CreateObject("ADODB.Command")
Set cmdP.ActiveConnection = conn
cmdP.CommandType = adCmdText
cmdP.CommandText = _
  "SELECT p.ProductID, p.Name, p.ShortDesc, p.LongDesc, p.Price, " & _
  "       img.ImagePath AS PrimaryImage " & _
  "FROM dbo.Products p " & _
  "OUTER APPLY (" & _
  "  SELECT TOP 1 ImagePath " & _
  "  FROM dbo.ProductImages i " & _
  "  WHERE i.ProductID = p.ProductID " & _
  "  ORDER BY i.IsPrimary DESC, i.SortOrder, i.ImageID " & _
  ") img " & _
  "WHERE p.IsActive=1 AND p.CategoryID=? " & _
  "ORDER BY p.SortOrder, p.Name;"
cmdP.Parameters.Append cmdP.CreateParameter("@CategoryID", adInteger, adParamInput, , CInt(catId))

Dim rs : Set rs = cmdP.Execute

Dim pidList : pidList = ""
Dim metaJson : metaJson = ""
%>
<!DOCTYPE html>
<html style="font-size: 16px;" lang="en">
<head>
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta charset="utf-8">
  <title><%=H(catName)%> - Creations</title>

  <!-- Keep Nicepage base styling + behavior -->
  <link rel="stylesheet" href="nicepage.css" media="screen">
  <script class="u-script" type="text/javascript" src="jquery.js" defer=""></script>
  <script class="u-script" type="text/javascript" src="nicepage.js" defer=""></script>

  <link id="u-page-google-font" rel="stylesheet"
        href="https://fonts.googleapis.com/css2?display=swap&family=Roboto:ital,wght@0,100;0,200;0,300;0,400;0,500;0,600;0,700;0,800;0,900;1,100;1,200;1,300;1,400;1,500;1,600;1,700;1,800;1,900&family=Open+Sans:ital,wght@0,300;0,400;0,500;0,600;0,700;0,800;1,300;1,400;1,500;1,600;1,700;1,800">

  <style>
    /* Catalog styling (scoped) */
    .wj-crumb { color: #cfcfcf; margin: 10px 0 0; font-size: 14px; }
    .wj-sub   { color: #cfcfcf; margin: 6px 0 16px; }

    .wj-grid{
      display:grid;
      grid-template-columns:repeat(auto-fill, minmax(240px, 1fr));
      gap:16px;
      margin-top:16px;
    }
    .wj-card{
      background: rgba(255,255,255,.06);
      border: 1px solid rgba(255,255,255,.14);
      border-radius: 14px;
      overflow:hidden;
      box-shadow:0 1px 0 rgba(0,0,0,.02);
      cursor:pointer;
      transition: transform .12s ease, background .12s ease;
      outline: none;
    }
    .wj-card:hover{ transform: translateY(-2px); background: rgba(255,255,255,.09); }
    .wj-img{
      height:170px;
      background: rgba(255,255,255,.05);
      display:flex; align-items:center; justify-content:center;
    }
    .wj-img img{ width:100%; height:100%; object-fit:cover; display:block; }
    .wj-pad{ padding: 12px 12px 14px; }
    .wj-name{ margin:0 0 6px; font-weight:800; }
    .wj-desc{ margin:0 0 10px; color:#e0e0e0; font-size:14px; min-height:38px; }
    .wj-meta{ display:flex; justify-content:space-between; align-items:center; gap:10px; }
    .wj-price{ font-weight:800; color:#fff; }
    .wj-hint{ color:#cfcfcf; font-size:13px; }

    .wj-empty{
      margin-top:16px;
      padding:14px;
      border:1px dashed rgba(255,255,255,.25);
      border-radius:12px;
      color:#cfcfcf;
      background: rgba(255,255,255,.04);
    }

    /* Lightbox */
    .wj-lb{position:fixed; inset:0; display:none; z-index:9999;}
    .wj-lb.on{display:block;}
    .wj-lb-backdrop{position:absolute; inset:0; background:rgba(0,0,0,.75);}
    .wj-lb-panel{position:relative; max-width:980px; margin:40px auto; background:#111; border-radius:14px; overflow:hidden; border:1px solid rgba(255,255,255,.15);}
    .wj-lb-top{display:flex; justify-content:space-between; align-items:center; padding:10px 12px; background:rgba(0,0,0,.35); color:#fff;}
    .wj-lb-title{font-weight:700;}
    .wj-lb-close{background:transparent; border:0; color:#fff; font-size:22px; cursor:pointer; padding:4px 8px;}
    .wj-lb-main{position:relative; background:#000; display:flex; align-items:center; justify-content:center; height:520px;}
    .wj-lb-main img{max-width:100%; max-height:100%; display:block;}
    .wj-lb-nav{position:absolute; top:0; bottom:0; width:20%; display:flex; align-items:center;}
    .wj-lb-prev{left:0; justify-content:flex-start;}
    .wj-lb-next{right:0; justify-content:flex-end;}
    .wj-lb-nav button{background:rgba(0,0,0,.35); border:1px solid rgba(255,255,255,.2); color:#fff; font-size:20px; cursor:pointer; padding:10px 12px; border-radius:12px; margin:0 10px;}
    .wj-lb-cap{padding:10px 12px; background:#111; color:#ddd; font-size:14px; display:flex; justify-content:space-between; gap:12px;}
    .wj-lb-thumbs{display:flex; gap:8px; overflow:auto; padding:10px 12px; background:#0b0b0b; border-top:1px solid rgba(255,255,255,.08);}
    .wj-lb-thumbs img{height:64px; width:88px; object-fit:cover; border-radius:10px; border:2px solid transparent; cursor:pointer; opacity:.85;}
    .wj-lb-thumbs img.sel{border-color:#fff; opacity:1;}

    @media (max-width:700px){
      .wj-lb-panel{margin:14px;}
      .wj-lb-main{height:360px;}
      .wj-lb-nav{width:30%;}
    }
  </style>
</head>

<body data-path-to-root="./" data-include-products="false" class="u-body u-clearfix u-grey-80 u-xl-mode" data-lang="en">

<!-- Header copied from Home structure (with .asp links) -->
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

<section class="u-clearfix u-section-1" id="category">
  <div class="u-clearfix u-sheet u-sheet-1">
    <div class="wj-crumb">
      <a href="./" style="color:#cfcfcf;">Home</a> &nbsp;›&nbsp;
      <a href="Creations.asp" style="color:#cfcfcf;">Creations</a> &nbsp;›&nbsp;
      <span><%=H(catName)%></span>
    </div>

    <h1 class="u-text u-text-default" style="margin-top:10px;"><%=H(catName)%></h1>
    <div class="wj-sub">Click a product to view images.</div>

    <% If rs.EOF Then %>
      <div class="wj-empty">No products yet in this category.</div>
    <% Else %>
      <div class="wj-grid">
        <% Do While Not rs.EOF %>
          <%
            Dim pid, pName, pShort, pPriceDisp, prim
            pid = CLng(rs("ProductID"))
            pName = rs("Name")
            pShort = ""
            If Not IsNull(rs("ShortDesc")) Then pShort = rs("ShortDesc")

            pPriceDisp = ""
            If Not IsNull(rs("Price")) Then pPriceDisp = "$" & FormatNumber(rs("Price"), 2, -1, 0, -1)

            prim = ""
            If Not IsNull(rs("PrimaryImage")) Then prim = rs("PrimaryImage")

            If Len(pidList) > 0 Then pidList = pidList & ","
            pidList = pidList & pid

            Dim pLong
            pLong = ""
            If Not IsNull(rs("LongDesc")) Then pLong = rs("LongDesc")

            If Len(metaJson) > 0 Then metaJson = metaJson & ","
            metaJson = metaJson & Chr(34) & pid & Chr(34) & ":{""name"":""" & JsonEsc(pName) & """,""price"":""" & JsonEsc(pPriceDisp) & """,""primary"":""" & JsonEsc(prim) & """,""long"":""" & JsonEsc(pLong) & """}"
          %>

          <div class="wj-card" data-pid="<%=pid%>" tabindex="0" role="button" aria-label="View <%=H(pName)%>">
            <div class="wj-img">
              <% If Len(prim) > 0 Then %>
                <img src="<%=H(prim)%>" alt="<%=H(pName)%>">
              <% Else %>
                <span style="color:#cfcfcf;font-size:13px;">No image</span>
              <% End If %>
            </div>
            <div class="wj-pad">
              <p class="wj-name"><%=H(pName)%></p>
              <p class="wj-desc"><%=H(pShort)%></p>
              <div class="wj-meta">
                <div class="wj-price"><%=H(pPriceDisp)%></div>
                <div class="wj-hint">Click to view</div>
              </div>
            </div>
          </div>

        <% rs.MoveNext : Loop %>
      </div>
    <% End If %>
  </div>
</section>

<!-- Lightbox -->
<div class="wj-lb" id="lb">
  <div class="wj-lb-backdrop" id="lbBack"></div>
  <div class="wj-lb-panel" role="dialog" aria-modal="true">
    <div class="wj-lb-top">
      <div class="wj-lb-title" id="lbTitle">Preview</div>
      <button class="wj-lb-close" id="lbClose" aria-label="Close">×</button>
    </div>
    <div class="wj-lb-main">
      <div class="wj-lb-nav wj-lb-prev"><button id="lbPrev">‹</button></div>
      <img id="lbImg" src="" alt="">
      <div class="wj-lb-nav wj-lb-next"><button id="lbNext">›</button></div>
    </div>
    <div class="wj-lb-cap">
      <div style="flex:1; min-width:0;">
        <div id="lbShort" style="font-weight:700;"></div>
        <div id="lbLong" style="margin-top:6px; color:#ddd; white-space:pre-wrap;"></div>
        <div id="lbAlt" style="margin-top:6px; color:#aaa;"></div>
      </div>
      <div id="lbPrice" style="font-weight:800; white-space:nowrap;"></div>
    </div>
    <div class="wj-lb-thumbs" id="lbThumbs"></div>
  </div>
</div>

<%
'--- Pull all images for products on this category page ---
Dim imagesJson : imagesJson = ""

If Len(pidList) > 0 Then
  Dim sqlI
  sqlI = ""
  sqlI = sqlI & "SELECT ProductID, ImagePath, AltText, SortOrder, IsPrimary, ImageID "
  sqlI = sqlI & "FROM dbo.ProductImages "
  sqlI = sqlI & "WHERE ProductID IN (" & pidList & ") "
  sqlI = sqlI & "ORDER BY ProductID, IsPrimary DESC, SortOrder, ImageID;"

  Dim rsI : Set rsI = conn.Execute(sqlI)

  Dim curPid : curPid = -1
  Dim firstInPid : firstInPid = True

  Do While Not rsI.EOF
    Dim ipid, ipath, ialt
    ipid = CLng(rsI("ProductID"))
    ipath = rsI("ImagePath")
    ialt = ""
    If Not IsNull(rsI("AltText")) Then ialt = rsI("AltText")

    If ipid <> curPid Then
      If curPid <> -1 Then imagesJson = imagesJson & "],"
      imagesJson = imagesJson & Chr(34) & ipid & Chr(34) & ":["
      curPid = ipid
      firstInPid = True
    End If

    If Not firstInPid Then imagesJson = imagesJson & ","
    imagesJson = imagesJson & "{""src"":""" & JsonEsc(ipath) & """,""alt"":""" & JsonEsc(ialt) & """}"
    firstInPid = False

    rsI.MoveNext
  Loop

  If curPid <> -1 Then imagesJson = imagesJson & "]"

  rsI.Close : Set rsI = Nothing
End If
%>

<script>
  const PRODUCT_META   = { <%=metaJson%> };
  const PRODUCT_IMAGES = { <%=imagesJson%> };

  const lb = document.getElementById('lb');
  const lbBack = document.getElementById('lbBack');
  const lbClose = document.getElementById('lbClose');
  const lbImg = document.getElementById('lbImg');
  const lbTitle = document.getElementById('lbTitle');
  const lbAlt = document.getElementById('lbAlt');
  const lbPrice = document.getElementById('lbPrice');
  const lbThumbs = document.getElementById('lbThumbs');
  const lbPrev = document.getElementById('lbPrev');
  const lbNext = document.getElementById('lbNext');

  let curPid = null;
  let curIndex = 0;
  let curList = [];

  function openLb(pid){
    curPid = String(pid);
    const meta = PRODUCT_META[curPid] || {name:"", price:"", primary:""};
    lbShort.textContent = meta.short || "";
    lbLong.textContent  = meta.long  || "";

    curList = PRODUCT_IMAGES[curPid] || [];

    // fallback: show primary image if no image rows yet
    if (!curList.length && meta.primary) curList = [{src: meta.primary, alt: ""}];

    curIndex = 0;
    lbTitle.textContent = meta.name || "Preview";
    lbPrice.textContent = meta.price || "";
    renderLb();
    lb.classList.add('on');
    document.body.style.overflow = "hidden";
  }

  function closeLb(){
    lb.classList.remove('on');
    document.body.style.overflow = "";
    curPid = null;
    curList = [];
    curIndex = 0;
  }

  function renderLb(){
    if (!curList.length){
      lbImg.src = "";
      lbAlt.textContent = "No images for this product yet.";
      lbThumbs.innerHTML = "";
      return;
    }
    const item = curList[curIndex];
    lbImg.src = item.src;
    lbAlt.textContent = item.alt || "";

    lbThumbs.innerHTML = "";
    curList.forEach((it, idx) => {
      const t = document.createElement('img');
      t.src = it.src;
      if (idx === curIndex) t.classList.add('sel');
      t.addEventListener('click', () => { curIndex = idx; renderLb(); });
      lbThumbs.appendChild(t);
    });
  }

  function prevImg(){
    if (!curList.length) return;
    curIndex = (curIndex - 1 + curList.length) % curList.length;
    renderLb();
  }
  function nextImg(){
    if (!curList.length) return;
    curIndex = (curIndex + 1) % curList.length;
    renderLb();
  }

  document.querySelectorAll('.wj-card[data-pid]').forEach(card => {
    card.addEventListener('click', () => openLb(card.dataset.pid));
    card.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openLb(card.dataset.pid); }
    });
  });

  lbBack.addEventListener('click', closeLb);
  lbClose.addEventListener('click', closeLb);
  lbPrev.addEventListener('click', prevImg);
  lbNext.addEventListener('click', nextImg);

  document.addEventListener('keydown', (e) => {
    if (!lb.classList.contains('on')) return;
    if (e.key === 'Escape') closeLb();
    if (e.key === 'ArrowLeft') prevImg();
    if (e.key === 'ArrowRight') nextImg();
  });
</script>

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
Set cmdP = Nothing
conn.Close : Set conn = Nothing
%>