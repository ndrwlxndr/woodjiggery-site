<%
Option Explicit
Response.Buffer = True
%>
<!--#include virtual="/includes/adovbs.asp"-->
<!--#include virtual="/includes/datacon.asp"-->
<%
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

Dim conn : Set conn = OpenConn()

Dim slug : slug = Trim(Request.QueryString("slug"))
If Len(slug) = 0 Then
  Response.Write "Missing category slug."
  Response.End
End If

' Category
Dim catId, catName
Dim cmdC : Set cmdC = Server.CreateObject("ADODB.Command")
Set cmdC.ActiveConnection = conn
cmdC.CommandType = adCmdText
cmdC.CommandText = "SELECT CategoryID, Name FROM dbo.Categories WHERE IsActive=1 AND Slug=?;"
cmdC.Parameters.Append cmdC.CreateParameter("@Slug", adVarWChar, adParamInput, 100, slug)

Dim rsC : Set rsC = cmdC.Execute
If rsC.EOF Then
  Response.Write "Category not found."
  Response.End
End If
catId = CLng(rsC("CategoryID"))
catName = rsC("Name")
rsC.Close : Set rsC = Nothing
Set cmdC = Nothing

' Products + primary image
Dim cmdP : Set cmdP = Server.CreateObject("ADODB.Command")
Set cmdP.ActiveConnection = conn
cmdP.CommandType = adCmdText
cmdP.CommandText = _
  "SELECT p.ProductID, p.Name, p.ShortDesc, p.Price, " & _
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

' Collect product IDs for image query + build meta JSON as we output cards
Dim pidList : pidList = ""
Dim metaJson : metaJson = ""
%>
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title><%=H(catName)%> - WoodJiggery</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <style>
    body{font-family:Segoe UI,Arial; margin:0; padding:24px; background:#fafafa; color:#111;}
    .wrap{max-width:1100px; margin:0 auto;}
    .crumb{color:#666; font-size:14px;}
    .grid{display:grid; grid-template-columns:repeat(auto-fill, minmax(240px, 1fr)); gap:16px; margin-top:18px;}
    .card{background:#fff; border:1px solid #e2e2e2; border-radius:14px; overflow:hidden; box-shadow:0 1px 0 rgba(0,0,0,.02); cursor:pointer;}
    .img{height:170px; background:#f0f0f0; display:flex; align-items:center; justify-content:center;}
    .img img{width:100%; height:100%; object-fit:cover; display:block;}
    .pad{padding:12px 12px 14px;}
    .name{font-weight:800; margin:0 0 6px;}
    .desc{margin:0 0 10px; color:#444; font-size:14px; min-height:38px;}
    .price{font-weight:800;}
    .muted{color:#666;}
    .empty{color:#666; margin-top:18px; background:#fff; border:1px dashed #ccc; border-radius:12px; padding:16px;}

    /* Lightbox */
    .lb{position:fixed; inset:0; display:none; z-index:9999;}
    .lb.on{display:block;}
    .lb-backdrop{position:absolute; inset:0; background:rgba(0,0,0,.75);}
    .lb-panel{position:relative; max-width:980px; margin:40px auto; background:#111; border-radius:14px; overflow:hidden; border:1px solid rgba(255,255,255,.15);}
    .lb-top{display:flex; justify-content:space-between; align-items:center; padding:10px 12px; background:rgba(0,0,0,.35); color:#fff;}
    .lb-title{font-weight:700;}
    .lb-close{background:transparent; border:0; color:#fff; font-size:22px; cursor:pointer; padding:4px 8px;}
    .lb-main{position:relative; background:#000; display:flex; align-items:center; justify-content:center; height:520px;}
    .lb-main img{max-width:100%; max-height:100%; display:block;}
    .lb-nav{position:absolute; top:0; bottom:0; width:20%; display:flex; align-items:center;}
    .lb-prev{left:0; justify-content:flex-start;}
    .lb-next{right:0; justify-content:flex-end;}
    .lb-nav button{background:rgba(0,0,0,.35); border:1px solid rgba(255,255,255,.2); color:#fff; font-size:20px; cursor:pointer; padding:10px 12px; border-radius:12px; margin:0 10px;}
    .lb-cap{padding:10px 12px; background:#111; color:#ddd; font-size:14px; display:flex; justify-content:space-between; gap:12px;}
    .lb-thumbs{display:flex; gap:8px; overflow:auto; padding:10px 12px; background:#0b0b0b; border-top:1px solid rgba(255,255,255,.08);}
    .lb-thumbs img{height:64px; width:88px; object-fit:cover; border-radius:10px; border:2px solid transparent; cursor:pointer; opacity:.85;}
    .lb-thumbs img.sel{border-color:#fff; opacity:1;}
    @media (max-width: 700px){
      .lb-panel{margin:14px;}
      .lb-main{height:360px;}
      .lb-nav{width:30%;}
    }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="crumb"><a href="/index.asp">Home</a> &nbsp;›&nbsp; <a href="/creations.asp">Creations</a></div>
    <h1 style="margin:8px 0 0;"><%=H(catName)%></h1>

    <% If rs.EOF Then %>
      <div class="empty">No products yet in this category.</div>
    <% Else %>
      <div class="grid">
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

            If Len(metaJson) > 0 Then metaJson = metaJson & ","
            metaJson = metaJson & Chr(34) & pid & Chr(34) & ":{""name"":""" & JsonEsc(pName) & """,""price"":""" & JsonEsc(pPriceDisp) & """,""primary"":""" & JsonEsc(prim) & """}"
          %>

          <div class="card" data-pid="<%=pid%>">
            <div class="img">
              <% If Len(prim) > 0 Then %>
                <img src="<%=H(prim)%>" alt="<%=H(pName)%>">
              <% Else %>
                <span class="muted">No image</span>
              <% End If %>
            </div>
            <div class="pad">
              <p class="name"><%=H(pName)%></p>
              <p class="desc"><%=H(pShort)%></p>
              <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;">
                <div class="price"><%=H(pPriceDisp)%></div>
                <div class="muted" style="font-size:13px;">Click to view</div>
              </div>
            </div>
          </div>

        <% rs.MoveNext : Loop %>
      </div>
    <% End If %>
  </div>

  <!-- Lightbox -->
  <div class="lb" id="lb">
    <div class="lb-backdrop" id="lbBack"></div>
    <div class="lb-panel" role="dialog" aria-modal="true">
      <div class="lb-top">
        <div class="lb-title" id="lbTitle">Preview</div>
        <button class="lb-close" id="lbClose" aria-label="Close">×</button>
      </div>
      <div class="lb-main">
        <div class="lb-nav lb-prev"><button id="lbPrev">‹</button></div>
        <img id="lbImg" src="" alt="">
        <div class="lb-nav lb-next"><button id="lbNext">›</button></div>
      </div>
      <div class="lb-cap">
        <div id="lbAlt"></div>
        <div id="lbPrice" style="font-weight:800;"></div>
      </div>
      <div class="lb-thumbs" id="lbThumbs"></div>
    </div>
  </div>

<%
' Pull all images for products on this category page
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
  // Product meta + images loaded from server-side JSON
  const PRODUCT_META = { <%=metaJson%> };
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
    curList = PRODUCT_IMAGES[curPid] || [];
    // If no image rows yet, fall back to primary image (if any)
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

    // thumbs
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

  document.querySelectorAll('.card[data-pid]').forEach(card => {
    card.addEventListener('click', () => openLb(card.dataset.pid));
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

</body>
</html>
<%
rs.Close : Set rs = Nothing
Set cmdP = Nothing
conn.Close : Set conn = Nothing
%>
