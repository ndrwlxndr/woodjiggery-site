<%
' /includes/upload.asp
' Simple upload/parser for Classic ASP
' - Works for multipart/form-data AND application/x-www-form-urlencoded
' - Exposes: up.Form("x"), up.Files("imgFile")
' - Safe binary save via ADODB.Stream

Class Upload
  Private mForm, mFiles

  Public Property Get Form()
    Set Form = mForm
  End Property

  Public Property Get Files()
    Set Files = mFiles
  End Property

  Private Sub Class_Initialize()
    Set mForm = Server.CreateObject("Scripting.Dictionary")
    mForm.CompareMode = 1 ' vbTextCompare

    Set mFiles = Server.CreateObject("Scripting.Dictionary")
    mFiles.CompareMode = 1 ' vbTextCompare

    ParseRequest
  End Sub

Private Sub ParseRequest()
  Dim total, ct
  total = Request.TotalBytes
  If total <= 0 Then Exit Sub

  ct = LCase("" & Request.ServerVariables("CONTENT_TYPE"))

  If InStr(ct, "multipart/form-data") > 0 And InStr(ct, "boundary=") > 0 Then
    Dim bin
    bin = Request.BinaryRead(total)   ' ONLY for multipart uploads
    ParseMultipart bin, ct
  Else
    ' Normal POST: do NOT BinaryRead
    Dim k
    For Each k In Request.Form
      mForm(k) = Request.Form(k)
    Next
  End If
End Sub

' Called when the client-side JS interceptor encoded the file as base64
' and posted everything as application/x-www-form-urlencoded.
Private Sub ParseBase64Upload(ByVal fileFieldName)
  Dim k
  For Each k In Request.Form
    If Left(k, 4) <> "_b64" Then
      mForm(k) = Request.Form(k)
    End If
  Next

  Dim b64Data : b64Data = Trim("" & Request.Form("_b64data"))
  Dim b64Name : b64Name = Trim("" & Request.Form("_b64name"))
  Dim b64Type : b64Type = Trim("" & Request.Form("_b64type"))

  If Len(b64Data) > 0 And Len(b64Name) > 0 Then
    Dim blob : blob = Base64DecodeBytes(b64Data)
    Dim fileObj : Set fileObj = New UploadFile
    fileObj.Name = fileFieldName
    fileObj.FileName = BaseFileName(b64Name)
    fileObj.ContentType = b64Type
    fileObj.Blob = blob
    mFiles(fileFieldName) = fileObj
  End If
End Sub

' Decode a base64 string to a binary byte array using MSXML
Private Function Base64DecodeBytes(ByVal sBase64)
  Dim oDoc, oNode
  Set oDoc = Server.CreateObject("MSXML2.DOMDocument")
  Set oNode = oDoc.createElement("b64")
  oNode.dataType = "bin.base64"
  oNode.text = sBase64
  Base64DecodeBytes = oNode.nodeTypedValue
  Set oNode = Nothing
  Set oDoc = Nothing
End Function

  Private Sub ParseUrlEncoded(bin)
    Dim s, parts, i, kv, k, v, p
    s = BStr2UStr(bin)
    If Len(s) = 0 Then Exit Sub

    parts = Split(s, "&")
    For i = 0 To UBound(parts)
      If Len(parts(i)) > 0 Then
        p = InStr(parts(i), "=")
        If p > 0 Then
          k = URLDecode(Left(parts(i), p-1))
          v = URLDecode(Mid(parts(i), p+1))
        Else
          k = URLDecode(parts(i))
          v = ""
        End If
        If Len(k) > 0 Then mForm(k) = v
      End If
    Next
  End Sub

  Private Sub ParseMultipart(bin, ct)
    Dim boundary, bBoundary, bCRLF, bCRLFCRLF
    Dim pos, nextPos, partStart, headerEnd, headers, bodyStart, bodyEnd
    Dim name, filename, ctype, dispo, fileObj, blob

    boundary = ExtractBoundary(ct)
    If Len(boundary) = 0 Then Exit Sub

    bBoundary   = UStr2BStr("--" & boundary)
    bCRLF       = UStr2BStr(vbCrLf)
    bCRLFCRLF   = UStr2BStr(vbCrLf & vbCrLf)

    pos = InStrB(1, bin, bBoundary)
    Do While pos > 0
      ' end marker?
      If MidB(bin, pos + LenB(bBoundary), 2) = UStr2BStr("--") Then Exit Do

      partStart = pos + LenB(bBoundary) + LenB(bCRLF)

      headerEnd = InStrB(partStart, bin, bCRLFCRLF)
      If headerEnd = 0 Then Exit Do

      headers = BStr2UStr(MidB(bin, partStart, headerEnd - partStart))

      bodyStart = headerEnd + LenB(bCRLFCRLF)
      nextPos = InStrB(bodyStart, bin, bBoundary)
      If nextPos = 0 Then Exit Do

      bodyEnd = nextPos - LenB(bCRLF)
      If bodyEnd < bodyStart Then bodyEnd = bodyStart

      name = ExtractBetween(headers, "name=""", """")
      filename = ExtractBetween(headers, "filename=""", """")

      ctype = ""
      Dim ctPos : ctPos = InStr(1, headers, "Content-Type:", vbTextCompare)
      If ctPos > 0 Then
        Dim lineEnd : lineEnd = InStr(ctPos, headers, vbCrLf)
        If lineEnd = 0 Then lineEnd = Len(headers) + 1
        ctype = Trim(Mid(headers, ctPos + Len("Content-Type:"), lineEnd - (ctPos + Len("Content-Type:"))))
      End If

      If Len(filename) > 0 Then
        blob = MidB(bin, bodyStart, bodyEnd - bodyStart)
        If LenB(blob) > 0 Then
          Set fileObj = New UploadFile
          fileObj.Name = name
          fileObj.FileName = BaseFileName(filename)
          fileObj.ContentType = ctype
          fileObj.Blob = blob
          mFiles(name) = fileObj
        End If
      Else
        Dim val : val = BStr2UStr(MidB(bin, bodyStart, bodyEnd - bodyStart))
        If Len(name) > 0 Then mForm(name) = val
      End If

      pos = nextPos
    Loop
  End Sub

  Private Function ExtractBoundary(ct)
    Dim p, b
    p = InStr(ct, "boundary=")
    If p = 0 Then ExtractBoundary = "" : Exit Function
    b = Mid(ct, p + Len("boundary="))
    b = Replace(b, """", "")
    ExtractBoundary = b
  End Function

  Private Function ExtractBetween(s, a, b)
    Dim p1, p2
    p1 = InStr(1, s, a, vbTextCompare)
    If p1 = 0 Then ExtractBetween = "" : Exit Function
    p1 = p1 + Len(a)
    p2 = InStr(p1, s, b, vbTextCompare)
    If p2 = 0 Then ExtractBetween = "" : Exit Function
    ExtractBetween = Mid(s, p1, p2 - p1)
  End Function

  Private Function BaseFileName(fullPath)
    Dim p : p = InStrRev(fullPath, "\")
    If p > 0 Then
      BaseFileName = Mid(fullPath, p + 1)
    Else
      BaseFileName = fullPath
    End If
  End Function

  Private Function URLDecode(Expression)
    Dim strSource, strTemp, strResult, i
    strSource = Replace(Expression, "+", " ")
    strResult = ""
    i = 1
    Do While i <= Len(strSource)
      strTemp = Mid(strSource, i, 1)
      If strTemp = "%" And i + 2 <= Len(strSource) Then
        strResult = strResult & Chr(CInt("&H" & Mid(strSource, i + 1, 2)))
        i = i + 3
      Else
        strResult = strResult & strTemp
        i = i + 1
      End If
    Loop
    URLDecode = strResult
  End Function

  Private Function BStr2UStr(BStr)
    Dim i, out
    out = ""
    For i = 1 To LenB(BStr)
      out = out & Chr(AscB(MidB(BStr, i, 1)))
    Next
    BStr2UStr = out
  End Function

  Private Function UStr2BStr(UStr)
    Dim i, out, ch
    out = ""
    For i = 1 To Len(UStr)
      ch = Mid(UStr, i, 1)
      out = out & ChrB(AscB(ch))
    Next
    UStr2BStr = out
  End Function
End Class


Class UploadFile
  Private mName, mFileName, mContentType, mBlob

  Public Property Get Name(): Name = mName: End Property
  Public Property Let Name(v): mName = v: End Property

  Public Property Get FileName(): FileName = mFileName: End Property
  Public Property Let FileName(v): mFileName = v: End Property

  Public Property Get ContentType(): ContentType = mContentType: End Property
  Public Property Let ContentType(v): mContentType = v: End Property

  Public Property Get Blob(): Blob = mBlob: End Property
  Public Property Let Blob(v): mBlob = v: End Property

' In Class UploadFile

Public Sub SaveTo(ByVal physFolder)
  SaveToAs physFolder, mFileName
End Sub

Public Sub SaveToAs(ByVal physFolder, ByVal newFileName)
  Dim fso, stm, filePath, finalName

  finalName = Trim("" & newFileName)
  If Len(finalName) = 0 Then finalName = mFileName

  Set fso = Server.CreateObject("Scripting.FileSystemObject")
  If Not fso.FolderExists(physFolder) Then fso.CreateFolder physFolder
  filePath = fso.BuildPath(physFolder, finalName)

  Set stm = Server.CreateObject("ADODB.Stream")
  stm.Type = 1 ' adTypeBinary
  stm.Open
  stm.Write mBlob
  stm.SaveToFile filePath, 2 ' overwrite
  stm.Close

  Set stm = Nothing
  Set fso = Nothing
End Sub
End Class
%>