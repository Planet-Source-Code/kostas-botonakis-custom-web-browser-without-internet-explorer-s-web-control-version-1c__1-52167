Attribute VB_Name = "modTags"
Public BackURL As String
Public Back As Boolean
Public DocTitle As String
Public DocURL As String
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_PASTE = &H302
Public Const SW_SHOW = 5
Public Const SWP_SHOWWINDOW = &H40
Public Const EM_GETLINECOUNT = &HBA
Function HTMLText(Text As String)
'Replace html characters with known characters
Text = Replace(Text, "&nbsp;", " ")
Text = Replace(Text, "&amp;", "&")
Text = Replace(Text, "&reg;", "®")
Select Case LCase(CurrentTag)
    'Load Page's title

    Case "td"
        Con = Trim(Text)
        frm.Page.TextRTF = frm.Page.TextRTF & "{\par\trowd\trgaph108\trleft-108\trbrdrt\brdrs\brdrw20 \trbrdrl\brdrs\brdrw20 \trbrdrb\brdrs\brdrw20 \trbrdrr\brdrs\brdrw20 \clbrdrt\brdrw15\brdrs\clbrdrl\brdrw15\brdrs\clbrdrb\brdrw15\brdrs\clbrdrr\brdrw15\brdrs \cellx8748\pard\intbl " & Con & "\cell\row\pard\nowidctlpar\par}"
    'Load Page's text
    Case Else
        frm.Page.SelText = Trim(Text)
        'frm.Page.SelText & Trim(Text)
End Select
End Function
Function PicLoad(picFilenameS As String)
Dim PicFileName
PicFileName = LCase(picFilenameS)
Select Case PicFileName
'Check C
Case Left(PicFileName, 3) = "c:\"
    'If valid Picture's Filename
    frm.Pic.Picture = LoadPicture(PicFileName)
    'Clear the clipboard
    Clipboard.Clear
    'Copy to clipboard the picture of Pictures()
    Clipboard.SetData frm.Pic.Picture
    'Paste the picture into the RichTextBox(Page).
    SendMessage frm.Page.hWnd, WM_PASTE, 0, 0
'Check D
Case Left(PicFileName, 3) = "d:\"
'If valid Picture's Filename
    frm.Pic.Picture = LoadPicture(PicFileName)
    'Clear the clipboard
    Clipboard.Clear
    'Copy to clipboard the picture of Pictures()
    Clipboard.SetData frm.Pic.Picture
    'Paste the picture into the RichTextBox(Page).
    SendMessage frm.Page.hWnd, WM_PASTE, 0, 0
'Check E
Case Left(PicFileName, 3) = "e:\"
'If valid Picture's Filename
    frm.Pic.Picture = LoadPicture(PicFileName)
    'Clear the clipboard
    Clipboard.Clear
    'Copy to clipboard the picture of Pictures()
    Clipboard.SetData frm.Pic.Picture
    'Paste the picture into the RichTextBox(Page).
    SendMessage frm.Page.hWnd, WM_PASTE, 0, 0
Case Left(PicFileName, 3) = "f:\"
'If valid Picture's Filename
    frm.Pic.Picture = LoadPicture(PicFileName)
    'Clear the clipboard
    Clipboard.Clear
    'Copy to clipboard the picture of Pictures()
    Clipboard.SetData frm.Pic.Picture
    'Paste the picture into the RichTextBox(Page).
    SendMessage frm.Page.hWnd, WM_PASTE, 0, 0
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'Check http/ftp
Case Left(PicFileName, 7) = "http://"
'If valid Picture's Filename
    frm.Pic.Picture = LoadPicture(PicFileName)
    'Clear the clipboard
    Clipboard.Clear
    'Copy to clipboard the picture of Pictures()
    Clipboard.SetData frm.Pic.Picture
    'Paste the picture into the RichTextBox(Page).
    SendMessage frm.Page.hWnd, WM_PASTE, 0, 0
Case Left(PicFileName, 8) = "https://"
'If valid Picture's Filename
    frm.Pic.Picture = LoadPicture(PicFileName)
    'Clear the clipboard
    Clipboard.Clear
    'Copy to clipboard the picture of Pictures()
    Clipboard.SetData frm.Pic.Picture
    'Paste the picture into the RichTextBox(Page).
    SendMessage frm.Page.hWnd, WM_PASTE, 0, 0
Case Left(PicFileName, 4) = "www."
'If valid Picture's Filename
    frm.Pic.Picture = LoadPicture(PicFileName)
    'Clear the clipboard
    Clipboard.Clear
    'Copy to clipboard the picture of Pictures()
    Clipboard.SetData frm.Pic.Picture
    'Paste the picture into the RichTextBox(Page).
    SendMessage frm.Page.hWnd, WM_PASTE, 0, 0
Case Left(PicFileName, 6) = "ftp://"
'If valid Picture's Filename
    frm.Pic.Picture = LoadPicture(PicFileName)
    'Clear the clipboard
    Clipboard.Clear
    'Copy to clipboard the picture of Pictures()
    Clipboard.SetData frm.Pic.Picture
    'Paste the picture into the RichTextBox(Page).
    SendMessage frm.Page.hWnd, WM_PASTE, 0, 0
Case Else
    frm.Pic.Picture = LoadPicture(PicFileName)
    'Clear the clipboard
    Clipboard.Clear
    'Copy to clipboard the picture of Pictures()
    Clipboard.SetData frm.Pic.Picture
    'Paste the picture into the RichTextBox(Page).
    SendMessage frm.Page.hWnd, WM_PASTE, 0, 0
End Select
End Function
Public Function LoadURL(strURLs As String)
'Make it with lowser case characters
Dim strURL
strURL = LCase(strURLs)
'Exit if URL is empty
If strURL = "" Then Exit Function
frm.Page.Text = ""
Back = False

'check if strURL = "about:browser"
If strURL = "about:browser" Then
frm.Page.SelStart = 0
frm.Page.TextRTF = ""
Dim K
K = frm.txtAbout.Text
frm.Page.TextRTF = frm.Page.TextRTF & K
End If

'unlock the page to be loaded
frm.Page.Locked = False
'Save the previous doc url
If modTags.DocURL = "" Then
BackURL = ""
Back = False
Else
BackURL = modTags.DocURL
Back = True
End If
'Save the document's url
modTags.DocURL = strURL
'Check URL
Select Case strURL
Case Left(strURL, 7) = "http://"
    Dim HT
    HT = frm.Inet.OpenURL(strURL)
    frm.WebControl1.ParseHTML Trim(HT)
Case Left(strURL, 8) = "https://"
    Dim HTS
    HTS = frm.Inet.OpenURL(strURL)
    frm.WebControl1.ParseHTML Trim(HTS)
Case Left(strURL, 6) = "ftp://"
    Dim FT
    FT = frm.Inet.OpenURL(strURL)
    frm.WebControl1.ParseHTML Trim(FT)
Case Else
    Dim F1
    F1 = FreeFile
    Dim GText As String
    Dim Con As String
    If strURL = "about:browser" Then Exit Function
    If strURL = "error:browser" Then Exit Function
    Open strURL For Input As #F1
    Do While Not EOF(F1)
        Line Input #1, Con
        GText = GText & vbclrf & Con
    Loop
    Close #F1
    frm.WebControl1.ParseHTML Trim(GText)
End Select
frm.Page.SelStart = 1
frm.Page.Locked = True
End Function
Public Function LineCount(txtBox As TextBox) As Long
LineCount = SendMessage(txtBox.hWnd, EM_GETLINECOUNT, 0&, 0&)
End Function
Public Function BrowserBack()
If Back = False Then Exit Function
modTags.LoadURL modTags.BackURL
End Function

