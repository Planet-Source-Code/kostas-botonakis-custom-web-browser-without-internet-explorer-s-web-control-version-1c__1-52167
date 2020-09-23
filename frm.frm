VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open Source Web Browser by DarkX"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pic 
      Height          =   495
      Left            =   10200
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAbout 
      Height          =   285
      Left            =   10560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Text            =   "frm.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox PicAddress 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   340
      Left            =   2160
      ScaleHeight     =   345
      ScaleWidth      =   7935
      TabIndex        =   3
      Top             =   7450
      Width           =   7935
      Begin VB.TextBox txtAddress 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   840
         TabIndex        =   5
         ToolTipText     =   "Enter your URL"
         Top             =   40
         Width           =   4695
      End
      Begin VB.Label lblShowAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show URL Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblHideAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hide URL Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   6360
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Top             =   120
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   270
         Left            =   5640
         MouseIcon       =   "frm.frx":0258
         MousePointer    =   99  'Custom
         Picture         =   "frm.frx":0562
         ToolTipText     =   "Navigate"
         Top             =   70
         Width           =   420
      End
      Begin VB.Label lblAdress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog CDSave 
      Left            =   2280
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CDOpen 
      Left            =   2760
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   7485
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   3598
            MinWidth        =   1940
            Text            =   "Open Source Web Browser"
            TextSave        =   "Open Source Web Browser"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicPage 
      Height          =   3735
      Left            =   0
      ScaleHeight     =   3675
      ScaleWidth      =   4185
      TabIndex        =   0
      Top             =   0
      Width           =   4245
      Begin RichTextLib.RichTextBox Page 
         Height          =   3375
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   5953
         _Version        =   393217
         BorderStyle     =   0
         HideSelection   =   0   'False
         ScrollBars      =   3
         MousePointer    =   1
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frm.frx":06E7
         MouseIcon       =   "frm.frx":0765
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   120
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin Web_Control.WebControl WebControl1 
      Left            =   4440
      Top             =   240
      _ExtentX        =   3413
      _ExtentY        =   1931
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpenURL 
         Caption         =   "&Open URL"
      End
      Begin VB.Menu mnuFileSavePage 
         Caption         =   "&Save this page"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsBrowser 
         Caption         =   "&Browser Options"
      End
      Begin VB.Menu mnuOptionsInternet 
         Caption         =   "&Internet Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpBrowser 
         Caption         =   "&Browser Help"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentTag As String
Dim ColorExtract As String
Dim FaceArray() As String
Dim Con As String
Dim strTitle As String



Private Sub Form_Load()
EnableURLs hWnd, frm.Page.hWnd
'modTags.LoadURL "about:browser"
modTags.LoadURL "Test.htm"
'Change Window 's Title before loading any page
frm.Caption = "Open Source Web Browser"
'Resize main picture box
Dim G As Integer
frm.PicPage.Width = frm.Width - 90
G = frm.Height - (SB.Height + PicAddress.Height + 500)
frm.PicPage.Height = frm.Height - frm.PicAddress.Height - 50
'Parse the html code
WebControl1.ParseHTML Trim(GText)
'Resize Rich Text Box
frm.Page.Height = frm.PicPage.Height - 150
frm.Page.Width = frm.PicPage.Width - 150
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SB.Panels(1).Text = "Open Source Web Browser"
lblHideAddress.FontUnderline = False
lblShowAddress.FontUnderline = False
End Sub
Private Sub Form_Resize()
Dim G1 As Integer
G1 = frm.Height - (SB.Height + PicAddress.Height + 500)
frm.PicPage.Height = G1
'Resize Rich Text Box
frm.Page.Height = frm.PicPage.Height - 50
frm.Page.Width = frm.PicPage.Width - 50
End Sub
Private Sub Form_Unload(Cancel As Integer)
DisableURLs hWnd, frm.Page.hWnd
Unload frmMeta
Unload frm
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SB.Panels(1).Text = "Open Source Web Browser"
lblHideAddress.FontUnderline = False
End Sub
Private Sub mnuBack_Click()
modTags.BrowserBack
End Sub
Private Sub lblHideAddress_Click()
Image1.Visible = False
txtAddress.Visible = False
frm.lblAdress.Visible = False
lblHideAddress.Visible = False
lblShowAddress.Visible = True
End Sub
Private Sub lblHideAddress_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHideAddress.FontUnderline = True
End Sub
Private Sub lblShowAddress_Click()
lblHideAddress.Visible = True
Image1.Visible = True
txtAddress.Visible = True
lblAdress.Visible = True
lblShowAddress.Visible = False
End Sub
Private Sub lblShowAddress_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblShowAddress.FontUnderline = True
End Sub
Private Sub mnuFileExit_Click()
Unload frmMeta
Unload frm
End Sub
Private Sub mnuFileOpenURL_Click()
Load frmOpen
frmOpen.Show vbModal, frm
End Sub
Private Sub mnuHelpAbout_Click()
Load frmAbout
frmAbout.Show vbModal, frm
End Sub
Private Sub Page_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SB.Panels(1).Text = "Open Source Web Browser"
End Sub
Private Sub PicAddress_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SB.Panels(1).Text = "Open Source Web Browser"
lblHideAddress.FontUnderline = False
lblShowAddress.FontUnderline = False
End Sub
Private Sub PicPage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SB.Panels(1).Text = "Open Source Web Browser"
End Sub
Private Sub SB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SB.Panels(1).Text = "By DarkX"
lblShowAddress.FontUnderline = False
lblHideAddress.FontUnderline = False
End Sub
Private Sub txtAddress_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SB.Panels(1).Text = "Open Source Web Browser"
lblHideAddress.FontUnderline = False

End Sub

Private Sub WebControl1_HTMLProperty(Property As String, PropertyValue As String)
Select Case CurrentTag
    'Load meta tag(s)
    Case "meta"
        If LCase(Property) = "http-equiv" Then
            frmMeta.httpequiv.Caption = "http-equiv = " & PropertyValue
        End If
        If LCase(Property) = "content" Then
            frmMeta.content.Caption = "Content = " & PropertyValue
        End If
        If LCase(Property) = "name" Then
            frmMeta.lblName.Caption = "Name = " & PropertyValue
        End If
    'Load body tag
    Case "body"
        If LCase(Property) = "bgcolor" Then
            ColorExtract = Right(PropertyValue, 6)
            frm.Page.BackColor = RGB(CByte("&H" & Left(ColorExtract, 2)), CByte("&H" & Mid(ColorExtract, 3, 2)), CByte("&H" & Right(ColorExtract, 2)))
        End If
    'Load any images in the page
    Case "img"
        If LCase(Property) = "src" Then
            modTags.PicLoad PropertyValue
        End If
    'Paragragh formating --> <p></p> tags
    Case "p"
        If Property = "align" Then
            Select Case LCase(PropertyValue)
                Case "left"
                    frm.Page.SelAlignment = 0
                Case "right"
                    frm.Page.SelAlignment = 1
                Case "center"
                    frm.Page.SelAlignment = 2
                Case Else
                    frm.Page.SelAlignment = 0
            End Select
        End If
    'Load Text Alignment
    Case "div"
        If Property = "align" Then
            Select Case LCase(PropertyValue)
                Case "left"
                    frm.Page.SelAlignment = 0
                Case "right"
                    frm.Page.SelAlignment = 1
                Case "center"
                    frm.Page.SelAlignment = 2
                Case Else
                    frm.Page.SelAlignment = 0
            End Select
        End If
    'Load font tag(s) , text formating
    Case "font"
        Select Case LCase(Property)
            Case "face"
                FaceArray = Split(PropertyValue, ",")
                frm.Page.SelFontName = FaceArray(0)
            Case "size"
                Select Case Abs(PropertyValue)
                    Case 1
                        frm.Page.SelFontSize = 8
                    Case 2
                        frm.Page.SelFontSize = 10
                    Case 3
                        frm.Page.SelFontSize = 12
                    Case 4
                        frm.Page.SelFontSize = 14
                    Case 5
                        frm.Page.SelFontSize = 18
                    Case 6
                        frm.Page.SelFontSize = 24
                    Case 7
                        frm.Page.SelFontSize = 36
                End Select
            Case "color"
                ColorExtract = Right(PropertyValue, 6)
                frm.Page.SelColor = RGB(CByte("&H" & Left(ColorExtract, 2)), CByte("&H" & Mid(ColorExtract, 3, 2)), CByte("&H" & Right(ColorExtract, 2)))
        End Select
    End Select
End Sub
Private Sub WebControl1_HTMLTagBegin(Tag As String)
Select Case LCase(Tag)
    Case "title"

    'Load simple text tag
    Case "text"
        frm.Page.SelText = ""
    'Load Custom <script> </script> tag
    Case "script"
        
    'Load bulleting tag
    Case "li"
        frm.Page.SelText = vbCrLf & "  "
        frm.Page.SelBullet = True
    'Load new lines tag
    Case "br"
        frm.Page.SelText = vbCrLf
    'Load Line with shade tag
    Case "hr"
        frm.Page.SelRTF = "{\par }{\insrsid6488218 {\pict{\*\picprop\shplid1025{\sp{\sn shapeType}{\sv 1}}{\sp{\sn fFlipH}{\sv 0}}{\sp{\sn fFlipV}{\sv 0}}{\sp{\sn fillColor}{\sv 8421504}}{\sp{\sn fFilled}{\sv 1}}" & _
         "{\sp{\sn fLine}{\sv 0}}{\sp{\sn alignHR}{\sv 1}}{\sp{\sn dxHeightHR}{\sv 30}}{\sp{\sn fStandardHR}{\sv 1}}{\sp{\sn fHorizRule}{\sv 1}}{\sp{\sn fLayoutInCell}{\sv 1}}}\picscalex831\picscaley6\piccropl0\piccropr0\piccropt0\piccropb0" & _
         "\picw1764\pich882\picwgoal1000\pichgoal500\wmetafile8\bliptag604941812\blipupi96{\*\blipuid 240eadf4c162681dd8f6fbbcaeadc235}0100090000038900000004001c00000000000400000003010800050000000b0200000000050000000c020f00fa02040000002e01180004000000020101000500" & _
         "00000902000000021c000000fb02f0ff0000000000009001000000000440001254696d6573204e657720526f6d616e0000000000000000000000000000000000" & _
         "040000002d0100000d000000320a0d00ffff01000400fffffffff8020d0020dd0700030000001e0007000000fc020000808080000000040000002d0101000c00" & _
         "000040092100f0000000000000000e00f8020000000008000000fa0200000000000000000000040000002d01020007000000fc020000ffffff000000040000002d010300040000002701ffff030000000000}}{\insrsid2369514" & _
         "\par }}"
    'Load Line with noshade tag
    Case "hr noshade"
        frm.Page.SelRTF = "\par }{\insrsid5275356 {\pict{\*\picprop\shplid1025{\sp{\sn shapeType}{\sv 1}}{\sp{\sn fFlipH}{\sv 0}}{\sp{\sn fFlipV}{\sv 0}}{\sp{\sn fillColor}{\sv 8421504}}{\sp{\sn fFilled}{\sv 1}}" & _
        "{\sp{\sn fLine}{\sv 0}}{\sp{\sn alignHR}{\sv 1}}{\sp{\sn dxHeightHR}{\sv 30}}{\sp{\sn fStandardHR}{\sv 1}}{\sp{\sn fNoshadeHR}{\sv 1}}{\sp{\sn fHorizRule}{\sv 1}}{\sp{\sn fLayoutInCell}{\sv 1}}}" & _
        "\picscalex6553\picscaley0\piccropl0\piccropr0\piccropt0\piccropb0\picw1764\pich882\picwgoal1000\pichgoal500\wmetafile8\bliptag1287174130\blipupi96{\*\blipuid 4cb8b7f2426ceedac924f56d61c83902}" & _
        "0100090000038900000004001c00000000000400000003010800050000000b0200000000050000000c020f00fa02040000002e01180004000000020101000500" & _
        "00000902000000021c000000fb02f0ff0000000000009001000000000440001254696d6573204e657720526f6d616e0000000000000000000000000000000000" & _
        "040000002d0100000d000000320a0d00ffff01000400fffffffff8020d0020000700030000001e0007000000fc020000808080000000040000002d0101000c00" & _
        "000040092100f0000000000000000e00f8020000000008000000fa0200000000000000000000040000002d01020007000000fc020000ffffff000000040000002d010300040000002701ffff030000000000}}{\insrsid1969843" & _
        "\par }}"
    'Load BOLD text tag
    Case "b"
        frm.Page.SelBold = True
    'Load ITALIC text tag
    Case "i"
        frm.Page.SelItalic = True
    'Load UNDERLINED text tag
    Case "u"
        frm.Page.SelUnderline = True
    Case "h1"
        frm.Page.SelBold = True
    Case "h2"
        frm.Page.SelBold = True
    Case "h3"
        frm.Page.SelBold = True
    Case "h4"
        frm.Page.SelBold = True
    Case "h5"
        frm.Page.SelBold = True
    Case "h6"
        frm.Page.SelBold = True
    'Load STRIKETHROUGH text tag
    Case "strike"
        frm.Page.SelStrikeThru = True
    'Load BOLD text tag
    Case "strong"
        frm.Page.SelBold = True
    'Case Else
     '   frm.Page.SelText = Tag

    End Select
CurrentTag = LCase(Tag)
End Sub
Private Sub WebControl1_HTMLTagClose(Tag As String)
Select Case LCase(Tag)
    Case "title"
    'End of simple text tag
    Case "text"
    frm.Page.SelText = ""
    'End of Bulleting tag
    Case "li"
        frm.Page.SelText = vbCrLf
        frm.Page.SelBullet = False
    'End of BOLD text tag
    Case "b"
        frm.Page.SelBold = False
    'End of ITALIC text tag
    Case "i"
        frm.Page.SelItalic = False
    'End of UNDERLINED text tag
    Case "u"
        frm.Page.SelUnderline = False
    'End of TEXT alignment
    Case "div"
        frm.Page.SelText = vbCrLf
    Case "h1"
        frm.Page.SelText = vbCrLf
        frm.Page.SelBold = False
    Case "h2"
        frm.Page.SelText = vbCrLf
        frm.Page.SelBold = False
    Case "h3"
        frm.Page.SelText = vbCrLf
        frm.Page.SelBold = False
    Case "h4"
        frm.Page.SelText = vbCrLf
        frm.Page.SelBold = False
    Case "h5"
        frm.Page.SelText = vbCrLf
        frm.Page.SelBold = False
    Case "h6"
        frm.Page.SelText = vbCrLf
        frm.Page.SelBold = False
    'End of STRIKETHROUGH text tag
    Case "strike"
        frm.Page.SelStrikeThru = False
    'End of BOLD text tag
    Case "strong"
        frm.Page.SelBold = False
End Select
End Sub

Private Sub WebControl1_HTMLText(Text As String)
'Replace html characters with known characters
Text = Replace(Text, "&nbsp;", " ")
Text = Replace(Text, "&amp;", "&")
Text = Replace(Text, "&reg;", "®")
Select Case LCase(CurrentTag)
    'Load Page's title
    Case "title"
            frm.Caption = "Open Source Web Browser - " & Trim(Text)
    Case "td"
        Con = Trim(Text)
        frm.Page.TextRTF = frm.Page.TextRTF & "{\par\trowd\trgaph108\trleft-108\trbrdrt\brdrs\brdrw20 \trbrdrl\brdrs\brdrw20 \trbrdrb\brdrs\brdrw20 \trbrdrr\brdrs\brdrw20 \clbrdrt\brdrw15\brdrs\clbrdrl\brdrw15\brdrs\clbrdrb\brdrw15\brdrs\clbrdrr\brdrw15\brdrs \cellx8748\pard\intbl " & Con & "\cell\row\pard\nowidctlpar\par}"
    'Load Page's text
    Case Else
        frm.Page.SelText = Trim(Text)
        'frm.Page.SelText & Trim(Text)
End Select
End Sub
