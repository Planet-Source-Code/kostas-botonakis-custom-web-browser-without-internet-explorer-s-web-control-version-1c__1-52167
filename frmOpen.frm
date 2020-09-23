VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOpen 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Open URL - Open Source Web Browser"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd 
      Left            =   480
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Browse"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Open It"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "URL Address:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   990
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()
cd.DialogTitle = "Open local URL"
cd.Filter = "All Files (*.*)|*.*"
cd.ShowOpen
If cd.FileName <> "" Then
modTags.LoadURL cd.FileName
Unload frmOpen
End If

End Sub

Private Sub cmdCancel_Click()
Unload frmOpen

End Sub

Private Sub cmdOpen_Click()
If txtAddress.Text <> "" Then
modTags.LoadURL txtAddress.Text
Unload frmOpen
Else
MsgBox "Please enter a valid filename.", vbInformation, "Invalid Filename"
End If

End Sub
