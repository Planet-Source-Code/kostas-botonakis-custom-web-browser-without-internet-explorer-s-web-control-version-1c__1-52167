VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Open Source Web Browser"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Close"
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "This web control was developed by DarkX(Kostas Botonakis) and Phreak(Simon Whitehead)"
      Height          =   735
      Left            =   2280
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   1800
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2100
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"frmAbout.frx":4DF0
      Height          =   855
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload frmAbout

End Sub

