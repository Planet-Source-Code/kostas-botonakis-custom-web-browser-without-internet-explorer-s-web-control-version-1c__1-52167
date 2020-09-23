VERSION 5.00
Begin VB.Form frmMeta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Meta Tag"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Close"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      Begin VB.Label Lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMeta.frx":0000
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   4335
      End
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label content 
      BackStyle       =   0  'Transparent
      Caption         =   "content"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Label httpequiv 
      BackStyle       =   0  'Transparent
      Caption         =   "http-equiv"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   4455
   End
End
Attribute VB_Name = "frmMeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
