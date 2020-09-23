VERSION 5.00
Begin VB.UserControl WebControl 
   ClientHeight    =   1170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1995
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1170
   ScaleWidth      =   1995
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "[Control]"
      Height          =   195
      Left            =   1200
      TabIndex        =   2
      Top             =   480
      Width           =   585
   End
   Begin VB.Shape Shape2 
      Height          =   1060
      Left            =   0
      Top             =   0
      Width           =   1900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Open Source"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   930
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      FillColor       =   &H00808080&
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DarkX Web Browser"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1425
   End
End
Attribute VB_Name = "WebControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event HTMLTagBegin(Tag As String)
Attribute HTMLTagBegin.VB_Description = "Fires when a tag begins"
Public Event HTMLProperty(Property As String, PropertyValue As String)
Attribute HTMLProperty.VB_Description = "Returns a property's name and value"
Public Event HTMLTagEnd(Tag As String)
Attribute HTMLTagEnd.VB_Description = "Fires when a tag ends, i.e. tag begin, properties and values, tag end,html text, tag close"
Public Event HTMLTagClose(Tag As String)
Attribute HTMLTagClose.VB_Description = "Fires when a tag closes i.e. </"
Public Event HTMLText(Text As String)
Attribute HTMLText.VB_Description = "Fires when text in between HTML tags is found"
Public Sub ParseHTML(HTML As String)
Attribute ParseHTML.VB_Description = "Parses HTML and fires HTML events"
'Go through the HTML, character by character
'checking first for <, then for spaces, then
'quotation marks, and finally /. As we find
'them we fire events and continue parsing.
'-------------------------------------------------------
Dim IsValue, IsProperty, IsTag, RaisedTagBegin As Boolean
Dim i As Long
Dim CurrentChar As String
Dim CurrentProperty As String
Dim CurrentPropertyValue As String
Dim CurrentTag As String
Dim CurrentText As String
'Remove tabs and returns, they have no place in HTML
HTML = Replace(HTML, vbCrLf, "")
HTML = Replace(HTML, vbTab, "")
'Start our searching
For i = 1 To Len(HTML)
    CurrentChar = Mid(HTML, i, 1)
    If IsTag = True Then
        If IsProperty = True Then
            If IsValue = True Then
                If CurrentChar = Chr(34) Then
                    IsValue = False
                    IsProperty = False
                    CurrentPropertyValue = Trim(CurrentPropertyValue)
                    CurrentProperty = Trim(CurrentProperty)
                    RaiseEvent HTMLProperty(Left(CurrentProperty, Len(CurrentProperty) - 1), CurrentPropertyValue)
                    CurrentPropertyValue = ""
                    CurrentProperty = ""
                Else
                    CurrentPropertyValue = CurrentPropertyValue & CurrentChar
                End If
            ElseIf CurrentChar = Chr(34) Then
                IsValue = True
            Else
                CurrentProperty = CurrentProperty & CurrentChar
            End If
        Else
            If CurrentChar = " " Then
                IsProperty = True
                CurrentTag = Trim(CurrentTag)
                CurrentTag = CurrentTag
                If RaisedTagBegin = False Then
                    RaiseEvent HTMLTagBegin(CurrentTag)
                    RaisedTagBegin = True
                End If
            ElseIf CurrentChar = ">" Then
                IsTag = False
                If Left(CurrentTag, 1) = "/" Then
                    RaiseEvent HTMLTagClose(Right(CurrentTag, Len(CurrentTag) - 1))
                ElseIf RaisedTagBegin = False Then
                    RaiseEvent HTMLTagBegin(CurrentTag)
                    RaiseEvent HTMLTagEnd(CurrentTag)
                    RaisedTagBegin = True
                Else
                    RaiseEvent HTMLTagEnd(CurrentTag)
                End If
                CurrentTag = ""
                
            Else
                CurrentTag = CurrentTag & CurrentChar
            End If
        End If
    Else
        If CurrentChar = "<" Then
            IsTag = True
            RaisedTagBegin = False
            If Trim(CurrentText) <> "" Then
               RaiseEvent HTMLText(Trim(CurrentText))
'                HTMLText Trim(CurrentText)
                CurrentText = ""
            End If
        Else
            CurrentText = CurrentText & CurrentChar
        End If
    End If
Next i
End Sub

Private Sub UserControl_Resize()
'Resize the control
UserControl.Height = Shape1.Height
UserControl.Width = Shape1.Width
End Sub
Private Sub UserControl_Show()
'Resize the control
UserControl.Height = Shape1.Height
UserControl.Width = Shape1.Width
End Sub
