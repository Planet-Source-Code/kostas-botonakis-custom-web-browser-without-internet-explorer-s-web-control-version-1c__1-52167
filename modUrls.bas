Attribute VB_Name = "modUrls"
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
     ByVal hWnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     lParam As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
     ByVal hWnd As Long, _
     ByVal lpOperation As String, _
     ByVal lpFile As String, _
     ByVal lpParameters As String, _
     ByVal lpDirectory As String, _
     ByVal nShowCmd As Long) As Long
Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type
Private Type NMHDR_RICHEDIT
    hwndFrom As Long
    wPad1 As Integer
    idfrom As Integer
    code As Integer
    wPad2 As Integer
End Type
Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type
Private Type TEXTRANGE
    chrg As CHARRANGE
    lpstrText As Long
End Type
Private Type ENLINK
    NMHDR As NMHDR_RICHEDIT
    msg As Integer
    wPad1 As Integer
    wParam As Integer
    wPad2 As Integer
    lParam As Integer
    chrg As CHARRANGE
End Type
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const GWL_WNDPROC = (-4)
Private Const WM_USER = &H400
Private Const WM_NOTIFY = &H4E
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const EM_SETEVENTMASK = (WM_USER + 69)
Private Const EM_GETTEXTRANGE = (WM_USER + 75)
Private Const EM_AUTOURLDETECT = (WM_USER + 91)
Private Const EM_EXSETSEL = (WM_USER + 55)
Private Const ENM_LINK = &H4000000
Private Const ENM_NONE = &H0
Private Const EN_LINK = &H70B&
Private mlWndProc As Long
Public Sub EnableURLs(ByVal hWndParent As Long, ByVal hWndRTB As Long)
  ' Turn on URL Detection
  Call SendMessage(hWndRTB, EM_AUTOURLDETECT, 1, ByVal 0)
  ' Tell it to send the EN_LINK notification message
  Call SendMessage(hWndRTB, EM_SETEVENTMASK, 0&, ByVal ENM_LINK)
  ' Subclass the RTB Parent
  mlWndProc = SetWindowLong(hWndParent, GWL_WNDPROC, AddressOf SubClassedWindowProc)
End Sub
Public Sub DisableURLs(ByVal hWndParent As Long, ByVal hWndRTB As Long)
  ' Turn off subclassing
  Call SetWindowLong(hWndParent, GWL_WNDPROC, mlWndProc)
  ' Turn off URL Detection
  Call SendMessage(hWndRTB, EM_AUTOURLDETECT, 0&, ByVal 0&)
  ' Turn off EN_LINK notifications
  Call SendMessage(hWndRTB, EM_SETEVENTMASK, 0&, ByVal ENM_NONE)
End Sub
Private Function SubClassedWindowProc(ByVal hWnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim tNMHDR As NMHDR
  Dim tENLINK As ENLINK
  Dim tTEXT As TEXTRANGE
  Dim sBuffer As String
  ' Look for Notification Messages (WM_NOTIFY)
  If nMsg = WM_NOTIFY Then
    Call CopyMemory(tNMHDR, ByVal lParam, Len(tNMHDR))
    ' If it's the Link Notification...
    If tNMHDR.code = EN_LINK Then
      ' Copy the Notification into a Link Notification Structure
      Call CopyMemory(tENLINK, ByVal lParam, Len(tENLINK))
      ' See what event caused this notification,
      ' If it was a Left Mouse Button Down event, process it..
      If tENLINK.msg = WM_LBUTTONDOWN Then
        ' Transfer the character range contianing the URL
        ' to the TEXTRANGE structure
        LSet tTEXT.chrg = tENLINK.chrg
        ' Create a Buffer to hold the URL text
        sBuffer = String(tTEXT.chrg.cpMax - tTEXT.chrg.cpMin, Chr(0))
        ' Assign the buffer to the TEXTRANGE structure
        tTEXT.lpstrText = StrPtr(sBuffer)
        ' Tell the RTB to give us the text in the given range
        Call SendMessage(tNMHDR.hwndFrom, EM_GETTEXTRANGE, 0, tTEXT)
        ' Strip out the null characters
        sBuffer = Replace(StrConv(sBuffer, vbUnicode), Chr(0), "")
        ' Launch the URL
        ShellExecute hWnd, "OPEN", sBuffer, "", "", 1
      End If
    End If
  End If
  SubClassedWindowProc = CallWindowProc(mlWndProc, hWnd, nMsg, wParam, lParam)
End Function
