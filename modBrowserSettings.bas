Attribute VB_Name = "modBrowserSettings"
Const App = "Open Source Web Browser"
Const Sect = "Browser"
Function BrowserSettings(What As String)
'Save/Load settings to/from Registry
Select Case What
'Save settings in registry
Case "Save"
'Save Browser's Title
SaveSetting App, Sect, "Browser Title", frm.Caption
'Save Browser's Default Url
SaveSetting App, Sect, "Browser Url", ""
'-------------------------------
'-------------------------------
'-------------------------------
'Load Settings from registry
Case "Load"
Dim frmTitle As String
Dim frmUrl As String
frmTitle = GetSetting(App, Sect, "Browser Title")
frmUrl = GetSetting(App, Sect, "Browser Url")
frm.Caption = frmTitle
'... = frmUrl
End Select
End Function
Function SaveHistory(strURLHistory As String)
Dim Histories As Integer
Histories = GetSetting(App, Sect, "History Urls")
If Histories > 0 Then
Histories = Histories + 1
SaveSetting App, Sect, "HistoryURL_" & Histories, strURLHistory
Else
Histories = 1
SaveSetting App, Sect, "HistoryURL_" & Histories, strURLHistory
End If
End Function
