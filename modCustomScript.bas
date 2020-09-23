Attribute VB_Name = "modCustomScript"
Public Author As String
Public Company As String
Public Comments As String
Public Menu As Boolean
Public OnlineBrowsing As Boolean
Function LoadCustomScript(cScript As String)
'Check the script's characters length
Dim Length As Integer
Length = Len(cScript)
Select Case cScript
'Date script
Case "date"
    frm.Page.SelText = "Date: " & Date
Case "{date}"
    frm.Page.SelText = "Date: " & Date
'Time Script
Case "time"
    frm.Page.SelText = "Time: " & Time
Case "{time}"
    frm.Page.SelText = "Time: " & Time
'Date and Time script
Case "date+time"
    frm.Page.SelText = "Date: " & Date & " - Time: " & Time
Case "date + time"
    frm.Page.SelText = "Date: " & Date & " - Time: " & Time
Case "{date + time}"
    frm.Page.SelText = "Date: " & Date & " - Time: " & Time
Case "{date+time}"
    frm.Page.SelText = "Date: " & Date & " - Time: " & Time
'Time and Date script
Case "time+date"
    frm.Page.SelText = "Time: " & Time & " - Date: " & Date
Case "time + date"
    frm.Page.SelText = "Time: " & Time & " - Date: " & Date
Case "{time+date}"
    frm.Page.SelText = "Time: " & Time & " - Date: " & Date
Case "{time + date}"
    frm.Page.SelText = "Time: " & Time & " - Date: " & Date
'Page's Author Script
Case Left(cScript, 7) = "author="
    Author = Mid(cScript, 8)
Case Left(cScript, 7) = "Author="
    Author = Mid(cScript, 8)
Case Left(cScript, 8) = "author ="
    Author = Mid(cScript, 9)
Case Left(cScript, 8) = "Author ="
    Author = Mid(cScript, 9)
'Page's Company Script
Case Left(cScript, 8) = "company="
    Company = Mid(cScript, 9)
Case Left(cScript, 8) = "Company="
    Company = Mid(cScript, 9)
Case Left(cScript, 9) = "company ="
    Company = Mid(cScript, 10)
Case Left(cScript, 9) = "Company ="
    Company = Mid(cScript, 10)
'Page's Comments Script
Case Left(cScript, 9) = "comments="
    Comments = Mid(cScript, 10)
Case Left(cScript, 9) = "Comments="
    Comments = Mid(cScript, 10)
Case Left(cScript, 10) = "comments ="
    Comments = Mid(cScript, 11)
Case Left(cScript, 10) = "Comments ="
    Comments = Mid(cScript, 11)
Case Left(cScript, 5) = "menu="
    Dim M
    M = Mid(cScript, 6)
    If M = "true" Then
    Menu = True
    ElseIf M = "True" Then
    Menu = True
    Else
    Menu = False
    End If
Case Left(cScript, 16) = "online browsing="
    Dim B
    B = Mid(cScript, 17)
    If B = "true" Then
    OnlineBrowsing = True
    ElseIf B = "True" Then
    OnlineBrowsing = True
    Else
    OnlineBrowsing = False
    End If
End Select
End Function
