Attribute VB_Name = "modForm"
Declare Function SetSysColors Lib "user32" _
(ByVal nChanges As Long, lpSysColor As _
Long, lpColorValues As Long) As Long

Public Const COLOR_SCROLLBAR = 0 'The Scrollbar colour
Public Const COLOR_BACKGROUND = 1 'Colour of the background with no wallpaper
Public Const COLOR_ACTIVECAPTION = 2 'Caption of Active Window
Public Const COLOR_INACTIVECAPTION = 3 'Caption of Inactive window
Public Const COLOR_MENU = 4 'Menu
Public Const COLOR_WINDOW = 5 'Windows background
Public Const COLOR_WINDOWFRAME = 6 'Window frame
Public Const COLOR_MENUTEXT = 7 'Window Text
Public Const COLOR_WINDOWTEXT = 8 '3D dark shadow (Win95)
Public Const COLOR_CAPTIONTEXT = 9 'Text in window caption
Public Const COLOR_ACTIVEBORDER = 10 'Border of active window
Public Const COLOR_INACTIVEBORDER = 11 'Border of inactive window
Public Const COLOR_APPWORKSPACE = 12 'Background of MDI desktop
Public Const COLOR_HIGHLIGHT = 13 'Selected item background
Public Const COLOR_HIGHLIGHTTEXT = 14 'Selected menu item
Public Const COLOR_BTNFACE = 15 'Button
Public Const COLOR_BTNSHADOW = 16 '3D shading of button
Public Const COLOR_GRAYTEXT = 17 'Grey text, of zero if dithering is used.
Public Const COLOR_BTNTEXT = 18 'Button text
Public Const COLOR_INACTIVECAPTIONTEXT = 19 'Text of inactive window
Public Const COLOR_BTNHIGHLIGHT = 20 '3D highlight of button
Function WindowsFormActiveCaptionColour()
'Load it with:
t& = SetSysColors(1, COLOR_ACTIVECAPTION, RGB(255, 0, 0))
End Function


