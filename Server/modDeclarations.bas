Attribute VB_Name = "modDeclarations"
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long


Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Const SW_HIDE = 0
Const SW_SHOW = 5
Const GW_CHILD = 5



Public Connection As String
Public ErrDU As String
Public SrcPath As String






'Show/Hide desktop Icons

Sub ShowDesktopIcons(ByVal Visible As Boolean)
Dim Desktop_Icons As Long
Desktop_Icons = GetWindow(FindWindow("Progman", "Program Manager"), GW_CHILD)

If Visible Then
    ShowWindow Desktop_Icons, SW_SHOW
Else
    ShowWindow Desktop_Icons, SW_HIDE
End If

End Sub

'Show/Hide StartBar

Sub ShowStartBar(ByVal Visible As Boolean)
Dim Start_Bar As Long
Start_Bar = FindWindow("Shell_TrayWnd", "")

If Visible Then
    ShowWindow Start_Bar, SW_SHOW
Else
    ShowWindow Start_Bar, SW_HIDE
End If

End Sub

