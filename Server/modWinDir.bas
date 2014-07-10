Attribute VB_Name = "modWinDir"

'Per trovare le cartelle di sistema di Windows

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
(ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" _
(ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Fso As New FileSystemObject    'Oggetto FileSystemObject da usare per le operazioni con i files e cartelle



'Trova le cartelle di Windows e System32
Public Function GetWinDir() As String
    Dim Temp As String * 256
    Dim x As Integer
    x = GetWindowsDirectory(Temp, Len(Temp))
    GetWinDir = Left$(Temp, x)
    GetWinDir = LCase(IIf(Right(GetWinDir, 1) = "\", GetWinDir, GetWinDir & "\"))
End Function

Public Function GetWinSysDir() As String
    Dim Temp As String * 256
    Dim x As Integer
    x = GetSystemDirectory(Temp, Len(Temp))
    GetWinSysDir = Left$(Temp, x)
    GetWinSysDir = LCase(IIf(Right(GetWinSysDir, 1) = "\", GetWinSysDir, GetWinSysDir & "\"))
End Function

