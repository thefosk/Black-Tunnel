Attribute VB_Name = "modCompName"

Declare Function GetComputerName Lib "kernel32" Alias _
"GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) _
 As Long
 
Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long



Public Function PCname() As String
    Dim temp As String * 128

    Dim l As Long
    l = GetComputerName(temp, 128)
    PCname = temp
End Function


Public Function SetPCname(NewName As String) As Boolean
On Error GoTo Err

SetComputerName (NewName)
SetPCname = True

Exit Function
Err:
SetPCname = False
End Function
