Attribute VB_Name = "modParameters"
Public Source As String
Public Destination As String
Public Arg As String
Public Data As String



Sub FindArg(StartCommand As Integer)

Dim Char As String

For i = 1 To Len(Data)

    Char = Mid(Data, i, 1)
    
    If Char = ";" Then    'If It finds the ---> ; <---

    Source = Mid(Data, StartCommand, i - 1 - StartCommand + 1)
    Destination = Mid(Data, i + 1)
    End If
    
Next i

End Sub

Sub FindX()

Dim Char As String

For i = 1 To Len(Data)

    Char = Mid(Data, i, 1)
    
    If Char = "-" Then
    Arg = Mid(Data, i + 1, 1)
    Source = Mid(Data, i + 3)
    End If
    
Next i

End Sub


Sub Reset()
Source = ""
Destination = ""
Arg = ""
End Sub

'Reset Data
Sub RD()
Data = ""
End Sub
