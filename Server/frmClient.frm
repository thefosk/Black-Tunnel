VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmClient 
   BackColor       =   &H00CEDBDE&
   Caption         =   "TFTPClient"
   ClientHeight    =   6420
   ClientLeft      =   300
   ClientTop       =   585
   ClientWidth     =   4695
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   4695
   Visible         =   0   'False
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   5640
      Width           =   3615
   End
   Begin VB.TextBox txtOK 
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   6000
      Width           =   3615
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3960
      Top             =   5880
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   180
      TabIndex        =   9
      Text            =   "C:\MyFile.exe"
      Top             =   4260
      Width           =   4335
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00EFF7F7&
      Height          =   1845
      Left            =   2400
      TabIndex        =   7
      Top             =   1140
      Width           =   2115
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00EFF7F7&
      Height          =   1890
      Left            =   180
      TabIndex        =   6
      Top             =   1140
      Width           =   2115
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00EFF7F7&
      Height          =   315
      Left            =   180
      TabIndex        =   5
      Top             =   720
      Width           =   4335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00DEE7EF&
      Caption         =   "Send File To Server"
      Height          =   795
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   2115
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00DEE7EF&
      Caption         =   "Connect To Server"
      Height          =   795
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   2115
   End
   Begin MSWinsockLib.Winsock WskClient 
      Left            =   2040
      Top             =   900
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Remote Destination"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   180
      TabIndex        =   10
      Top             =   4020
      Width           =   2475
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No Connected"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   300
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      Height          =   555
      Left            =   180
      TabIndex        =   4
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lbFileSend 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "File Sended:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   675
      Left            =   180
      TabIndex        =   1
      Top             =   5100
      Width           =   4335
   End
   Begin VB.Label lbByteSend 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Byte Sended:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   4680
      Width           =   4335
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SrcPath As String
Dim DstPath As String
Dim IsReceived As Boolean



Private Sub SendFile()
    Dim BufFile As String
    Dim LnFile As Long
    Dim nLoop As Long
    Dim nRemain As Long
    Dim Cn As Long
    
    On Error GoTo GLocal:
    LnFile = FileLen(SrcPath)
    If LnFile > 8192 Then
        nLoop = Fix(LnFile / 8192)
        
        nRemain = LnFile Mod 8192
    Else
        nLoop = 0
        nRemain = LnFile
    End If
    
    If LnFile = 0 Then
        ErrDU = ErrDU & "Ivalid Source File" & vbCrLf
        Exit Sub
    End If
    
    Open SrcPath For Binary As #1
    If nLoop > 0 Then
        For Cn = 1 To nLoop
            BufFile = String(8192, " ")
            Get #1, , BufFile
            WskClient.SendData BufFile
            IsReceived = False
            lbByteSend.Caption = "Bytes Sent: " & Cn * 81092 & " Of " & LnFile
            lbByteSend.Refresh
            While IsReceived = False
                DoEvents
            Wend
        Next
        If nRemain > 0 Then
            BufFile = String(nRemain, " ")
            Get #1, , BufFile
            WskClient.SendData BufFile
            IsReceived = False
            lbByteSend.Caption = "Bytes Sent: " & LnFile & " Of " & LnFile
            lbByteSend.Refresh
            While IsReceived = False
                DoEvents
            Wend
        End If
    Else
        BufFile = String(nRemain, " ")
        Get #1, , BufFile
        WskClient.SendData BufFile
        IsReceived = False
        While IsReceived = False
            DoEvents
        Wend
    End If
    WskClient.SendData "Msg_Eof_"    'end of file tag
    Close #1
    
    
    
    Exit Sub
GLocal:
    ErrDU = ErrDU & Err.Description & vbCrLf
    
End Sub
Private Sub ConnectToServer()
        'connecting to localhost
        'if you want to connect to another
        'client replace IP address with
        'remote computer name
        On Error Resume Next

        WskClient.Connect Connection, 34045
        If Err <> 0 Then
            WskClient.Close
        End If
End Sub
Private Sub Command1_Click()
    ConnectToServer
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    If SrcPath = "" Then
        ErrDU = ErrDU & "Select File to transfer!" & vbCrLf
        Exit Sub
    End If
    lbFileSend.Caption = SrcPath
    'send to server the remote path
    WskClient.SendData "Msg_Dst_" & DstPath
    Exit Sub
    

    
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Left(Drive1.Drive, 2) & "\"
End Sub

Private Sub File1_Click()
    SrcPath = File1.Path
    If Right(SrcPath, 1) <> "\" Then
        SrcPath = SrcPath & "\"
    End If
    SrcPath = SrcPath & File1.FileName
    'default destination path
    'if client and server are running on
    'the same machine
    Text1.Text = "C:\" & File1.FileName
    'to prevent overwrite source destination file
    'when client and server are running on
    'the same machine
    If Text1.Text = SrcPath Then
        Text1.Text = "C:\TFTPFile." & Right(File1.FileName, 3)
    End If
    lbFileSend.Caption = SrcPath
End Sub
'This code was edited by BOmBeR from an other code found on www.freevbcode.com

Private Sub Form_Load()
    WskClient.Protocol = sckTCPProtocol
    
End Sub

Private Sub Text1_Change()
    DstPath = Text1.Text
End Sub

Private Sub Timer1_Timer()
If Left(txtOK, 1) = "1" Then
    
    SrcPath = txtPath
    Text1 = Mid(txtOK, 3)
    DstPath = Text1
    Call Command1_Click
    
End If
End Sub



Private Sub WskClient_Close()
    lblInfo.Caption = "Not Connected..."
    WskClient.Close
End Sub

Private Sub WskClient_Connect()
    Call Command3_Click
    txtOK = ""
End Sub

Private Sub WskClient_DataArrival(ByVal bytesTotal As Long)
    Dim RecBuffer As String
    
    WskClient.GetData RecBuffer
    
    Select Case Left(RecBuffer, 7)
    Case "Msg_Rec"  'Block Received
        IsReceived = True
    Case "Msg_OkS"  'Ok you can begin to send file
        SendFile
    Case "Msg_Res"  'resent bad block
        'implement this case
    Case "Msg_Err"  'error
        'implement this case
    End Select
End Sub

Private Sub WskClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Number <> 0 Then
        lblInfo.Caption = "Not Connected"
    End If
End Sub

Private Sub WskClient_SendComplete()
'    WskClient.SendData "Msg_Eof_"    'end of file tag

End Sub

Private Sub WskClient_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
'    DoEvents
'    lbByteSend.Caption = "Bytes Sent: " & bytesSent & " Of " & bytesRemaining
'    lbByteSend.Refresh
    
End Sub
