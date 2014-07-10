VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSvr 
   BackColor       =   &H00CEDBDE&
   Caption         =   "TFTPServer"
   ClientHeight    =   2850
   ClientLeft      =   6780
   ClientTop       =   825
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   4680
   Begin MSWinsockLib.Winsock WskServer 
      Index           =   0
      Left            =   180
      Top             =   2340
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lbBytesReceived 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bytes Received"
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
      TabIndex        =   2
      Top             =   1920
      Width           =   4275
   End
   Begin VB.Label lbFilereceived 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "File Received"
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
      TabIndex        =   1
      Top             =   1500
      Width           =   4275
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Server Is Listen..."
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
      Height          =   975
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   4275
   End
End
Attribute VB_Name = "frmSvr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DstPath As String
Dim BytesRec As Long
Dim FL As Integer

Private Sub Form_Load()
    WskServer(0).Protocol = sckTCPProtocol
    WskServer(0).LocalPort = frmMain.txtDownload.Text
    WskServer(0).Listen
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    For Cn = 0 To WskServer.Count - 1
        WskServer(Cn).Close
    Next
End Sub

Private Sub WskServer_Close(Index As Integer)
    lblInfo.Caption = "Connection Closed..."

End Sub

Private Sub WskServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Load WskServer(Index + 1)
    WskServer(Index + 1).Accept requestID
    lblInfo.Caption = "Connection Estabilished..."
End Sub

Private Sub WskServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim recBuffer As String
    Dim Ret As Integer
    
    On Error GoTo GLocal
    
    WskServer(Index).GetData recBuffer
    Select Case Left(recBuffer, 8)
        Case "Msg_Eof_"
            Close #FL
            lblInfo.Caption = "File Received..."
        Case "Msg_Dst_"
            DstPath = Right(recBuffer, Len(recBuffer) - 8)
                
            FL = FreeFile
            'overwrite file
            On Error Resume Next
            If Len(Dir(DstPath)) > 0 Then
                Ret = MsgBox("File Already Exist!!!" & vbCrLf & " You Wont Overwrite It??", vbQuestion + vbYesNo, "TFTServer Message")
                If Ret = vbYes Then
                    Kill DstPath
                Else
                    'insert code to notify the error
                    'to client
                    Unload Me
                End If
            End If
            Open DstPath For Binary As #FL
            lbFilereceived.Caption = DstPath
            WskServer(Index).SendData "Msg_OkS"
        
        Case Else
            BytesRec = BytesRec + Len(recBuffer)
            Put #FL, , recBuffer
            lbBytesReceived.Caption = "Bytes received: " & BytesRec
            WskServer(Index).SendData "Msg_Rec"
    End Select
    
    Exit Sub
GLocal:
    MsgBox Err.Description
    Unload Me
    
End Sub

Private Sub WskServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Number <> 0 Then
        MsgBox Number
    End If
End Sub
