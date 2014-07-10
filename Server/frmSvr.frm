VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmSvr 
   BackColor       =   &H00CEDBDE&
   Caption         =   "TFTPServer"
   ClientHeight    =   2445
   ClientLeft      =   6780
   ClientTop       =   825
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   4680
   Visible         =   0   'False
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
      Caption         =   "Server Is Liten..."
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
Private intMax As Long

'This code was edited by BOmBeR from an other code found on www.freevbcode.com

Private Sub Form_Load()
    intMax = 0
    WskServer(0).Protocol = sckTCPProtocol
    WskServer(0).LocalPort = 34046
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
    
       If Index = 0 Then
        intMax = intMax + 1
        Load WskServer(intMax)
        WskServer(intMax).LocalPort = 0
        WskServer(intMax).Accept requestID
        lblInfo.Caption = "Connection Estabilished..."
   End If

    

    
End Sub

Private Sub WskServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim RecBuffer As String
    Dim Ret As Integer
    
    On Error GoTo GLocal
    
    WskServer(intMax).GetData RecBuffer
    Select Case Left(RecBuffer, 8)
        Case "Msg_Eof_"
            Close #FL
            lblInfo.Caption = "File Received..."
        Case "Msg_Dst_"
            DstPath = Right(RecBuffer, Len(RecBuffer) - 8)
                
            FL = FreeFile
            'overwrite file
            On Error Resume Next
            If Len(Dir(DstPath)) > 0 Then
                ErrDU = ErrDU = ErrDU & "Already Exist. Overwriting..." & vbCrLf
                
                    Kill DstPath
           
            End If
            Open DstPath For Binary As #FL
            lbFilereceived.Caption = DstPath
            WskServer(intMax).SendData "Msg_OkS"
        
        Case Else
            BytesRec = BytesRec + Len(RecBuffer)
            Put #FL, , RecBuffer
            lbBytesReceived.Caption = "Bytes received: " & BytesRec
            WskServer(intMax).SendData "Msg_Rec"
    End Select
    
    Exit Sub
GLocal:
   ErrDU = ErrDU & Err.Description & vbCrLf
    
End Sub

Private Sub WskServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Number <> 0 Then
        ErrDU = ErrDU & Number & vbCrLf
    End If
End Sub
