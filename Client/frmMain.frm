VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Black Tunnel 1.0 Client"
   ClientHeight    =   8415
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDownload 
      Height          =   285
      Left            =   4320
      TabIndex        =   16
      Text            =   "34045"
      Top             =   7920
      Width           =   1575
   End
   Begin VB.TextBox txtUpload 
      Height          =   285
      Left            =   4320
      TabIndex        =   15
      Text            =   "34046"
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "Upload Server"
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download Client"
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   6960
      Width           =   1815
   End
   Begin RichTextLib.RichTextBox txtReceived 
      Height          =   4455
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7858
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "SEND"
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   6615
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Text"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   6960
      Width           =   855
   End
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   7080
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtPORT 
      Height          =   285
      Left            =   4440
      TabIndex        =   4
      Text            =   "12098"
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Download Port:"
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Upload Port:"
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "PORT:"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "IP:"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "BLACK TUNNEL v1.0 CLIENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu MnuDelFile 
         Caption         =   "Delete File"
      End
      Begin VB.Menu MnuCopyFile 
         Caption         =   "Copy File"
      End
      Begin VB.Menu MnuNameFile 
         Caption         =   "Rename File"
      End
      Begin VB.Menu MnuMoveFile 
         Caption         =   "Move File"
      End
      Begin VB.Menu MnuLoadFile 
         Caption         =   "Execute file"
      End
      Begin VB.Menu MnuExistFile 
         Caption         =   "Check if a file exists"
      End
      Begin VB.Menu MnuFileAttributes 
         Caption         =   "Set File attributes"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
txtReceived.Text = ""
End Sub

Private Sub cmdConnect_Click()
On Error GoTo Err
sckClient.Connect txtIP, txtPORT
cmdDisconnect.Enabled = True
cmdConnect.Enabled = False
Exit Sub
Err:
MsgBox "Errore! " & Err.Description, vbCritical, "Err No: " & Err.Number
End Sub

Private Sub cmdDisconnect_Click()
On Error GoTo Err
sckClient.Close
Unload frmClient
Unload frmSvr
cmdDisconnect.Enabled = False
cmdConnect.Enabled = True
Exit Sub
Err:
MsgBox "Errore! " & Err.Description, vbCritical, "Err No: " & Err.Number
End Sub



Private Sub cmdDownload_Click()
frmSvr.Show
End Sub

Private Sub cmdSend_Click()
Dim Data As String
txtReceived.Text = txtReceived.Text & txtSend.Text & vbCrLf
Data = txtSend
sckClient.SendData Data & vbCrLf
txtSend.Text = ""
End Sub

Private Sub cmdUpload_Click()
frmClient.Show
End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
On Error GoTo Err
Dim Data As String
sckClient.GetData Data, vbString
txtReceived.Text = txtReceived.Text & Data
Exit Sub
Err:
MsgBox "Errore! " & Err.Description, vbCritical, "Err No: " & Err.Number
End Sub

