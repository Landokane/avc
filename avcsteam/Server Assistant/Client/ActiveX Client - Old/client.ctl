VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl UserControl1 
   ClientHeight    =   6705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9990
   ScaleHeight     =   6705
   ScaleWidth      =   9990
   Begin VB.Frame Frame4 
      Caption         =   "Connection"
      Height          =   2415
      Left            =   60
      TabIndex        =   22
      Top             =   4140
      Visible         =   0   'False
      Width           =   3555
      Begin VB.CommandButton Command11 
         Caption         =   "Connect"
         Height          =   375
         Left            =   2340
         TabIndex        =   32
         Top             =   1980
         Width           =   1155
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   840
         TabIndex        =   31
         Text            =   "209.166.25.149"
         Top             =   420
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   840
         TabIndex        =   25
         Text            =   "26000"
         Top             =   780
         Width           =   915
      End
      Begin VB.TextBox Text4 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   24
         Text            =   "bat"
         Top             =   1620
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   840
         TabIndex        =   23
         Text            =   "Avatar-X"
         Top             =   1260
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Connect to:"
         Height          =   195
         Left            =   60
         TabIndex        =   30
         Top             =   180
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "IP"
         Height          =   195
         Left            =   60
         TabIndex        =   29
         Top             =   420
         Width           =   150
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Port"
         Height          =   195
         Left            =   60
         TabIndex        =   28
         Top             =   840
         Width           =   285
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Username"
         Height          =   195
         Left            =   60
         TabIndex        =   27
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Left            =   60
         TabIndex        =   26
         Top             =   1680
         Width           =   690
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection"
      Height          =   975
      Left            =   8640
      TabIndex        =   18
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command5 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   33
         Top             =   600
         Width           =   1155
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Connect"
         Height          =   315
         Left            =   60
         TabIndex        =   19
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Log Detail"
      Height          =   2415
      Left            =   7140
      TabIndex        =   9
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton Command3 
         Caption         =   "Select None"
         Height          =   255
         Left            =   60
         TabIndex        =   17
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Select All"
         Height          =   255
         Left            =   60
         TabIndex        =   16
         Top             =   1740
         Width           =   1335
      End
      Begin VB.CheckBox LogDetail 
         Caption         =   "Misc"
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   15
         ToolTipText     =   "Name changes, sentry, etc"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox LogDetail 
         Caption         =   "Goals"
         Height          =   255
         Index           =   4
         Left            =   60
         TabIndex        =   14
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox LogDetail 
         Caption         =   "Admin Speech"
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   13
         Top             =   720
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox LogDetail 
         Caption         =   "Kills"
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox LogDetail 
         Caption         =   "Team Speech"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   11
         Top             =   480
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox LogDetail 
         Caption         =   "Speech"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   10
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   315
      Left            =   5940
      TabIndex        =   8
      Top             =   60
      Width           =   795
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "Rcon Command Box"
      Top             =   60
      Width           =   5835
   End
   Begin VB.Frame Frame3 
      Caption         =   "Server"
      Height          =   1395
      Left            =   8640
      TabIndex        =   3
      Top             =   1020
      Width           =   1275
      Begin VB.CommandButton Command6 
         Caption         =   "File Manager"
         Height          =   315
         Left            =   60
         TabIndex        =   6
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Scripts"
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   600
         Width           =   1155
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Players"
         Height          =   315
         Left            =   60
         TabIndex        =   4
         Top             =   960
         Width           =   1155
      End
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Text            =   "Chat Here!"
      ToolTipText     =   "Chat Box"
      Top             =   3780
      Width           =   7035
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   9000
      Top             =   3720
   End
   Begin VB.CommandButton Command8 
      Caption         =   "C"
      Height          =   315
      Left            =   6780
      TabIndex        =   1
      ToolTipText     =   "Clear"
      Top             =   60
      Width           =   315
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   8520
      Top             =   3720
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3315
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   5847
      _Version        =   393217
      BackColor       =   0
      ScrollBars      =   3
      TextRTF         =   $"client.ctx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock TCP1 
      Left            =   8040
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RT2 
      Height          =   735
      Left            =   9660
      TabIndex        =   21
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      _Version        =   393217
      BackColor       =   0
      TextRTF         =   $"client.ctx":00BA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   555
      Left            =   4320
      TabIndex        =   34
      Top             =   4620
      Width           =   4695
   End
   Begin VB.Label lblUpdate 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   7140
      TabIndex        =   20
      Top             =   2460
      Width           =   2775
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Const m_def_BackColor = 0
'Property Variables:
Dim m_BackColor As Long


Type typButtons
    ButtonName As String
    ButtonText As String
    OptionOn As String
    OptionOff As String
    Type As Integer
End Type

Type typConnectUsers
    Name As String
    IP As String
    Version As String
End Type

Type comds
    Name As String
    Exec As String
    NumParams As Integer
    MustHave As Integer
    ScriptName As String
    NumButtons As Integer
    Buttons() As typButtons
End Type

Type typPlayerPos
    x As Integer
    y As Integer
    Z As Integer
End Type

Type Playersss
    Name As String      'name
    UniqueID As String
    IP As String
    Class As Integer    'currentclass (-1 = civ, 0 = random, 1-9 = scout-engy) [only applies to TFC servers]
    UserID As Integer   'server userid
    Team As Integer     'current team
    Status As Boolean
    RealName As String
    Pos As typPlayerPos
    JoinTime As Date
    EntryName  As String
    NumKickVotes As Integer
    Port As Long
    ShutUp As Integer
    Warn As Integer
    Points As Long
    LastEvent As Date
    
    
End Type

Type typPreset
    Allowed As String
    Flags As Variant
    Name As String
End Type

Type typSpeech 'used when the server talks to the client
    ClientText As String 'what someone has to say
    Answers() As String 'what the server will do in return
    NumAnswers As Integer 'how many possibilities there are
End Type

Type OtherSettings
    FontName As Variant
    FontBold As Variant
    FontStrikethru As Variant
    FontColor As Variant
    FontUnderline As Variant
    FontItalic As Variant
    FontSize As Variant
    CommandWidth As Long
    CommandHeight As Long
End Type

Type typUsers
    Name As String
    Password As String
    Allowed As String
    Flags As Variant
    ICQ As String
    Email As String
End Type

Type typRGB 'color set
    r As Byte
    g As Byte
    b As Byte
End Type

Type typKickBans 'people who are BANNED
    Name As String 'persons name
    Clan As String 'persons clan
    UID As String 'persons uniqueid
    Type As Byte    'type of ban:
                    '1 - kick this person
                    '2 - kick anyone in this persons clan
                    '4 - kick anyone with this UID
                    '8 - immidiately BAN this person by putting his uniqueid in the servers ban list
End Type

Type serv
    ServerPort As String
    GamePath As String
    HLPath As String
    LocalConnectPort As String
    RconListenPort As String
    LocalIP As String
End Type

Type typClanMember
    Name As String
    UIN As String
    LastIP As String
End Type

Type typClan
    'Provided to allow you to restrict the use of certain names, and identify players by uniqueid
    Clan As String
    JoinPass As String
    Flags As Integer
    NumMembers As Integer
    Members(1 To 200) As typClanMember
End Type

Type typRealPlayer
    'this type is used to store information about certain players
    'they are identified by uniqueid, and can be found regardless of
    'their current name.
    RealName As String  'stored name
    UniqueID As String  'uniqueid list
    LastName As String  'last name seen with
    Points As String ' is this user in probation mode?
    TimesSeen As String ' Number of times this player has been seen (once a day)
    LastTime As Date    'last time player was seen
    Flags As Long
End Type

Type typSwearWords
    BadWord As String
    Flags As Long
End Type

Type typGeneral
    NoAutoVotes As Boolean 'disable automatic map votes
    NoKickVotes As Boolean 'disable automatic kick votes
    MaxMsg As Integer           'Maximum number of messages before kick (spam protection)
    MaxKickVotes As Integer     'Maximum number of kick votes
    MaxTime As Integer          'Time for MaxMsg
    MaxKicks As Integer         'Kickvoted this many times...
    BanTime As Integer          'Ban him this many minutes
    VotePercent As Integer      'Yes vote for a kickvote
    LoggingDisabled As Boolean  'Is logging enabled?
    LastMapsDisabled As Boolean 'Is last maps disabled?
    Flags As Long               'Checkable Flags
    MaxSpeech As Integer
    MaxSpeechTime As Integer
    MapVoteStartTime As Integer
    MapVoteStartTimeMode As Integer
    MapVoteMode As String
    AutoAddReal As Integer      'Automatically Add RealPlayers
    AutoAddRealTimes As String  'Turn into real realplayers after seen more than this many times
    AutoAddRealDays As String   'Delete them if they are not seen after this many days
    SameSpamTime As String      'Time allowance for sending the same message over again
    SameSpamNum As String       'Number of samespams allowed
    AutoAdminName As String     'the name of the autoadmin
    MapChangeMode As String     ' CHANGELEVEL or MAP
End Type

Type typWebInfo
    'stores info on the built-in web logging
    Enabled As Boolean          'Are we actually using this feature?
    LogPath As String           'Path where the files go
    LogFlags As Long
        'LogFlags:
        '1 - Speech (global / team / admin)
        '2 - Kills
        '4 - Goals
        '8 - Name changes
        '16 - Class changes
        '32 - Team changes
        '64 - Joins / Leaves
    Colors(1 To 21) As typRGB   'List of colours
End Type

Type typEvent
    Mde As Integer '0 - once, 1 - once every xxx, 2 - once on these days
    Times As Integer 'the number of (every) to run
    Every As Integer '0 - weeks, 1 - days, 2 - hours, 3 - minutes, 4 - seconds
    Days(0 To 6) As Boolean '0 - monday, 6 - sunday, true - selected, false - not selected
    FirstCheck As Date 'the first time to run or check on this event (in other words, the first time the sub has to look at it)
    WhatToDo As Integer '0 - Script, 1 - RCON
    ScriptName As String 'The script to run
    ComPara As String 'Commands or parameters (depending on what was selected)
    Name As String      'what is it CALLED!?
End Type

Type typDirList ' Directory Listing
    Name As String      'File/Dir name
    FullPath As String  'Full path (on server) to this file
    Type As Byte        '0 - file, 1 - dir
    DateTime As Date    'guess what this is?
    Size As String      'size of file
End Type

Type SHFILEOPSTRUCT
   hwnd As Long
   wFunc As Long
   pFrom As String
   pTo As String
   fFlags As Long
   fAnyOperationsAborted As Long
   hNameMappings As Long
   lpszProgressTitle As String
End Type

Type typLogFound
    LogFile As String
    LogLine As String
End Type

Type typMessages     'Kinda like email
    MsgText As String
    MsgFor As String
    MsgFrom As String
    MsgTimeSent As Date
    MsgSubj As String
    MsgID As Integer
    Flags As Integer        '1 - request reciept, 2 - unread
End Type

Type typTeleportExit
    x As Integer
    y As Integer
    Z As Integer
    Name As String
End Type

'Teleporters
Dim Tele() As typTeleportExit
Dim NumTele As Integer

'Messages
Dim Messages() As typMessages
Dim NumMessages As Integer

Dim RichColors(0 To 10) As typRGB

'Log Searching
Dim NumLogFound As Long
Dim LogFound() As typLogFound
Dim LogSearchString As String

Dim SendSize As Long
Dim FindReal As String

Dim Swears() As typSwearWords
Dim NumSwears As Integer

Dim DllEnabled As Boolean
Dim GameMode As Integer

Dim ServVersion As String

Dim LastKnownState As Integer


'Files and File Transfer Vars
Dim DirList0() As typDirList
Dim DirList1() As typDirList
Dim NumDirs(0 To 1) As Long
Dim DirFullPath(0 To 1) As String     'path to the current window
Dim FileBuffer As String
Dim FileMode As Integer               '0 - unset, 1 - edit this file, 2 - download this file
Dim FilePath As String                'Path to download to for mode 2 above
Dim FileLocalPath As String           'Where to save this file LOCALLY
Dim FileSize As Long                  'Size of file
Dim LastRefresh As Integer
Dim FileRecieveMode As Boolean
Dim SendingFile As Boolean            'Are we currently sending a file?
Dim EditMode As Boolean              'Editing this file now?
Dim TheEditFile As String               'File we are editing on server
Dim EditFileTemp As String
Dim FileWriteNum As Integer
Dim BytesTransferred As Long
Dim LastData As Long
Dim ByteCount As Long
Dim FileStop As Boolean


'Scheduler :)
Dim Events() As typEvent
Dim NumEvents As Integer

'time/update
Dim SecondsLeft As Integer
Dim MapName As String
Dim PlayersOn As String

Dim ShowPlayers As Boolean
Dim ShowUsers As Boolean
Dim ShowMap As Boolean

Dim EmailCheckCounter As Integer

'web
Dim Web As typWebInfo

'real players
Dim RealPlayers() As typRealPlayer
Dim NumRealPlayers As Integer

Dim Server As serv
Dim Presets(1 To 200) As typPreset
Dim NumPresets As Integer

Dim EditedButton As Integer

Dim ConnectUsers(1 To 400) As typConnectUsers
Dim NumConnectUsers As Integer

Dim NumKickBans As Integer
Dim KickBans(1 To 200) As typKickBans

Const SB_VERT = 1

Public Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
Private Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Const FO_MOVE As Long = &H1
Const FO_COPY As Long = &H2
Const FO_DELETE As Long = &H3
Const FO_RENAME As Long = &H4

Const FOF_MULTIDESTFILES As Long = &H1
Const FOF_CONFIRMMOUSE As Long = &H2
Const FOF_SILENT As Long = &H4
Const FOF_RENAMEONCOLLISION As Long = &H8
Const FOF_NOCONFIRMATION As Long = &H10
Const FOF_WANTMAPPINGHANDLE As Long = &H20
Const FOF_CREATEPROGRESSDLG As Long = &H0
Const FOF_ALLOWUNDO As Long = &H40
Const FOF_FILESONLY As Long = &H80
Const FOF_SIMPLEPROGRESS As Long = &H100
Const FOF_NOCONFIRMMKDIR As Long = &H200
Const GW_CHILD = 5
Const GW_HWNDFIRST = 0
Const GW_HWNDLAST = 1
Const GW_HWNDNEXT = 2
Const GW_HWNDPREV = 3
Const GW_OWNER = 4
Const SW_SHOW = 5


Const CB_ERR = (-1)
Const WM_USER = &H400
Const CB_FINDSTRING = &H14C

Dim Commands() As comds
Dim NumCommands As Integer

Dim Players(1 To 400) As Playersss
Dim NumPlayers As Integer

Dim Users(1 To 200) As typUsers
Dim NumUsers As Integer

Dim General As typGeneral


'clannies
Dim Clans(1 To 20) As typClan
Dim NumClans As Integer

'speech
Dim Speech() As typSpeech
Dim NumSpeech As Integer

Dim ChosenClan As Integer

Dim Settings As OtherSettings
Dim DataFile As String
Dim DataFile2 As String
Dim DataFile3 As String
Dim RecData As String
Dim LoginName As String
Dim LoginPass As String
Dim UserEditNum As Integer

'Map
Dim MapArray(0 To 64, 0 To 64) As Integer ' Stores Z coord at this location




Private Sub Command11_Click()

AttemptConnect Text5, Text1, Text3, Text4

LoginName = Text3
LoginPass = Text4
TCP1.RemoteHost = Text5
TCP1.RemotePort = Val(Text1)
TCP1.Connect


End Sub

Private Sub Command4_Click()
Frame4.Visible = True

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Command1_Click
    End If
    Dim CB As Long
    Dim FindString As String

    If KeyAscii < 32 Or KeyAscii > 127 Then Exit Sub

    If Combo1.SelLength = 0 Then
        FindString = Combo1.Text & Chr$(KeyAscii)
    Else
        FindString = Left$(Combo1.Text, Combo1.SelStart) & Chr$(KeyAscii)
    End If

    CB = SendMessage(Combo1.hwnd, CB_FINDSTRING, -1, ByVal FindString)

    If CB <> CB_ERR Then
        Combo1.ListIndex = CB
        Combo1.SelStart = Len(FindString)
        Combo1.SelLength = Len(Combo1.Text) - Combo1.SelStart
        KeyAscii = 0
    End If
    
End Sub

Private Sub Command1_Click()
'    hed$ = Chr(255) + Chr(255) + Chr(255) + Chr(255) + "rcon" + Chr(32)
'    UDP1.SendData hed$ + " " + Text2 + " " + Combo1.Text + Chr(255) + Chr(255) + Chr(255) + Chr(255)
        
    SendPacket "RC", Combo1.Text
        
    Dim CB As Long
    
    CB = SendMessage(Combo1.hwnd, CB_FINDSTRING, -1, ByVal Combo1.Text)

    If CB = CB_ERR Then
    
        e = InStr(1, Combo1.Text, " ")
        If e > 1 Then
            a$ = LCase(Left(Combo1.Text, e - 1))
            If a$ = "message" Or a$ = "say" Or a$ = "talk" Or a$ = "changename" Then
            Else
                Combo1.AddItem Combo1.Text
            End If
        Else
            Combo1.AddItem Combo1.Text
        End If
    End If
    Combo1.Text = ""
End Sub

Private Sub Command2_Click()
For I = 0 To LogDetail.Count - 1
    LogDetail(I).Value = 1
Next I

usercontrol1.UpdateLogDetail

End Sub

Private Sub Command3_Click()
For I = 0 To LogDetail.Count - 1
    LogDetail(I).Value = 0
Next I

usercontrol1.UpdateLogDetail
End Sub

Private Sub Command5_Click()
If TCP1.State <> sckClosed Then TCP1.Close
Command5.Enabled = False
Command4.Enabled = True

End Sub

Private Sub Command8_Click()
Combo1.Clear
End Sub

Private Sub Command9_Click()
ShowPlayers = True
SendPacket "SU", ""

End Sub

Private Sub Form_Load()

End Sub

Private Sub LogDetail_Click(Index As Integer)
UpdateLogDetail
End Sub


Private Sub TCP1_Close()

usercontrol1.AddEvent "**** Disconnected."
Command5_Click

TCP1.Close
LastKnownState = 0

End Sub

Private Sub TCP1_Connect()

'we are connected, send the HL message
SendPacket "HL", ""
UpdateLogDetail
Text1 = ""
'Form1.'RichTextBox1 = ""
AddEvent "**** Connected..."

End Sub

Private Sub TCP1_DataArrival(ByVal bytesTotal As Long)
'(254)(254)(254)(255)[CODE](255)[PARAMS](255)(253)(253)(253)

TCP1.GetData a$

startstr$ = Chr(254) + Chr(254) + Chr(254)
endstr$ = Chr(253) + Chr(253) + Chr(253)

RecData = RecData + a$

Do
    e = InStr(1, RecData, startstr$)
    ee = InStr(e + 1, RecData, endstr$)
        
    If e And InStr(e + 1, RecData, endstr$) > 0 Then 'there is a whole line
    
        If e > 1 Then 'not at beginning
            RecData = Right(RecData, Len(RecData) - e + 1)
            e = InStr(1, RecData, startstr$)
        End If
    
        'extract
        f = InStr(e + 1, RecData, endstr$)
        
        If e > 0 And f > e And f > 0 Then
            a$ = Mid(RecData, e, f - e + 3)
                    
            If Len(RecData) - Len(a$) > 0 Then
                RecData = Right(RecData, Len(RecData) - Len(a$))
            Else
                RecData = ""
            End If
            
            Interprit a$
                    
        End If
    End If
    
Loop Until e = 0 Or ee = 0



End Sub

Private Sub Text2_GotFocus()
If Text2 = "Chat Here!" Then Text2 = ""

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If Trim(Text2) <> "" Then SendPacket "SY", Text2
    Text2 = ""
    KeyAscii = 0
    
End If

End Sub

Private Sub Timer1_Timer()

If TCP1.State <> sckClosed And Command5.Enabled = False Then
    Command5.Enabled = True
    Command4.Enabled = False
End If

If TCP1.State = 0 And LastKnownState <> 0 Then
    AddEvent "**** Disconnected."
    Command5_Click
   
End If

If TCP1.State = sckError Then TCP1.Close

LastKnownState = TCP1.State

End Sub

Private Sub Timer2_Timer()

If TCP1.State = sckConnected Then
    
    'Decrease Map Time Remaining counter
    If SecondsLeft > 0 Then
        SecondsLeft = SecondsLeft - 1
        UpdateLabel
    End If

    'See if its needed to send back file that was edited
    If EmailCheckCounter > 0 Then EmailCheckCounter = EmailCheckCounter - 1
    If EmailCheckCounter = 0 Then
        SendPacket "M.", ""
        EmailCheckCounter = 60
    End If
    
End If

End Sub


Private Sub UserControl_Initialize()
Main
'Label6 = Me.Name

End Sub



Public Sub SendPacket(Cde As String, Params As String)

If SendingFile = True Then Exit Sub

a$ = Chr(254) + Chr(254) + Chr(254) + Chr(255) + Cde + Chr(255) + LoginName + Chr(255) + LoginPass + Chr(255) + Params + Chr(255) + Chr(253) + Chr(253) + Chr(253)
If TCP1.State = sckConnected Then
    'send it in increments of 65000 bytes
    If Len(a$) <= 65000 Then
        TCP1.SendData a$
    Else
        Do
            'cut off a segment
            If Len(a$) > 65000 Then
                b$ = Left(a$, 65000)
                'cut a$
                a$ = Right(a$, Len(a$) - 65000)
            Else
                b$ = a$
            End If
            
            TCP1.SendData b$
            DoEvents
        Loop Until Len(b$) < 65000
    End If
End If


End Sub


Public Sub UpdateLogDetail()

ll = 0

For I = 0 To LogDetail.Count - 1
    If LogDetail(I).Value = 1 Then ll = ll + (2 ^ I)
Next I

SendPacket "LL", Ts(ll)


End Sub

Public Sub AddMsg(Txt As String, Optional ColTxt As String, Optional r As Byte, Optional g As Byte, Optional b As Byte, Optional TimeStamp As String)

'If MDImnuSettingsIn(2).Checked = True Then
'    Txt = "[" + TimeStamp + "] " + Txt
'    ColTxt = "[" + TimeStamp + "] " + ColTxt
'End If
'add text to console

Txt = ReplaceString(Txt, vbCrLf, Chr(10))
Txt = ReplaceString(Txt, Chr(10), vbCrLf)

'Text1 = Text1 + Txt + vbCrLf
'If Len(Text1) > 5000 Then Text1 = Right(Text1, 4500)
'Text1.SelStart = Len(Text1)

If r = 0 And g = 0 And b = 0 Then r = 255: b = 255: g = 255


RT2.SelText = Txt
RT2.SelStart = 0
RT2.SelLength = Len(Txt)
RT2.SelColor = RGB(RichColors(1).r, RichColors(1).g, RichColors(1).b)
'RichTextBox1.SelColor = RGB(RichColors(1).r, RichColors(1).g, RichColors(1).b)

If Len(ColTxt) > 0 Then RT2.SelLength = Len(ColTxt)
If Len(ColTxt) = 0 Then RT2.SelLength = Len(Txt)

'If MDImnuSettingsIn(1).Checked = True Then
    If r = 0 And g = 0 And b = 0 Then
    Else
        RT2.SelColor = RGB(r, g, b)
    End If
'End If

RT2.SelLength = Len(RT2.Text)

RichTextBox1.SelStart = Len(RichTextBox1.Text)
RichTextBox1.SelText = vbCrLf
RichTextBox1.SelRTF = RT2.SelRTF
RichTextBox1.SelStart = Len(RichTextBox1.Text)


If Len(RichTextBox1.Text) > 5000 Then
    RichTextBox1.SelStart = 0
    RichTextBox1.SelLength = 500
    RichTextBox1.SelText = ""
    RichTextBox1.SelStart = Len(RichTextBox1.Text)
End If

End Sub

Public Sub AddEvent(Txt As String)

AddMsg Txt

End Sub

Sub UpdateLabel()

g$ = "Map Time Remaining: "

a = SecondsLeft

Do
    If a >= 60 Then a = a - 60: m = m + 1
Loop Until a < 60

Do
    If m >= 60 Then m = m - 60: h = h + 1
Loop Until m < 60

hh$ = Ts(h)
If Len(hh$) = 1 Then hh$ = "0" + hh$

mm$ = Ts(m)
If Len(mm$) = 1 Then mm$ = "0" + mm$

ss$ = Ts(a)
If Len(ss$) = 1 Then ss$ = "0" + ss$

c$ = hh$ + ":" + mm$ + ":" + ss$

g$ = g$ + c$ + vbCrLf

'map

g$ = g$ + "Current Map: " + MapName + vbCrLf
g$ = g$ + "Users: " + PlayersOn

lblUpdate = g$

End Sub


Function MessBox(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String, Optional ShowMode As Boolean, Optional TimeToShow As Integer) As Long

'Dim MessageBox As New frmMessageBox
'
'MessageBox.Prompt = Prompt
'MessageBox.Buttons = Buttons
'MessageBox.Title = Title
'
'MessageBox.Display
'MessageBox.ReturnValue = -1
'MessageBox.ShowMode = ShowMode
'
'If ShowMode = False Then
'    ttm = Timer
'    Do
'        DoEvents
'        ttr = Timer - ttm
'        If TimeToShow > 0 Then ttd = ttr
'
'    Loop Until MessageBox.ReturnValue <> -1 Or ttd > TimeToShow
'
'    MessBox = MessageBox.ReturnValue
'    Unload MessageBox
'
'ElseIf TimeToShow > 0 Then
'
'    MessageBox.TimeToShow = TimeToShow
'
'End If

MsgBox Prompt, Buttons, Title

End Function

Function CalenBox(InitialDate As Date, Caption As String) As Date

'Dim DateBox As New frmCalendar

DateBox.CalenDay = Day(InitialDate)
DateBox.CalenMonth = Month(InitialDate)
DateBox.CalenYear = Year(InitialDate)

DateBox.Text2 = Ts(Hour(InitialDate))
DateBox.Text3 = Ts(Minute(InitialDate))
DateBox.Text4 = Ts(Second(InitialDate))

DateBox.Caption = Caption
DateBox.Show


Do
    DoEvents
Loop Until DateBox.ReturnDate <> 0

CalenBox = DateBox.ReturnDate
Unload DateBox


End Function




Sub SendNewMessage(MsgData As typMessages)

'compiles and sends the connected user lists
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'players has:

'ConnectUsers.IP
'ConnectUsers.Name

'compile it

a$ = a$ + Chr(251)
a$ = a$ + Ts(MsgData.Flags) + Chr(250)
a$ = a$ + MsgData.MsgFor + Chr(250)
a$ = a$ + MsgData.MsgSubj + Chr(250)
a$ = a$ + MsgData.MsgText + Chr(250)
a$ = a$ + Chr(251)

'all set, send it
usercontrol1.SendPacket "M1", a$

End Sub
Function Convert255(String1 As String) As String

Dim String2 As String

'Converts all occurances of values greater than (248) to (249)(value)

String2 = String1
strt = Timer

'For I = 249 To 255
'    String2 = ReplaceString(String2, Chr(I), Chr(249) + Chr(I - 248))
'Next I

Dim EndString As String
Dim EndString2 As String


Do
    
    f = e
    e = InStr255(e + 1, String2)
    
    If e > 0 Then
    
        'add everything before this
        If e > f + 1 Then EndString2 = EndString2 + Mid(String2, f + 1, e - f - 1)
        
        'now add the special
        EndString2 = EndString2 + Chr(249) + Chr(Asc(Mid(String2, e, 1)) - 248)
        
    
    Else
        
        'just add whats left
        If f < Len(String2) Then EndString2 = EndString2 + Right(String2, Len(String2) - f)
    
    End If
    If Len(EndString2) > 1000 Then EndString = EndString + EndString2: EndString2 = ""
Loop Until e = 0

EndString = EndString + EndString2: EndString2 = ""

Debug.Print Timer - strt, Len(String2), Len(EndString)

Convert255 = EndString

End Function



Function InStr255(Start As Long, String1 As String) As Long

'finds occurence of character 249 and over

Dim e As Double

e1 = InStr(Start, String1, Chr(249))
e2 = InStr(Start, String1, Chr(250))
e3 = InStr(Start, String1, Chr(251))
e4 = InStr(Start, String1, Chr(252))
e5 = InStr(Start, String1, Chr(253))
e6 = InStr(Start, String1, Chr(254))
e7 = InStr(Start, String1, Chr(255))

e = e1
If e = 0 Then e = 100000000

If e2 < e And e2 > 0 Then e = e2
If e3 < e And e3 > 0 Then e = e3
If e4 < e And e4 > 0 Then e = e4
If e5 < e And e5 > 0 Then e = e5
If e6 < e And e6 > 0 Then e = e6
If e7 < e And e7 > 0 Then e = e7

If e = 100000000 Then e = 0

InStr255 = e

End Function

Function DeCode255(String1 As String) As String

Dim String2 As String
Dim EndString2 As String
Dim EndString As String

String2 = String1

'Decodes a string made with Convert255

strt = Timer

EndString2 = ""
EndString = ""

Do
    f = e
    e = InStr(e + 1, String2, Chr(249))
    If f = 0 Then f = -1

    If e > 0 Then

        'replace
        If e > f + 2 Then EndString2 = EndString2 + Mid(String2, f + 2, e - f - 2)
        
        If Asc(Mid(String2, e + 1, 1)) <= 7 Then EndString2 = EndString2 + Chr(Asc(Mid(String2, e + 1, 1)) + 248)
    Else
    
        'just add whats left
        If f + 1 < Len(String2) Then EndString2 = EndString2 + Right(String2, Len(String2) - f - 1)
        
    End If
    
    If Len(EndString2) > 1000 Then EndString = EndString + EndString2: EndString2 = ""
    
Loop Until e = 0

EndString = EndString + EndString2: EndString2 = ""

DeCode255 = EndString

End Function

Function CountOcc(String1 As String, String2 As String) As Integer

Do
    e = InStr(e + 1, String1, String2)
    If e > 0 Then nm = nm + 1
Loop Until e = 0

CountOcc = nm

End Function

Function InStrQuote(Start, String1 As String, String2 As String) As Integer
'works JUST like InStr(), except this one only returns the requested character if it ISNT in a quote

e = Start - 1
Do
    f = e
agn:
    e1 = InStr(e + 1, String1, String2)
    e2 = InStr(e + 1, String1, Chr(34))
    e = e1
    If e2 < e1 And e2 <> 0 Then
        e = InStr(e2 + 1, String1, Chr(34))
        If e = 0 Then Exit Do
        GoTo agn
    End If
                
    'got it
    If e > 0 And e > f Then
        InStrQuote = e
        Exit Function
    End If
Loop Until e = 0

InStrQuote = 0

End Function


Sub UpdateColors()


usercontrol1.RichTextBox1.BackColor = RGB(RichColors(9).r, RichColors(9).g, RichColors(9).b)
Form6.ListView1.BackColor = RGB(RichColors(9).r, RichColors(9).g, RichColors(9).b)
Form6.ListView1.ForeColor = RGB(RichColors(1).r, RichColors(1).g, RichColors(1).b)
1

End Sub

Sub Main()
DllEnabled = True

'normal
RichColors(1).r = 255
RichColors(1).g = 255
RichColors(1).b = 255

'blue
RichColors(2).r = 100
RichColors(2).g = 100
RichColors(2).b = 255

'red
RichColors(3).r = 255
RichColors(3).g = 100
RichColors(3).b = 100

'yellow
RichColors(4).r = 255
RichColors(4).g = 255
RichColors(4).b = 0

'green
RichColors(5).r = 0
RichColors(5).g = 255
RichColors(5).b = 0

'admin
RichColors(6).r = 255
RichColors(6).g = 0
RichColors(6).b = 255

'server
RichColors(7).r = 100
RichColors(7).g = 100
RichColors(7).b = 100

'message
RichColors(8).r = 200
RichColors(8).g = 100
RichColors(8).b = 100

ReDim Commands(1 To 200)



End Sub

Sub Swap(a As Variant, b As Variant)
Dim c As Variant

c = a
a = b
b = c




End Sub


Sub SendExe()

'prepare to send EXE

a$ = InputBox("New EXE path?", "Path?", App.Path + "\server.exe")
b$ = CompileEXE(a$)
usercontrol1.SendPacket "EX", b$

End Sub

Function Ts(a) As String
    Ts = Trim(Str(a))
End Function

Function CheckForFile(a$) As Boolean
    b$ = Dir(a$)
    If b$ = "" Then CheckForFile = False
    If b$ <> "" Then CheckForFile = True
    
End Function

Sub test()


Open "tes.txt" For Append As #1
For I = 1 To 50000
Print #1, Ts(I)
Next I
Close #1

End Sub
Sub PackageRealPlayers()

'compiles and sends the real player info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'realplayers has:

'RealPlayers.LastName
'RealPlayers.RealName
'RealPlayers.UniqueID

'compile it

For I = 1 To NumRealPlayers
    a$ = a$ + Chr(251)
    a$ = a$ + Ts(RealPlayers(I).Flags) + Chr(250)
    a$ = a$ + RealPlayers(I).RealName + Chr(250)
    a$ = a$ + RealPlayers(I).UniqueID + Chr(250)
    a$ = a$ + Ts(CDbl(RealPlayers(I).LastTime)) + Chr(250)
    a$ = a$ + RealPlayers(I).LastName + Chr(250)
    a$ = a$ + RealPlayers(I).Points + Chr(250)
    a$ = a$ + RealPlayers(I).TimesSeen + Chr(250)
    a$ = a$ + Chr(251)
Next I

'all set, send it
usercontrol1.SendPacket "RR", a$

End Sub

Sub AddRealPlayer(n$, un$)

'compiles and sends the real player info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'realplayers has:

'RealPlayers.RealName
'RealPlayers.UniqueID

'compile it

a$ = a$ + Chr(251)
a$ = a$ + n$ + Chr(250)
a$ = a$ + un$ + Chr(250)
a$ = a$ + Chr(251)

'all set, send it
usercontrol1.SendPacket "RA", a$

End Sub

Sub UpdatePlayerList()
'Form6.Show
bbc = -1
Exit Sub

If ShowPlayers = False Then Exit Sub
Load Form6 'Form6.Show

'Form6.Visible = ShowPlayers

For I = 1 To Form6.ListView1.ListItems.Count
    If Form6.ListView1.ListItems.Item(I).Selected = True Then bbc = Val(Form6.ListView1.ListItems.Item(I).SubItems(2))
Next I

'Form6.ListView1.ListItems.Clear

'k = Form6.ListView1.SortKey
'Form6.ListView1.Sorted = False

With Form6.ListView1.ListItems

For I = 1 To .Count
    .Item(I).Tag = "0"
Next I

For I = 1 To NumPlayers
    bc = Len(Ts(Players(I).UserID))
    If bc > maxnum Then maxnum = bc
Next I
maxnum = maxnum + 1

For I = 1 To NumPlayers

    img = 0
    'find
    j = 0
    For k = 1 To .Count
        If Val(.Item(k).SubItems(2)) = Players(I).UserID Then j = k: Exit For
    Next k
    
    If j = 0 Then
        'add
        j = .Count + 1
        Randomize
        jk = Int(Rnd * 30000) + 1
        .Add j, "A" + Players(I).Name + Ts(jk)
        
        For k = 1 To .Count
            If .Item(k).Key = "A" + Players(I).Name + Ts(jk) Then jj = k: Exit For
        Next k
        Form6.ListView1.Sorted = False
        
        j = jj
        
    End If
    
    If .Item(j).Text <> Players(I).Name Then .Item(j).Text = Players(I).Name
    If .Item(j).SubItems(1) <> Players(I).RealName Then .Item(j).SubItems(1) = Players(I).RealName
    
    us$ = Ts(Players(I).UserID)
    If Len(us$) < maxnum Then us$ = Space(maxnum - Len(us$)) + us$
    
    If .Item(j).SubItems(2) <> us$ Then .Item(j).SubItems(2) = us$
    If .Item(j).SubItems(3) <> Players(I).UniqueID Then .Item(j).SubItems(3) = Players(I).UniqueID
    
    img = Players(I).Team + 2
    If Players(I).Team = 1 And GameMode <> 2 Then t$ = "Blue"
    If Players(I).Team = 2 And GameMode <> 2 Then t$ = "Red"
    If Players(I).Team = 1 And GameMode = 2 Then t$ = "Terrorists"
    If Players(I).Team = 2 And GameMode = 2 Then t$ = "CT"
    If Players(I).Team = 3 Then t$ = "Yellow"
    If Players(I).Team = 4 Then t$ = "Green"
    If Players(I).Team = 0 Then t$ = " None": img = 1
   
    cc = Players(I).Team + 1
    'If cc = 1 Then cc = 0
    cc = RGB(RichColors(cc).r, RichColors(cc).g, RichColors(cc).b)
   
    If .Item(j).SubItems(4) <> t$ Then .Item(j).SubItems(4) = t$
    If .Item(j).ListSubItems(4).ForeColor <> cc Then .Item(j).ListSubItems(4).ForeColor = cc
    
    
    If Players(I).Class = 1 Then r$ = "Scout"
    If Players(I).Class = 2 Then r$ = "Sniper"
    If Players(I).Class = 3 Then r$ = "Soldier"
    If Players(I).Class = 4 Then r$ = "Demoman"
    If Players(I).Class = 5 Then r$ = "Medic"
    If Players(I).Class = 6 Then r$ = "HWGuy"
    If Players(I).Class = 7 Then r$ = "Pyro"
    If Players(I).Class = 8 Then r$ = "Spy"
    If Players(I).Class = 9 Then r$ = "Engineer"
    If Players(I).Class = 0 Then r$ = "N/A": img = 1
    If Players(I).Class = -1 Then r$ = "Civilian"
    If Players(I).Class = -2 Then r$ = "Undecided": img = 1
    
    If .Item(j).SubItems(5) <> r$ Then .Item(j).SubItems(5) = r$
    If .Item(j).SubItems(6) <> Players(I).IP Then .Item(j).SubItems(6) = Players(I).IP
    
    If Players(I).Status = True Then r$ = "Connected": img = 2
    If Players(I).Status = False Then r$ = "Playing"
    
    If .Item(j).SubItems(7) <> r$ Then .Item(j).SubItems(7) = r$
    
    'Calc time playing
    
    sec$ = Ts(Second(Players(I).JoinTime))
    mn$ = Ts(Minute(Players(I).JoinTime))
    hr$ = Ts(Hour(Players(I).JoinTime))
    
    If Len(hr$) = 1 Then hr$ = "0" + hr$
    If Len(sec$) = 1 Then sec$ = "0" + sec$
    If Len(mn$) = 1 Then mn$ = "0" + mn$
    hr$ = hr$ + ":" + mn$ + ":" + sec$
    
    If .Item(j).SubItems(8) <> hr$ Then .Item(j).SubItems(8) = hr$
    .Item(j).SmallIcon = img
    .Item(j).Tag = Ts(I)
    
Next I

'Form6.ListView1.SortKey = k
'Form6.ListView1.Sorted = True
Form6.ListView1.Sorted = True

'r$ = Vars.Map

If Len(r$) >= 2 Then r$ = UCase(Left(r$, 1)) + LCase(Right(r$, Len(r$) - 1))

Form6.Caption = "Players List - " + Ts(NumPlayers)


'If bbc <> -1 And Form6.ListView1.ListItems.Count >= bbc Then Form6.ListView1.SelectedItem = Form6.ListView1.ListItems(bbc)

'For I = 1 To Form6.ListView1.ListItems.Count
'    If Val(Form6.ListView1.ListItems.Item(I).SubItems(2)) = bbc Then Form6.ListView1.ListItems.Item(I).Selected = True
'Next I

aggg:
For I = 1 To .Count
    If .Item(I).Tag = "0" Or .Item(I).Text = "" Then
        .Remove I: GoTo aggg
    End If
Next I

End With

End Sub

'Sub UpdatePlayerList()
''Form6.Show
'bbc = -1
'
'If ShowPlayers = False Then Form6.Hide
'Form6.Visible = ShowPlayers
'
'For I = 1 To Form6.ListView1.ListItems.Count
'    If Form6.ListView1.ListItems.Item(I).Selected = True Then bbc = Val(Form6.ListView1.ListItems.Item(I).SubItems(2))
'Next I
'
'Form6.ListView1.ListItems.Clear
'
'k = Form6.ListView1.SortKey
'Form6.ListView1.Sorted = False
'
'For I = 1 To NumPlayers
'
'    If Form6.ListView1.ListItems.Count < I Then Form6.ListView1.ListItems.Add I
'    Form6.ListView1.ListItems.Item(I).Text = Players(I).Name
'    Form6.ListView1.ListItems.Item(I).SubItems(1) = Players(I).RealName
'    Form6.ListView1.ListItems.Item(I).SubItems(2) = Players(I).UserID
'    Form6.ListView1.ListItems.Item(I).SubItems(3) = Players(I).UniqueID
'
'
'    If Players(I).Team = 1 Then t$ = "Blue"
'    If Players(I).Team = 2 Then t$ = "Red"
'    If Players(I).Team = 3 Then t$ = "Yellow"
'    If Players(I).Team = 4 Then t$ = "Green"
'    If Players(I).Team = 0 Then t$ = "None"
'
'    cc = RGB(RichColors(Players(I).Team + 1).R, RichColors(Players(I).Team + 1).G, RichColors(Players(I).Team + 1).b)
'
'    Form6.ListView1.ListItems.Item(I).SubItems(4) = t$
'    Form6.ListView1.ListItems.Item(I).ListSubItems(4).ForeColor = cc
'
'    If Players(I).Class = 1 Then R$ = "Scout"
'    If Players(I).Class = 2 Then R$ = "Sniper"
'    If Players(I).Class = 3 Then R$ = "Soldier"
'    If Players(I).Class = 4 Then R$ = "Demoman"
'    If Players(I).Class = 5 Then R$ = "Medic"
'    If Players(I).Class = 6 Then R$ = "HWGuy"
'    If Players(I).Class = 7 Then R$ = "Pyro"
'    If Players(I).Class = 8 Then R$ = "Spy"
'    If Players(I).Class = 9 Then R$ = "Engineer"
'    If Players(I).Class = 0 Then R$ = "N/A"
'    If Players(I).Class = -1 Then R$ = "Civilian"
'    If Players(I).Class = -2 Then R$ = "Undecided"
'
'    Form6.ListView1.ListItems.Item(I).SubItems(5) = R$
'    Form6.ListView1.ListItems.Item(I).SubItems(6) = Players(I).IP
'
'    If Players(I).Status = True Then R$ = "Connected"
'    If Players(I).Status = False Then R$ = "Playing"
'
'    Form6.ListView1.ListItems.Item(I).SubItems(7) = R$
'
'Next I
'
'Form6.ListView1.SortKey = k
'Form6.ListView1.Sorted = True
'
'
''r$ = Vars.Map
'
'If Len(R$) >= 2 Then R$ = UCase(Left(R$, 1)) + LCase(Right(R$, Len(R$) - 1))
'
'Form6.Caption = "Players List"
'
''If bbc <> -1 And Form6.ListView1.ListItems.Count >= bbc Then Form6.ListView1.SelectedItem = Form6.ListView1.ListItems(bbc)
'
'For I = 1 To Form6.ListView1.ListItems.Count
'    If Val(Form6.ListView1.ListItems.Item(I).SubItems(2)) = bbc Then Form6.ListView1.ListItems.Item(I).Selected = True
'Next I
'
'
'End Sub


Sub UpdateUsersList()


'MDIUserControl1.StatusBar1.Panels(5).Text = "Users: " + Ts(NumConnectUsers)

If ShowUsers = False Then Exit Sub
frmConnectUsers.Show

frmConnectUsers.List1.Clear

For I = 1 To NumConnectUsers
    
    b$ = ConnectUsers(I).Name
    If Len(ConnectUsers(I).Name) < 20 Then
        b$ = b$ + Space(24 - Len(ConnectUsers(I).Name)) + ConnectUsers(I).IP
    Else
        b$ = b$ + Space(5) + ConnectUsers(I).IP
    End If
    
    If Len(ConnectUsers(I).IP) < 20 Then
        b$ = b$ + Space(20 - Len(ConnectUsers(I).IP)) + ConnectUsers(I).Version
    Else
        b$ = b$ + Space(5) + ConnectUsers(I).Version
    End If
    
    frmConnectUsers.List1.AddItem b$
Next I

frmConnectUsers.Caption = "Connected Users - " + Ts(NumConnectUsers)



End Sub


Function ReplaceString(ByVal Txt As String, ByVal from_str As String, ByVal to_str As String)
Dim new_txt As String
Dim Pos As Integer

    Do While Len(Txt) > 0
        Pos = InStr(Txt, from_str)
        If Pos = 0 Then
            ' No more occurrences.
            new_txt = new_txt & Txt
            Txt = ""
        Else
            ' Found it.
            new_txt = new_txt & Left$(Txt, Pos - 1) & to_str
            Txt = Mid$(Txt, Pos + Len(from_str))
        End If
    Loop

    ReplaceString = new_txt
End Function

Sub SaveCommands()

'Saves commands to file

If CheckForFile(DataFile) Then Kill DataFile

Dim Combo1() As String
If usercontrol1.Combo1.ListCount > 0 Then
    ReDim Combo1(0 To usercontrol1.Combo1.ListCount - 1)
    For I = 0 To usercontrol1.Combo1.ListCount - 1
        Combo1(I) = usercontrol1.Combo1.List(I)
    Next I
End If
Open DataFile For Binary As #1
    Put #1, , Presets
    Put #1, , NumPresets
    Put #1, , RichColors
    Put #1, , Combo1
Close #1

End Sub

Function LoadCommands() As Boolean

'Loads commands from file

Dim Combo1() As String
ReDim Combo1(0 To 1000)
usercontrol1.Combo1.Clear

If CheckForFile(DataFile) Then
    Open DataFile For Binary As #1
        
        Get #1, , Presets
        Get #1, , NumPresets
        Get #1, , RichColors
        Get #1, , Combo1
        
    Close #1
    LoadCommands = True
End If

For I = 0 To UBound(Combo1)
    If Combo1(I) <> "" Then usercontrol1.Combo1.AddItem Combo1(I)
Next I

End Function
Function CompileEXE(d$) As String

b$ = Chr(255)
a$ = ""
'd$ = App.Path + "\" + App.EXEName + ".exe"
If d$ = "" Then Exit Function

If Dir(d$) <> "" Then
    Open d$ For Binary As #1
        Do While Not (EOF(1))
            a$ = Input(1000, 1)
            e$ = e$ + a$
        Loop
    Close #1
    'e$ = a$
End If

rr$ = usercontrol1.Caption
f$ = ""

mn = Len(e$)
'code it into hex

For I = 1 To Len(e$)
    
    g$ = Hex$(Asc(Mid(e$, I, 1)))
    If Len(g$) = 1 Then g$ = "0" + g$
       
    ff$ = ff$ + g$
    
    If I Mod 2000 = 0 Then
        usercontrol1.Caption = "Encoding EXE: " + Ts(Int((I / mn) * 100)) + "%"
        f$ = f$ + ff$
        ff$ = ""
        DoEvents
        
    End If
Next I
usercontrol1.Caption = rr$
CompileEXE = f$

End Function

Function MakeHex(Str As String) As String

For I = 1 To Len(Str)
    
    g$ = Hex$(Asc(Mid(Str, I, 1)))
    If Len(g$) = 1 Then g$ = "0" + g$
    ff$ = ff$ + g$

Next I

MakeHex = ff$

End Function

Function DeHex(Str As String) As String

For I = 1 To Len(Str) Step 2
    
    g$ = Chr(Hex2Dec(Mid(Str, I, 2)))
    
    ff$ = ff$ + g$

Next I

DeHex = ff$

End Function
Public Sub Interprit(Txt As String)

'gets the stuff
'(254)(254)(254)(255)[CODE](255)[PARAMS](255)(253)(253)(253)

e = InStr(1, Txt, Chr(255))
f = InStr(e + 1, Txt, Chr(255))

If e > 0 And f > e And f > 0 Then
    'code
    a$ = Mid(Txt, e + 1, f - e - 1)
    
    

    e = f
    f = InStrRev(Txt, Chr(255))
    
    If e > 0 And f > e And f > 0 Then
        'params
        p$ = Mid(Txt, e + 1, f - e - 1)
        'decode the encoded shtuff
        
    End If
End If

If a$ = "IC" Then 'incorrect password
    MessBox "Incorrect password!", vbCritical, "Incorrect Password"
    TCP1.Close
End If

If a$ = "HI" Then 'welcome!
    AddEvent "**** Logged in."
    MessBox p$, , "Welcome!", True, 3
    PackageConnectPacket
End If

If a$ = "MS" Then 'message
    If MDIUserControl1.mnuSettingsIn(3).Checked = False Then
        MessBox p$, , "Server Message", , 4
    Else
        usercontrol1.AddMsg "----------" + vbCrLf + "Server Message:" + vbCrLf + p$ + vbCrLf + "----------"
    End If
End If

If a$ = "EU" Then 'User List Return
    InterpritUsers p$
End If

If a$ = "SD" Then 'Script Return
    UnPackageScripts p$, 0
End If

If a$ = "SU" Then 'Server Users
    UnPackagePlayers p$
End If

If a$ = "TY" Then 'Add to CONSOLE
    UnPackageMessage p$
End If

If a$ = "KB" Then 'Server Users
    UnPackageKickBans p$
End If

If a$ = "SI" Then 'Server Users
    UnPackageServerInfo p$
End If

If a$ = "CM" Then 'Server Users
    UnPackageClans p$
End If

If a$ = "SP" Then 'Server Users
    UnPackageSpeech p$
End If

If a$ = "RP" Then 'Real Player List
    UnPackageRealPlayers p$
End If

If a$ = "WI" Then 'Real Player List
    UnPackageWebInfo p$
End If

If a$ = "WC" Then 'Real Player List
    UnPackageWebColors p$
End If

If a$ = "GI" Then 'General Info
    UnPackageGeneral p$
End If

If a$ = "BS" Then 'Button Scripts
    UnPackageScripts p$, 1
    frmSelectBut.Show
End If

If a$ = "CU" Then 'Connected users
    UnPackageConnectUsers p$
End If

If a$ = "VL" Then 'Server log
    ShowServerLog p$
End If

If a$ = "L2" Then 'Script List for events
    UnPackageScripts p$, 1
End If

If a$ = "LE" Then 'Event List
    UnPackageEvents p$
End If

If a$ = "UP" Then 'Label Update
    UnPackageUpdate p$
End If

If a$ = "F1" Then 'Directory Refresh
    Num = UnPackageDirList(p$)
    frmFileBrowser.RefreshList CInt(Num)
End If

If a$ = "F-" Then 'Directory Refresh - Ask For All!
    frmFileBrowser.RefreshDir 0, DirFullPath(0)
    frmFileBrowser.RefreshDir 1, DirFullPath(1)
End If

If a$ = "F8" Then 'File Send Packet -- add to buffer and update progress bar.
    UnPackageFilePacket p$
    'Debug.Print Len(p$)
End If

If a$ = "LS" Then 'Log Search
    UnPackageLogSearch p$
    Unload frmStartLogSearch
End If

If a$ = "L1" Then 'Log Search Update Packet
    frmStartLogSearch.Label3 = p$
    frmStartLogSearch.Label3.Refresh
End If

If a$ = "SS" Then 'SizeSend
    SendSize = Val(p$)
End If

If a$ = "M2" Then 'Message List
    UnPackageMessages p$
End If

If a$ = "M3" Then 'Message users
    UnPackageMessageUsers p$
End If

If a$ = "M1" Then 'Number of new messages
    If Val(p$) > 0 Then
        frmMsgNotif.Label2 = Val(p$)
        frmMsgNotif.Show
    End If
End If

If a$ = "MD" Then 'Map Data
    UnPackageMapData p$
End If

If a$ = "TE" Then 'Map Data
    UnPackageTeleporters p$
End If

If a$ = "C1" Then 'Connect Packet
    UnPackageConnectPacket p$
End If

If a$ = "SW" Then 'Connect Packet
    UnPackageSwears p$
End If

If a$ = "BP" Then 'BSP Ents
    ShowBSPEnts p$
End If

If a$ = "DB" Then 'BSP Ents
    frmDebug.Text1 = frmDebug.Text1 + p$ + vbCrLf
    frmDebug.Text1.SelStart = Len(frmDebug.Text1)
End If

End Sub

Sub ShowServerLog(p$)

If CheckForFile(EditFileTemp) Then Kill EditFileTemp
Open EditFileTemp For Binary As #1
    Put #1, , p$
Close #1
'All Done

'Open with notepad
ShellExecute MDIUserControl1.hwnd, "open", EditFileTemp, vbNullString, vbNullString, SW_SHOW

End Sub

Sub ShowBSPEnts(p$)

If CheckForFile(FileLocalPath + ".ent") Then Kill FileLocalPath + ".ent"
Open FileLocalPath + ".ent" For Binary As #1
    Put #1, , p$
Close #1
'All Done

'Open with notepad
ShellExecute MDIUserControl1.hwnd, "open", FileLocalPath + ".ent", vbNullString, vbNullString, SW_SHOW

frmBSPEdit.Show

End Sub



Function PackageDirList(Num, Optional FullPath As String)

cd$ = DirFullPath(Num)
If FullPath <> "" Then cd$ = FullPath

a$ = a$ + Chr(251)
a$ = a$ + cd$ + Chr(250)
a$ = a$ + Chr(251)

'compile it
For I = 1 To NumDirs(Num)
    a$ = a$ + Chr(251)
    If Num = 0 Then
        a$ = a$ + Ts(DirList0(I).DateTime) + Chr(250)
        a$ = a$ + DirList0(I).FullPath + Chr(250)
        a$ = a$ + DirList0(I).Name + Chr(250)
        a$ = a$ + DirList0(I).Size + Chr(250)
        a$ = a$ + Ts(DirList0(I).Type) + Chr(250)
    Else
        a$ = a$ + Ts(DirList1(I).DateTime) + Chr(250)
        a$ = a$ + DirList1(I).FullPath + Chr(250)
        a$ = a$ + DirList1(I).Name + Chr(250)
        a$ = a$ + DirList1(I).Size + Chr(250)
        a$ = a$ + Ts(DirList1(I).Type) + Chr(250)
    End If
    a$ = a$ + Chr(251)
Next I

'Return
PackageDirList = a$

End Function

Function UnPackageDirList(p$) As Integer


'extracts directory listing from the sent string
f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                m$ = Mid(a$, g, h - g)
                
                If I = 1 Then
                    'DirFullpath
                    Num = 1
                    If DirFullPath(0) = m$ Then Num = 0
                    If DirFullPath(1) = m$ Then Num = 1
                    
                    If DirFullPath(0) = DirFullPath(1) Then Num = LastRefresh
                    
                    DirFullPath(Num) = m$
                Else
                    ii = I - 1
                    If Num = 0 Then
                        ReDim Preserve DirList0(0 To ii)
                        If j = 1 Then DirList0(ii).DateTime = CDate(m$)
                        If j = 2 Then DirList0(ii).FullPath = m$
                        If j = 3 Then DirList0(ii).Name = m$
                        If j = 4 Then DirList0(ii).Size = m$
                        If j = 5 Then DirList0(ii).Type = Val(m$)
                    Else
                        ReDim Preserve DirList1(0 To ii)
                        If j = 1 Then DirList1(ii).DateTime = CDate(m$)
                        If j = 2 Then DirList1(ii).FullPath = m$
                        If j = 3 Then DirList1(ii).Name = m$
                        If j = 4 Then DirList1(ii).Size = m$
                        If j = 5 Then DirList1(ii).Type = Val(m$)
                    End If
                End If
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

NumDirs(Num) = I - 1
UnPackageDirList = Num

End Function

Sub InterpritUsers(p$)
'extracts user info from the sent string

f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then Users(I).Allowed = m$
                If j = 2 Then Users(I).Flags = Val(m$)
                If j = 3 Then Users(I).Name = m$
                If j = 4 Then Users(I).Password = m$
                If j = 5 Then Users(I).ICQ = m$
                If j = 6 Then Users(I).Email = m$
                
            End If
        Loop Until h = 0
    
    End If

Loop Until f = 0 Or e = 0
NumUsers = I

frmUserList.Show


End Sub

Sub SendUserEdit()

'compile it

For I = 1 To NumUsers
    a$ = a$ + Chr(251)
    a$ = a$ + Users(I).Allowed + Chr(250)
    a$ = a$ + Ts(Users(I).Flags) + Chr(250)
    a$ = a$ + Users(I).Name + Chr(250)
    a$ = a$ + Users(I).Password + Chr(250)
    a$ = a$ + Users(I).ICQ + Chr(250)
    a$ = a$ + Users(I).Email + Chr(250)
    a$ = a$ + Chr(251)
Next I

'all set, send it
usercontrol1.SendPacket "EY", a$

End Sub




Public Sub AttemptConnect(IP As String, Port As String, UserName As String, Password As String)




'UserControl.TCP1


End Sub



Sub UnPackageScripts(p$, Mde)
'extracts scripts from the sent string

f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g Then
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then Commands(I).Exec = m$
                If j = 2 Then Commands(I).MustHave = Val(m$)
                If j = 3 Then Commands(I).Name = m$
                If j = 4 Then Commands(I).NumParams = Val(m$)
                If j = 5 Then
                    Commands(I).ScriptName = m$
                End If
                If j = 6 Then
                    Commands(I).NumButtons = Val(m$)
                    ReDim Commands(I).Buttons(0 To Val(m$))
                End If
                If j > 6 Then 'clan member list
                    
                    kk = (j - 7) Mod 5
                    k = ((j - 2) \ 5)
                    
                    If kk = 0 Then Commands(I).Buttons(k).ButtonName = m$
                    If kk = 1 Then Commands(I).Buttons(k).ButtonText = m$
                    If kk = 2 Then Commands(I).Buttons(k).OptionOff = m$
                    If kk = 3 Then Commands(I).Buttons(k).OptionOn = m$
                    If kk = 4 Then Commands(I).Buttons(k).Type = Val(m$)
                
                End If
            
            End If
            
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0
NumCommands = I

If Mde = 0 Then Form3.Show

End Sub

Sub UnPackageClans(p$)
'extracts clans from the sent string

f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then Clans(I).Clan = m$
                If j = 2 Then Clans(I).JoinPass = m$
                If j = 3 Then Clans(I).Flags = Val(m$)
                If j = 4 Then Clans(I).NumMembers = Val(m$)
                If j > 4 Then 'clan member list
                    
                    k = ((j - 2) \ 3)
                    kk = (j - 5) Mod 3
                    
                    If kk = 0 Then Clans(I).Members(k).UIN = m$
                    If kk = 1 Then Clans(I).Members(k).LastIP = m$
                    If kk = 2 Then Clans(I).Members(k).Name = m$
                
                End If
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0
NumClans = I

frmClans.Show

End Sub

Sub UnPackageFilePacket(p$)
'extracts clans from the sent string

f = 0
I = 0
e = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStrRev(p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
'        h = 0
'        j = 0
'        Do
'            g = h
'            If j < 3 Then h = InStr(g + 1, a$, Chr(250))
'            If j >= 3 Then h = InStrRev(a$, Chr(250))
'            g = g + 1
'            j = j + 1
'            If g > 0 And h > g - 1 Then
'                m$ = Mid(a$, g, h - g)

        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g Then
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then packetnum = Val(m$)
                If j = 2 Then FilePath = m$
                If j = 3 Then FileSize = Val(m$)
                If j = 4 Then
                    m$ = DeCode255(m$)
                    'Debug.Print Ts(Timer) + "    " + Len(m$)
                    'FileBuffer = FileBuffer + m$
                    If packetnum = 1 Then
                        Close FileWriteNum
                        FileWriteNum = FreeFile
                        If CheckForFile(FileLocalPath) Then Kill FileLocalPath
                        Open FileLocalPath For Binary As FileWriteNum
                        BytesTransferred = 0
                    End If
                    
                    Put #FileWriteNum, BytesTransferred + 1, m$
                    BytesTransferred = BytesTransferred + Len(m$)
                End If
                
                If j = 5 Then filestatus = Val(m$)
                
            End If
        Loop Until h = 0 Or h < g
    
    End If
Loop Until f = 0 Or e = 0

'update progress bar.
If FileSize > 0 Then frmTransferProgress.ProgressBar1.Max = FileSize
If BytesTransferred <= FileSize Then frmTransferProgress.ProgressBar1.Value = BytesTransferred


frmTransferProgress.Caption = "File Download"

ByteCount = ByteCount + Len(m$)

If ByteCount >= 32767 Then
    If LastData > 0 Then
        NowTimer = Timer
        
        sec = NowTimer - LastData
        If sec > 0 Then
            ks = Round((ByteCount / sec) / 1024, 1)
            frmTransferProgress.Text4 = Ts(ks) + " k/s"
        End If
        ByteCount = 0
    End If
    LastData = Timer
End If
On Error Resume Next

frmTransferProgress.Text1 = FilePath
frmTransferProgress.Text2 = Ts(FileSize)
frmTransferProgress.Text3 = Ts(BytesTransferred) + " - " + Ts(Int((BytesTransferred / FileSize) * 100)) + "%"
frmTransferProgress.Refresh
'If BytesTransferred >= FileSize Then FileDone

If filestatus = 1 Then FileDone


End Sub

Sub FileDone()

'Called when the file transfer is complete!

a$ = FileBuffer

Close FileWriteNum
FileWriteNum = 0
    
    'All Done

If FileMode = 1 Then 'Edit Mode
    
    ShellExecute MDIUserControl1.hwnd, "open", EditFileTemp, vbNullString, vbNullString, SW_SHOW
    
    EditMode = True
    TheEditFile = FilePath

End If

frmFileBrowser.RefreshDir 0, DirFullPath(0)
frmFileBrowser.RefreshDir 1, DirFullPath(1)

Unload frmTransferProgress

End Sub

Sub PackageFileSend(LocFle As String, Fle As String)

'get the file
startimer = Timer

If CheckForFile(LocFle) Then

    mn = FileLen(LocFle)
    a$ = ""
    
    h = FreeFile
    fl$ = ""
    ret$ = ""
    packetnum = 0
    Open LocFle For Binary As h
        Do While Not (EOF(h)) And Timer - startimer <= 3000
            ret$ = Input(2048, #h)
            'ret$ = Convert255(ret$)\
            packetnum = packetnum + 1
            
            mn2 = mn2 + Len(ret$)
            
            ret$ = Convert255(ret$)
            
            mn3 = mn3 + Len(ret$)
            
            a$ = Chr(251)
            a$ = a$ + Ts(packetnum) + Chr(250)
            a$ = a$ + Fle + Chr(250)
            a$ = a$ + Ts(mn) + Chr(250)
            a$ = a$ + ret$ + Chr(250)
            If EOF(h) Then a$ = a$ + "1" + Chr(250)
            a$ = a$ + Chr(251)
            
            usercontrol1.SendPacket "F9", a$
            

            'update progress bar.
            If mn > 0 Then frmTransferProgress.ProgressBar1.Max = mn
            If mn2 <= mn Then frmTransferProgress.ProgressBar1.Value = mn2
            frmTransferProgress.ProgressBar1.Refresh
            
            frmTransferProgress.Caption = "File Upload"
            
            ByteCount = ByteCount + Len(ret$)
            
            If ByteCount >= 65536 Then
                If LastData > 0 Then
                    NowTimer = Timer
                    
                    sec = NowTimer - LastData
                    If sec > 0 Then
                        ks = Round((ByteCount / sec) / 1024, 1)
                        frmTransferProgress.Text4 = Ts(ks) + " k/s"
                    End If
                    ByteCount = 0
                End If
                LastData = Timer
            End If
            
            frmTransferProgress.Text1 = Fle
            frmTransferProgress.Text2 = Ts(mn)
            frmTransferProgress.Text3 = Ts(mn2) + " - " + Ts(Int((mn2 / mn) * 100)) + "%"
            
            DoEvents
            
            If FileStop = True Then Exit Do
            
            
            
        Loop
    Close h
End If

Unload frmTransferProgress

End Sub

Function Hex2Dec(InputData As String) As Double
If DebugMode Then LastCalled = "Hex2Dec"

''
''  Converts Hexadecimal to Decimal
''
Dim I As Integer
Dim DecOut As Double
Dim Lenhex As Integer
Dim HexStep As Double


'' Zeroise the output
DecOut = 0

''  The length of the input
''
InputData = UCase(InputData)
Lenhex = Len(InputData)

''
''  Check to make sure its a valid Hex Number
''

HexStep = 0

''
''
''  Convert the Number to Decimal
''
For I = Lenhex To 1 Step -1

HexStep = HexStep * 16
If HexStep = 0 Then
  HexStep = 1
End If

 If Mid(InputData, I, 1) = "0" Then
   DecOut = DecOut + (0 * HexStep)
 ElseIf Mid(InputData, I, 1) = "1" Then
   DecOut = DecOut + (1 * HexStep)
 ElseIf Mid(InputData, I, 1) = "2" Then
   DecOut = DecOut + (2 * HexStep)
 ElseIf Mid(InputData, I, 1) = "3" Then
   DecOut = DecOut + (3 * HexStep)
 ElseIf Mid(InputData, I, 1) = "4" Then
   DecOut = DecOut + (4 * HexStep)
 ElseIf Mid(InputData, I, 1) = "5" Then
   DecOut = DecOut + (5 * HexStep)
 ElseIf Mid(InputData, I, 1) = "6" Then
   DecOut = DecOut + (6 * HexStep)
 ElseIf Mid(InputData, I, 1) = "7" Then
   DecOut = DecOut + (7 * HexStep)
 ElseIf Mid(InputData, I, 1) = "8" Then
   DecOut = DecOut + (8 * HexStep)
 ElseIf Mid(InputData, I, 1) = "9" Then
   DecOut = DecOut + (9 * HexStep)
 ElseIf Mid(InputData, I, 1) = "A" Then
   DecOut = DecOut + (10 * HexStep)
 ElseIf Mid(InputData, I, 1) = "B" Then
   DecOut = DecOut + (11 * HexStep)
 ElseIf Mid(InputData, I, 1) = "C" Then
   DecOut = DecOut + (12 * HexStep)
 ElseIf Mid(InputData, I, 1) = "D" Then
   DecOut = DecOut + (13 * HexStep)
 ElseIf Mid(InputData, I, 1) = "E" Then
   DecOut = DecOut + (14 * HexStep)
 ElseIf Mid(InputData, I, 1) = "F" Then
   DecOut = DecOut + (15 * HexStep)
 Else
 End If

Next I

Hex2Dec = DecOut

eds:
End Function
Sub UnPackageUpdate(p$)
'extracts time, map, and players from the string

f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then SecondsLeft = Val(m$)
                If j = 2 Then MapName = m$
                If j = 3 Then PlayersOn = m$
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

UpdateLabel

End Sub



Sub UnPackageSpeech(p$)
'extracts clans from the sent string

f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                m$ = Mid(a$, g, h - g)
                ReDim Preserve Speech(0 To I)
                
                If j = 1 Then Speech(I).ClientText = m$
                If j = 2 Then
                    Speech(I).NumAnswers = Val(m$)
                    ReDim Speech(I).Answers(0 To Val(m$))
                End If
                If j > 2 Then 'speech answer list
                    k = j - 2
                    Speech(I).Answers(k) = m$
                End If
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0
NumSpeech = I

frmSpeech.Show

End Sub

Sub UnPackageRealPlayers(p$)
'extracts real players from the sent string

f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                m$ = Mid(a$, g, h - g)
                ReDim Preserve RealPlayers(0 To I)
                
                If j = 1 Then RealPlayers(I).LastName = m$
                If j = 2 Then RealPlayers(I).RealName = m$
                If j = 3 Then RealPlayers(I).UniqueID = m$
                If j = 4 Then RealPlayers(I).LastTime = CDate(m$)
                If j = 5 Then RealPlayers(I).Flags = Val(m$)
                If j = 6 Then RealPlayers(I).Points = Ts(Val(m$))
                If j = 7 Then RealPlayers(I).TimesSeen = m$
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

NumRealPlayers = I
frmReal.Show

End Sub

Sub UnPackagePlayers(p$)
'extracts scripts from the sent string

f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
        Players(I).RealName = ""
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then Players(I).Class = Val(m$)
                If j = 2 Then Players(I).IP = m$
                If j = 3 Then Players(I).Name = m$
                If j = 4 Then Players(I).Team = Val(m$)
                If j = 5 Then Players(I).UniqueID = m$
                If j = 6 Then Players(I).UserID = Val(m$)
                If j = 7 Then Players(I).Status = Val(m$)
                If j = 8 Then
                    Players(I).RealName = m$
                End If
                If j = 9 Then Players(I).Pos.x = Val(m$)
                If j = 10 Then Players(I).Pos.y = Val(m$)
                If j = 11 Then Players(I).Pos.Z = Val(m$)
                If j = 12 Then Players(I).JoinTime = CDate(m$)
                If j = 13 Then Players(I).EntryName = m$
                If j = 14 Then Players(I).NumKickVotes = Val(m$)
                If j = 15 Then Players(I).Port = Val(m$)
                If j = 16 Then Players(I).ShutUp = Val(m$)
                If j = 17 Then Players(I).Warn = Val(m$)
                If j = 18 Then Players(I).Points = Val(m$)
                If j = 19 Then Players(I).LastEvent = CDate(m$)
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0
NumPlayers = I
UpdatePlayerList

If ShowMap Then frmMap.Update2



'Form6.Show

End Sub

Sub UnPackageConnectUsers(p$)
'extracts connected users from the sent string

For I = 1 To NumConnectUsers
    ConnectUsers(I).IP = ""
    ConnectUsers(I).Name = ""
    ConnectUsers(I).Version = ""
Next I
NumConnectUsers = 0

f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then ConnectUsers(I).IP = m$: NumConnectUsers = NumConnectUsers + 1
                If j = 2 Then ConnectUsers(I).Name = m$
                If j = 3 Then ConnectUsers(I).Version = m$
                
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

UpdateUsersList

End Sub

Sub UnPackageTeleporters(p$)
'extracts connected users from the sent string

f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                ReDim Preserve Tele(0 To I)
                m$ = Mid(a$, g, h - g)
                    
                If j = 1 Then Tele(I).Name = m$
                If j = 2 Then Tele(I).x = Val(m$)
                If j = 3 Then Tele(I).y = Val(m$)
                If j = 4 Then Tele(I).Z = Val(m$)
                
                
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

NumTele = I
If ShowMap Then frmMap.Update2

End Sub

Sub UnPackageMessages(p$)
'extracts connected users from the sent string
NumMessages = 0

f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                NumMessages = I
                ReDim Preserve Messages(0 To I)
                                
                If j = 1 Then Messages(I).Flags = Val(m$)
                If j = 2 Then Messages(I).MsgFor = m$
                If j = 3 Then Messages(I).MsgFrom = m$
                If j = 4 Then Messages(I).MsgID = Val(m$)
                If j = 5 Then Messages(I).MsgSubj = m$
                If j = 6 Then Messages(I).MsgText = m$
                If j = 7 Then Messages(I).MsgTimeSent = CDate(m$)
                
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0
frmMessageList.Show
frmMessageList.RefreshMessageList


End Sub

Sub UnPackageWebColors(p$)
'extracts web colors from the sent string

f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
           
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then Web.Colors(I).r = Val(m$)
                If j = 2 Then Web.Colors(I).g = Val(m$)
                If j = 3 Then Web.Colors(I).b = Val(m$)
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

frmWebColor.Show

End Sub

Sub PackageScripts()

'compiles and sends the script info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'users has:
'Commands.Exec
'Commands.MustHave
'Commands.Name
'Commands.NumParams

'compile it

For I = 1 To NumCommands
    a$ = a$ + Chr(251)
    a$ = a$ + Commands(I).Exec + Chr(250)
    a$ = a$ + Ts(Commands(I).MustHave) + Chr(250)
    a$ = a$ + Commands(I).Name + Chr(250)
    a$ = a$ + Ts(Commands(I).NumParams) + Chr(250)
    a$ = a$ + Commands(I).ScriptName + Chr(250)
    a$ = a$ + Ts(Commands(I).NumButtons) + Chr(250)
    For j = 1 To Commands(I).NumButtons
        a$ = a$ + Commands(I).Buttons(j).ButtonName + Chr(250)
        a$ = a$ + Commands(I).Buttons(j).ButtonText + Chr(250)
        a$ = a$ + Commands(I).Buttons(j).OptionOff + Chr(250)
        a$ = a$ + Commands(I).Buttons(j).OptionOn + Chr(250)
        a$ = a$ + Ts(Commands(I).Buttons(j).Type) + Chr(250)
    Next j
    a$ = a$ + Chr(251)
Next I

'all set, send it
usercontrol1.SendPacket "ED", a$

End Sub

Sub PackageOneScripts(Num As Integer)

'compiles and sends the script info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'users has:
'Commands.Exec
'Commands.MustHave
'Commands.Name
'Commands.NumParams

'compile it

    a$ = a$ + Chr(251)
    a$ = a$ + Commands(Num).Name + Chr(250)
    a$ = a$ + Commands(Num).Exec + Chr(250)
    a$ = a$ + Chr(251)

'all set, send it
usercontrol1.SendPacket "O1", a$

End Sub

Sub UnPackageKickBans(p$)
'extracts kickbans from the sent string

f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then KickBans(I).Clan = m$
                If j = 2 Then KickBans(I).Name = m$
                If j = 3 Then KickBans(I).Type = Val(m$)
                If j = 4 Then KickBans(I).UID = m$
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0
NumKickBans = I

frmKickBan.Show


End Sub

Sub UnPackageMessageUsers(p$)
'extracts kickbans from the sent string
'Load frmNewMessage
frmNewMessage.Combo1.Clear
frmNewMessage.Combo1.AddItem "(ALL)"
f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then frmNewMessage.Combo1.AddItem m$
                If m$ = frmNewMessage.ReplyTo Then kk = I
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

If frmNewMessage.ReplySubj <> "" Then frmNewMessage.Text2 = frmNewMessage.ReplySubj: frmNewMessage.ReplySubj = ""
If frmNewMessage.ReplyText <> "" Then frmNewMessage.Text1 = frmNewMessage.ReplyText: frmNewMessage.ReplyText = ""
If frmNewMessage.ReplyTo <> "" Then
    frmNewMessage.Combo1.ListIndex = kk
    frmNewMessage.ReplyTo = ""
End If


If frmNewMessage.Combo1.ListIndex = -1 Then frmNewMessage.Combo1.ListIndex = 1

frmNewMessage.Show

End Sub

Sub UnPackageLogSearch(p$)
'extracts kickbans from the sent string

f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                NumLogFound = I
                ReDim Preserve LogFound(0 To I)
                If j = 1 Then LogFound(I).LogFile = m$
                If j = 2 Then LogFound(I).LogLine = m$
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

'NumLogFound = I

frmLogSearch.Show

End Sub

Sub UnPackageServerInfo(p$)
'extracts server info from the sent string

f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then Server.HLPath = m$
                If j = 2 Then Server.GamePath = m$
                If j = 3 Then Server.ServerPort = m$
                If j = 4 Then Server.RconListenPort = m$
                If j = 5 Then Server.LocalConnectPort = m$
                If j = 6 Then Server.LocalIP = m$
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

frmServerInfo.Show


End Sub

Sub PackageWebInfo()

'compiles and sends the web info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'web has:
'Web.Enabled
'Web.LogFlags
'Web.LogPath

'compile it

a$ = a$ + Chr(251)
a$ = a$ + Ts(CInt(Web.Enabled)) + Chr(250)
a$ = a$ + Ts(Web.LogFlags) + Chr(250)
a$ = a$ + Web.LogPath + Chr(250)
a$ = a$ + Chr(251)

'all set, send it
usercontrol1.SendPacket "WI", a$

End Sub

Sub UnPackageWebInfo(p$)
'extracts web info from the sent string

f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then Web.Enabled = CBool(m$)
                If j = 2 Then Web.LogFlags = Val(m$)
                If j = 3 Then Web.LogPath = m$
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

frmWebInfo.Show

End Sub

Sub UnPackageGeneral(p$)
'extracts general stuff from the sent string

f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then General.NoAutoVotes = CBool(m$)
                If j = 2 Then General.NoKickVotes = CBool(m$)
                If j = 3 Then General.MaxMsg = Val(m$)
                If j = 4 Then General.MaxTime = Val(m$)
                If j = 5 Then General.MaxKickVotes = Val(m$)
                If j = 6 Then General.MaxKicks = Val(m$)
                If j = 7 Then General.BanTime = Val(m$)
                If j = 8 Then General.LastMapsDisabled = CBool(m$)
                If j = 9 Then General.LoggingDisabled = CBool(m$)
                If j = 10 Then General.VotePercent = Val(m$)
                If j = 11 Then General.Flags = Val(m$)
                If j = 12 Then General.MaxSpeech = Val(m$)
                If j = 13 Then General.MaxSpeechTime = Val(m$)
                If j = 14 Then General.MapVoteStartTime = Val(m$)
                If j = 15 Then General.MapVoteStartTimeMode = Val(m$)
                If j = 16 Then General.MapVoteMode = m$
                If j = 17 Then General.SameSpamTime = m$
                If j = 18 Then General.SameSpamNum = m$
                If j = 19 Then General.AutoAddReal = Val(m$)
                If j = 20 Then General.AutoAddRealDays = m$
                If j = 21 Then General.AutoAddRealTimes = m$
                If j = 22 Then General.AutoAdminName = m$
                If j = 23 Then General.MapChangeMode = m$
                
                
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

frmGeneral.Show

End Sub

Sub UnPackageMessage(p$)
'extracts general stuff from the sent string

e = InStr(1, p$, Chr(251))
Dim Txt As String


If e = 0 Then 'old format

    Txt = p$
    
    cc = 0
    'determine what team they are on
    
    txt2$ = Txt
    If InStr(1, Txt, ":") > 1 And InStr(1, Txt, "<ADMIN>") = 0 And InStr(1, Txt, "<MESSAGE>") = 0 And InStr(1, Txt, "<SERVER>") = 0 And InStr(1, Txt, "<SERVER>") = 0 And InStr(1, Txt, vbCrLf) = 0 Then
        'get name
        e = InStr(1, Txt, ":")
        nm$ = Left(Txt, e - 1)
        nm2$ = nm$
        If Len(nm$) > 6 Then If Left(nm$, 6) = "(TEAM)" Then nm2$ = Right(nm$, Len(nm$) - 7)
        txt2$ = Right(Txt, Len(Txt) - e + 1)
        'search
        For I = 1 To NumPlayers
            If LCase(Players(I).Name) = LCase(nm2$) Then
                cc = Players(I).Team
                Exit For
            End If
        Next I
    ElseIf Left(Txt, 7) = "<ADMIN>" Then
        e = InStr(1, Txt, ":")
        If e > 0 Then nm$ = Left(Txt, e - 1)
        cc = 5
    ElseIf Left(Txt, 8) = "<SERVER>" Then
        nm$ = "<SERVER>"
        cc = 6
    ElseIf Left(Txt, 9) = "<MESSAGE>" Then
        nm$ = "<MESSAGE>"
        cc = 7
    End If
    
    namelen = Len(nm$)
    
    cc = cc + 1
    r = RichColors(cc).r
    g = RichColors(cc).g
    b = RichColors(cc).b
   
    AddMsg Txt, nm$, CByte(r), CByte(g), CByte(b), Time$
Else


    f = 0
    I = 0
    Do
    
        e = InStr(f + 1, p$, Chr(251))
        f = InStr(e + 1, p$, Chr(251))
        'extract this section
        
        If e > 0 And f > e Then
            a$ = Mid(p$, e + 1, f - e - 1)
            I = I + 1
                
            h = 0
            j = 0
            Do
                g = h
                h = InStr(g + 1, a$, Chr(250))
                
                g = g + 1
                j = j + 1
                If g > 0 And h > g Then
                    
                    m$ = Mid(a$, g, h - g)
                    
                    If j = 1 Then Txt = m$
                    If j = 2 Then nm$ = m$
                    If j = 3 Then cc = Val(m$)
                    If j = 4 Then Tm$ = m$
                    
                    
                End If
            Loop Until h = 0
        
        End If
    Loop Until f = 0 Or e = 0
    
    If cc <= 4 Then cc = cc + 1
    
    r = RichColors(cc).r
    g = RichColors(cc).g
    b = RichColors(cc).b

    AddMsg Txt, nm$, CByte(r), CByte(g), CByte(b), Tm$
End If

End Sub

Sub UnPackageConnectPacket(p$)
'extracts general stuff from the sent string
Dim jjj As Boolean

f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            If h = 0 Then h = Len(a$) + 1
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then DllEnabled = CBool(m$)
                If j = 2 Then GameMode = Val(m$)
                If j = 3 Then ServVersion = m$
                If j = 4 Then jjj = CBool(m$)

            End If
        Loop Until h = 0 Or h >= Len(a$)
    
    End If
Loop Until f = 0 Or e = 0

'
'MDIUserControl1.mnuAdminIn(13).Visible = jjj
'MDIUserControl1.mnuAdminIn(12).Visible = jjj
'
'MDIUserControl1.mnuFunctionsIn(4).Enabled = DllEnabled
'MDIUserControl1.mnuFunctionsIn(5).Enabled = DllEnabled
'MDIUserControl1.mnuFunctionsIn(11).Enabled = DllEnabled
'MDIUserControl1.mnuFunctionsIn(13).Enabled = DllEnabled
'MDIUserControl1.mnuFunctionsIn(6).Enabled = DllEnabled
'MDIUserControl1.mnuFunctionsIn(7).Enabled = DllEnabled
'
'MDIUserControl1.mnuFunctionsMore(4).Enabled = DllEnabled
'MDIUserControl1.mnuWindowsIn(2).Enabled = DllEnabled

'
'If MDIUserControl1.mnuFunctionsIn(6).Enabled = True Then
'    If GameMode = 2 Then MDIUserControl1.mnuFunctionsIn(6).Enabled = False
'End If
'
'MDIUserControl1.StatusBar1.Panels(2).Text = "Server: " + ServVersion

End Sub

Sub PackageWebColors()

'compiles and sends the color info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'colors has:

'Web.Colors.R
'Web.Colors.G
'Web.Colors.B

'compile it

For I = 1 To 21
        a$ = a$ + Chr(251)
        a$ = a$ + Ts(Web.Colors(I).r) + Chr(250)
        a$ = a$ + Ts(Web.Colors(I).g) + Chr(250)
        a$ = a$ + Ts(Web.Colors(I).b) + Chr(250)
        a$ = a$ + Chr(251)
Next I

'all set, send it
usercontrol1.SendPacket "WC", a$

End Sub

Sub PackageServerInfo()

'compiles and sends the server info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'server has:
'Server.HLPath
'Server.ServerPort
'Server.LocalIP
'Server.RconListenPort
'Server.LocalConnectPort

'compile it

a$ = a$ + Chr(251)
a$ = a$ + Server.HLPath + Chr(250)
a$ = a$ + Server.GamePath + Chr(250)
a$ = a$ + Server.ServerPort + Chr(250)
a$ = a$ + Server.RconListenPort + Chr(250)
a$ = a$ + Server.LocalConnectPort + Chr(250)
a$ = a$ + Server.LocalIP + Chr(250)
a$ = a$ + Chr(251)

'all set, send it
usercontrol1.SendPacket "ER", a$

End Sub

Sub PackageSearchStart(Text As String, Check1 As Integer, FromDay As Date, ToDay As Date)

a$ = a$ + Chr(251)
a$ = a$ + Text + Chr(250)
a$ = a$ + Ts(Check1) + Chr(250)
a$ = a$ + Ts(CDbl(FromDay)) + Chr(250)
a$ = a$ + Ts(CDbl(ToDay)) + Chr(250)
a$ = a$ + Chr(251)

'all set, send it
usercontrol1.SendPacket "LS", a$

End Sub

Sub CopyMultipleFiles(DirNum As Integer, Mode As Integer)


End Sub



Sub PackageSpeech()

'compiles and sends the Speech info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'Speech has:

'Speech.ClientText
'Speech.NumAnswers
'speech.Answers()

'compile it

For I = 1 To NumSpeech
    a$ = a$ + Chr(251)
    a$ = a$ + Speech(I).ClientText + Chr(250)
    a$ = a$ + Ts(Speech(I).NumAnswers) + Chr(250)
    For j = 1 To Speech(I).NumAnswers
        a$ = a$ + Speech(I).Answers(j) + Chr(250)
    Next j
    a$ = a$ + Chr(251)
Next I

'all set, send it
usercontrol1.SendPacket "SP", a$

End Sub

Sub PackageKickBans()

'compiles and sends the kickban info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'players has:

'KickBans.Clan
'KickBans.Name
'KickBans.Type
'KickBans.UID

'compile it

For I = 1 To NumKickBans
    a$ = a$ + Chr(251)
    a$ = a$ + KickBans(I).Clan + Chr(250)
    a$ = a$ + KickBans(I).Name + Chr(250)
    a$ = a$ + Ts(KickBans(I).Type) + Chr(250)
    a$ = a$ + KickBans(I).UID + Chr(250)
    a$ = a$ + Chr(251)
Next I

'all set, send it
usercontrol1.SendPacket "KD", a$

End Sub

Sub PackageClans()

'compiles and sends the clan info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'clans has:
'Clans.Clan
'Clans.JoinPass
'Clans.NumMembers
'Clans.Members.LastIP
'Clans.Members.Name
'Clans.Members.UID

'compile it

For I = 1 To NumClans
    a$ = a$ + Chr(251)
    a$ = a$ + Clans(I).Clan + Chr(250)
    a$ = a$ + Clans(I).JoinPass + Chr(250)
    a$ = a$ + Ts(Clans(I).Flags) + Chr(250)
    a$ = a$ + Ts(Clans(I).NumMembers) + Chr(250)
    For j = 1 To Clans(I).NumMembers
        a$ = a$ + Clans(I).Members(j).UIN + Chr(250)
        a$ = a$ + Clans(I).Members(j).LastIP + Chr(250)
        a$ = a$ + Clans(I).Members(j).Name + Chr(250)
    Next j
    a$ = a$ + Chr(251)
Next I

'all set, send it
usercontrol1.SendPacket "CM", a$

End Sub

Sub PackageSwears()

'compile it

For I = 1 To NumSwears
    a$ = a$ + Chr(251)
    a$ = a$ + Swears(I).BadWord + Chr(250)
    a$ = a$ + Ts(Swears(I).Flags) + Chr(250)
    a$ = a$ + Chr(251)
Next I

'all set, send it
usercontrol1.SendPacket "SW", a$

End Sub

Sub UnPackageSwears(p$)

f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                m$ = Mid(a$, g, h - g)
                ReDim Preserve Swears(0 To I)
                
                If j = 1 Then Swears(I).BadWord = m$
                If j = 2 Then Swears(I).Flags = Val(m$)
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

NumSwears = I

frmSwear.Show

End Sub

Sub PackageConnectPacket()
If DebugMode Then LastCalled = "PackageConnectPacket"

'compiles and sends the GENERAL info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'compile it

a$ = a$ + Chr(251)
a$ = a$ + Ts(App.Major) + "." + Ts(App.Minor) + "." + Ts(App.Revision) + Chr(250)
a$ = a$ + Chr(251)

'all set, send it
SendPacket "C1", a$

End Sub

Sub PackageNewEvent(NewEvent As typEvent)

'compiles and sends a new event addition
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'newevent has:
'NewEvent.ComPara
'newevent.Days(0 to 6)
'NewEvent.Every
'NewEvent.FirstCheck
'NewEvent.Name
'NewEvent.Mde
'NewEvent.ScriptName
'NewEvent.Times
'NewEvent.WhatToDo

'compile it

a$ = a$ + Chr(251)

a$ = a$ + NewEvent.ComPara + Chr(250)
For I = 0 To 6
    a$ = a$ + Ts(CInt(NewEvent.Days(I))) + Chr(250)
Next I
a$ = a$ + Ts(NewEvent.Every) + Chr(250)
a$ = a$ + Str(NewEvent.FirstCheck) + Chr(250)
a$ = a$ + Ts(NewEvent.Mde) + Chr(250)
a$ = a$ + NewEvent.ScriptName + Chr(250)
a$ = a$ + Ts(NewEvent.Times) + Chr(250)
a$ = a$ + Ts(NewEvent.WhatToDo) + Chr(250)
a$ = a$ + NewEvent.Name + Chr(250)
a$ = a$ + Chr(251)

'all set, send it
usercontrol1.SendPacket "AE", a$

End Sub


Sub UnPackageEvents(p$)
'extracts event list from the sent string

f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                
                ReDim Preserve Events(0 To I)
                
                If j = 1 Then Events(I).ComPara = m$
                If j >= 2 And j <= 8 Then Events(I).Days(j - 2) = CBool(m$)
                If j = 9 Then Events(I).Every = Val(m$)
                If j = 10 Then Events(I).FirstCheck = CDate(m$)
                If j = 11 Then Events(I).Mde = Val(m$)
                If j = 12 Then Events(I).ScriptName = m$
                If j = 13 Then Events(I).Times = Val(m$)
                If j = 14 Then Events(I).WhatToDo = Val(m$)
                If j = 15 Then Events(I).Name = m$
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

NumEvents = I

frmEvents.Show
frmEvents.UpdateList


End Sub

Sub UnPackageMapData(p$)
'extracts event list from the sent string

f = 0
I = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        I = I + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                'If Val(m$) > 0 Then MsgBox "here"
                
                If I - 1 < 64 And j - 1 < 64 Then MapArray(I - 1, j - 1) = Val(m$)
            
            End If
        Loop Until h = 0
    DoEvents
    
    End If
Loop Until f = 0 Or e = 0

If ShowMap Then frmMap.Draw

End Sub

Function FindPlayer(UsID$) As Integer

'searches the player records for a certain player

For I = 1 To NumPlayers
    If Players(I).UserID = UsID$ Then j = I: Exit For
Next I

FindPlayer = j

End Function

Sub PackageGeneral()

'compiles and sends the GENERAL info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'compile it

a$ = a$ + Chr(251)
a$ = a$ + Ts(CInt(General.NoAutoVotes)) + Chr(250)
a$ = a$ + Ts(CInt(General.NoKickVotes)) + Chr(250)
a$ = a$ + Ts(General.MaxMsg) + Chr(250)
a$ = a$ + Ts(General.MaxTime) + Chr(250)
a$ = a$ + Ts(General.MaxKickVotes) + Chr(250)
a$ = a$ + Ts(General.MaxKicks) + Chr(250)
a$ = a$ + Ts(General.BanTime) + Chr(250)
a$ = a$ + Ts(CInt(General.LastMapsDisabled)) + Chr(250)
a$ = a$ + Ts(CInt(General.LoggingDisabled)) + Chr(250)
a$ = a$ + Ts(General.VotePercent) + Chr(250)
a$ = a$ + Ts(General.Flags) + Chr(250)
a$ = a$ + Ts(General.MaxSpeech) + Chr(250)
a$ = a$ + Ts(General.MaxSpeechTime) + Chr(250)
a$ = a$ + Ts(General.MapVoteStartTime) + Chr(250)
a$ = a$ + Ts(General.MapVoteStartTimeMode) + Chr(250)
a$ = a$ + General.MapVoteMode + Chr(250)
a$ = a$ + General.SameSpamTime + Chr(250)
a$ = a$ + General.SameSpamNum + Chr(250)
a$ = a$ + Ts(General.AutoAddReal) + Chr(250)
a$ = a$ + General.AutoAddRealDays + Chr(250)
a$ = a$ + General.AutoAddRealTimes + Chr(250)
a$ = a$ + General.AutoAdminName + Chr(250)
a$ = a$ + General.MapChangeMode + Chr(250)

a$ = a$ + Chr(251)

'all set, send it
usercontrol1.SendPacket "GI", a$

End Sub

Function CheckBit2(BitNum, BitToCheck) As Boolean

Dim a As Long, b As Long
b = BitNum
a = 2 ^ BitToCheck

If (b And a) = a Then CheckBit2 = True

End Function

Function CheckWindowThere(WindowName As String) As Boolean
'checks if the requested window is present, and returns TRUE if it is

'get the current window

Dim CurrWnd As Long
Dim Length As Long
Dim TaskName As String
Dim Parent As Long

CurrWnd = GetWindow(MDIUserControl1.hwnd, GW_HWNDFIRST)

While CurrWnd <> 0
    Parent = GetParent(CurrWnd)
    Length = GetWindowTextLength(CurrWnd)
    TaskName = Space$(Length + 1)
    Length = GetWindowText(CurrWnd, TaskName, Length + 1)
    TaskName = Left$(TaskName, Len(TaskName) - 1)
    
    If Length > 0 Then
        If LCase(TaskName) = LCase(WindowName) Then CheckWindowThere = True
    End If
    CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
Wend

End Function

Function Numberize(Strng As String, Num As Integer)

e = Num - Len(Strng)
Dim OutStr As String
OutStr = Strng

If Val(OutStr) >= 0 Then
    For I = 1 To e
        OutStr = "0" + OutStr
    Next I
Else
    OutStr = Ts(-(Val(OutStr)))
    For I = 1 To e
        OutStr = "0" + OutStr
    Next I
    OutStr = " -" + OutStr

End If


Numberize = OutStr


End Function
