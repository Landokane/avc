Attribute VB_Name = "Module1"
Public Type typButtons
    ButtonName As String
    ButtonText As String
    OptionOn As String
    OptionOff As String
    Type As Integer
End Type

Public Type typConnectUsers
    Name As String
    IP As String
    Version As String
End Type

Public Type comds
    Name As String
    Exec As String
    NumParams As Integer
    MustHave As Integer
    ScriptName As String
    NumButtons As Integer
    Buttons() As typButtons
End Type

Public Type typPlayerPos
    x As Integer
    y As Integer
    Z As Integer
End Type

Public Type Playersss
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

Public Type typPreset
    Allowed As String
    Flags As Variant
    Name As String
End Type

Public Type typSpeech 'used when the server talks to the client
    ClientText As String 'what someone has to say
    Answers() As String 'what the server will do in return
    NumAnswers As Integer 'how many possibilities there are
End Type

Public Type OtherSettings
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

Public Type typUsers
    Name As String
    Password As String
    Allowed As String
    Flags As Variant
    ICQ As String
    Email As String
End Type

Public Type typRGB 'color set
    r As Byte
    g As Byte
    b As Byte
End Type

Public Type typKickBans 'people who are BANNED
    Name As String 'persons name
    Clan As String 'persons clan
    UID As String 'persons uniqueid
    Type As Byte    'type of ban:
                    '1 - kick this person
                    '2 - kick anyone in this persons clan
                    '4 - kick anyone with this UID
                    '8 - immidiately BAN this person by putting his uniqueid in the servers ban list
End Type

Public Type serv
    ServerPort As String
    GamePath As String
    HLPath As String
    LocalConnectPort As String
    RconListenPort As String
    LocalIP As String
End Type

Public Type typClanMember
    Name As String
    UIN As String
    LastIP As String
End Type

Public Type typClan
    'Provided to allow you to restrict the use of certain names, and identify players by uniqueid
    Clan As String
    JoinPass As String
    Flags As Integer
    NumMembers As Integer
    Members(1 To 200) As typClanMember
End Type

Public Type typRealPlayer
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

Public Type typSwearWords
    BadWord As String
    Flags As Long
End Type

Public Type typGeneral
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

Public Type typWebInfo
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

Public Type typEvent
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

Public Type typDirList ' Directory Listing
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

Public Type typLogFound
    LogFile As String
    LogLine As String
End Type

Public Type typMessages     'Kinda like email
    MsgText As String
    MsgFor As String
    MsgFrom As String
    MsgTimeSent As Date
    MsgSubj As String
    MsgID As Integer
    Flags As Integer        '1 - request reciept, 2 - unread
End Type

Public Type typTeleportExit
    x As Integer
    y As Integer
    Z As Integer
    Name As String
End Type

'Teleporters
Public Tele() As typTeleportExit
Public NumTele As Integer

'Messages
Public Messages() As typMessages
Public NumMessages As Integer

Public RichColors(0 To 10) As typRGB

'Log Searching
Public NumLogFound As Long
Public LogFound() As typLogFound
Public LogSearchString As String

Public SendSize As Long
Public FindReal As String

Public Swears() As typSwearWords
Public NumSwears As Integer

Public DllEnabled As Boolean
Public GameMode As Integer

Public ServVersion As String

Public LastKnownState As Integer


'Files and File Transfer Vars
Public DirList0() As typDirList
Public DirList1() As typDirList
Public NumDirs(0 To 1) As Long
Public DirFullPath(0 To 1) As String     'path to the current window
Public FileBuffer As String
Public FileMode As Integer               '0 - unset, 1 - edit this file, 2 - download this file
Public FilePath As String                'Path to download to for mode 2 above
Public FileLocalPath As String           'Where to save this file LOCALLY
Public FileSize As Long                  'Size of file
Public LastRefresh As Integer
Public FileRecieveMode As Boolean
Public SendingFile As Boolean            'Are we currently sending a file?
Public EditMode As Boolean              'Editing this file now?
Public TheEditFile As String               'File we are editing on server
Public EditFileTemp As String
Public FileWriteNum As Integer
Public BytesTransferred As Long
Public LastData As Long
Public ByteCount As Long
Public FileStop As Boolean


'Scheduler :)
Public Events() As typEvent
Public NumEvents As Integer

'time/update
Public SecondsLeft As Integer
Public MapName As String
Public PlayersOn As String

Public ShowPlayers As Boolean
Public ShowUsers As Boolean
Public ShowMap As Boolean

Public EmailCheckCounter As Integer

'web
Public Web As typWebInfo

'real players
Public RealPlayers() As typRealPlayer
Public NumRealPlayers As Integer

Public Server As serv
Public Presets(1 To 200) As typPreset
Public NumPresets As Integer

Public EditedButton As Integer

Public ConnectUsers(1 To 400) As typConnectUsers
Public NumConnectUsers As Integer

Public NumKickBans As Integer
Public KickBans(1 To 200) As typKickBans

Public Const SB_VERT = 1
Public Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Public Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Public Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
Public Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Public Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Const FO_MOVE As Long = &H1
Public Const FO_COPY As Long = &H2
Public Const FO_DELETE As Long = &H3
Public Const FO_RENAME As Long = &H4

Public Const FOF_MULTIDESTFILES As Long = &H1
Public Const FOF_CONFIRMMOUSE As Long = &H2
Public Const FOF_SILENT As Long = &H4
Public Const FOF_RENAMEONCOLLISION As Long = &H8
Public Const FOF_NOCONFIRMATION As Long = &H10
Public Const FOF_WANTMAPPINGHANDLE As Long = &H20
Public Const FOF_CREATEPROGRESSDLG As Long = &H0
Public Const FOF_ALLOWUNDO As Long = &H40
Public Const FOF_FILESONLY As Long = &H80
Public Const FOF_SIMPLEPROGRESS As Long = &H100
Public Const FOF_NOCONFIRMMKDIR As Long = &H200
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const SW_SHOW = 5


Public Const CB_ERR = (-1)
Public Const WM_USER = &H400
Public Const CB_FINDSTRING = &H14C

Public Commands() As comds
Public NumCommands As Integer

Public Players(1 To 400) As Playersss
Public NumPlayers As Integer

Public Users(1 To 200) As typUsers
Public NumUsers As Integer

Public General As typGeneral


'clannies
Public Clans(1 To 20) As typClan
Public NumClans As Integer

'speech
Public Speech() As typSpeech
Public NumSpeech As Integer

Public ChosenClan As Integer

Public Settings As OtherSettings
Public DataFile As String
Public DataFile2 As String
Public DataFile3 As String
Public RecData As String
Public LoginName As String
Public LoginPass As String
Public UserEditNum As Integer

'Map
Public MapArray(0 To 64, 0 To 64) As Integer ' Stores Z coord at this location


