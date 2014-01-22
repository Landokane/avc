Attribute VB_Name = "definitions"
' ---------------------------------------------------------------------------
' ===========================================================================
' ==========================     SERVER ASSISTANT     =======================
' ===========================================================================
'
'      This code is copyright © 1999-2003 Avatar-X (avcode@cyberwyre.com)
'      and is protected by the GNU General Public License.
'      Basically, this means if you make any changes you must distrubute
'      them, you can't keep the code for yourself.
'
'      A copy of the license was included with this download.
'
'      ------------------------------------------------------------------
'
'      FILE: defs.bas
'      PURPOSE: This file simply contains all the definitions for custom types,
'      global variables, APIs, etc which are used by SA.
'
'
' ===========================================================================
' ---------------------------------------------------------------------------

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'TYPES
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=


Public Type SHELLEXECUTEINFO
        cbSize As Long
        fMask As Long
        hwnd As Long
        lpVerb As String
        lpFile As String
        lpParameters As String
        lpDirectory As String
        nShow As Long
        hInstApp As Long
        '  Optional fields
        lpIDList As Long
        lpClass As String
        hkeyClass As Long
        dwHotKey As Long
        hIcon As Long
        hProcess As Long
End Type

Public Type typButtons
    ButtonName As String
    ButtonText As String
    OptionOn As String
    OptionOff As String
    Type As Integer
End Type

Public Type typSpeech 'used when the server talks to the client
    ClientText As String 'what someone has to say
    Answers() As String 'what the server will do in return
    NumAnswers As Integer 'how many possibilities there are
End Type

Public Type typAdminBMP

    AdminName As String
    BMPFile As String
    
    ' later expansion
    Unused1 As Integer
    unused2 As Integer
    
    unused3 As String
    unused4 As String
    
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

Public Type typMapRecord
    MapName As String
    TimesPlayed As Long
    LastTimePlayed As Date
    Unused1 As String
    unused2 As String
    unused3 As String
End Type

Public Type typScripts
    Name As String
    Exec As String
    NumParams As Integer
    MustHave As Integer
    ScriptName As String
    NumButtons As Integer
    Buttons() As typButtons
    Group As String
    LogExec As Boolean
    AutoMakeVars As Boolean
    Unused1 As String
    unused2 As String
    unused3 As String
    ScriptID As Integer
    Unused5 As Integer
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
    NumMembers As Integer
    Flags As Integer
    Members(1 To 200) As typClanMember
End Type

Public Type typBans
    UIDs As String
    Name As String
    IP As String
    Reason As String
    BanTime As String
    Map As String
    RealName As String
    EntryName As String
    BannedAt As String
End Type

Public Type serv
    RCONPass As String
    HostName As String
    LocalIP As String
    ServerPort As String
    GamePath As String
    HLPath As String
    BothPath As String
    DataFile As String
    LocalConnectPort As String
    RconListenPort As String
    GameMode As Integer     '1 = TFC mode, 2 = CS mode, 3 = OTHER MOD mode
End Type

Public Type typKills 'Weapons that can kill the player
    Name As String          'ie "Sniper Rifle"
    Ent As String           'ie "sniperrifle"
    Award As String         'ie "Best Sniper Award"
End Type

Public Type Varss
    Map As String
    UserIP As String
    UserPort As Long
    UserName As String
    UserTCP As Boolean 'is the user talking to us by TCP?
    Command As String
    ScriptCommand As String
    MapTimeLeft As Long
    MapTimeElapsed As Integer 'time elapsed in MINUTES
    MapTimeTotal As Integer ' the server setting for total map time
    ClanBattle As Boolean ' if true, many automated features stop working
    VotedMap As String 'the map that was voted for
    AlreadyAutoVoted As Boolean ' did we already auto-vote?
    TimeCounter As Integer      'Tick counter
    TimeCounter2 As Integer      'Tick counter
    MapTimeLeftIP As String
    MapTimeLeftPort As Long
    UserIsTCP As Boolean
    Index As Integer
    MaxPlayers As Integer
    MapCycleFile As String
    TempMapCycleFile As String
    TempFileMade As Boolean
    
End Type

Public Type typGeneral
    NoAutoVotes As Boolean      'disable automatic map votes
    NoKickVotes As Boolean      'disable automatic kick votes
    MaxMsg As Integer           'Maximum number of messages before kick (spam protection)
    MaxKickVotes As Integer     'Maximum number of kick votes
    MaxTime As Integer          'Time for MaxMsg
    MaxKicks As Integer         'Kickvoted this many times...
    BanTime As Integer          'Ban him this many minutes
    VotePercent As Integer      'Yes vote for a kickvote
    LoggingDisabled As Boolean  'Is logging enabled?
    LastMapsDisabled As Boolean 'Is last maps disabled?
    Flags As Long               'Checkable Flags
    MaxSpeech As Integer        'Max amount of speeches for x time
    MaxSpeechTime As Integer    'Time for above
    MapVoteStartTime As Integer 'Seconds from either start of map or end of map when auto vote starts
    MapVoteStartTimeMode As Integer ' 1 = time from start of map, 0 = time from end of map
    MapVoteMode As String       '1 - changelevel mode, 2 - mapcycle mode
    SameSpamTime As String      'Time allowance for sending the same message over again
    SameSpamNum As String       'Number of samespams allowed
    AutoAddReal As Integer      'Automatically Add RealPlayers
    AutoAddRealTimes As String  'Turn into real realplayers after seen more than this many times
    AutoAddRealDays As String   'Delete them if they are not seen after this many days
    AutoAdminName As String     'the name of the autoadmin
    MapChangeMode As String     'CHANGELEVEL or MAP
    CustomFlag1 As String       'Custom flags for RealPlayers
    CustomFlag2 As String       '2
    CustomFlag3 As String       '3
    CustomFlag4 As String       '4
    NewestClient As String      'Newest version of the client for autodownload
    MaxFileSend As Integer      'Max rate for sending files (to prevent lag) in K/s
    SendToDisco As Integer      'Send UDP Data to disco Stu Server
    UnUsed8 As Integer
    UnUsed9 As Integer
    UnUsed10 As Integer
End Type

Public Type typPlayerPos
    X As Integer
    Y As Integer
    Z As Integer
End Type

Public Type typLastMsgs
    Text As String
    When As Single
End Type

Public Type Playersss
    Name As String                 'name
    UniqueID As String
    IP As String
    Port As Long
    Class As Integer            'currentclass (-1 = civ, 0 = random, 1-9 = scout-engy) [only applies to TFC servers]
    UserID As Integer           'server userid
    Team As Integer             'current team
    RemoveMe As Boolean         'if the player will be needed later, DONT remove him!
    ConnectOnly As Boolean      'This player has connected but not yet entered the game
    ThereFlag As Boolean
    KillsWith() As Integer      'Kills Stats
    RealName As String          'the REAL name of this player, regardless of whatever crap they change their name to
    NumKickVotes As Integer     'Number of kickvotes this player has initiated
    EntryName As String         'The name they joined the game with
    BroadcastType As Byte       'Broadcasting to whole server. 0-off, 1-sa_talk or say, 2-message
    MessInMin As Integer        'Number of messages in last minute (spam protection)
    Pos As typPlayerPos
    ShutUp As Boolean           ' cant speak
    TimeJoined As Date
    Warn As Integer
    LastMsgs(1 To 10) As typLastMsgs
    TempRealMode As Boolean
    Points As Long
    LastEvent As Date           'The last time this player was heard from.
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

Public Type Votess
    UserID As String
    MapChoice As String
    Double As Boolean
End Type

Public Type typUsers
    Name As String
    PassWord As String
    Allowed As String
    Flags As Variant
    ICQ As String
    FTPRoot As String
End Type

Public Type typUsersOld
    Name As String
    PassWord As String
    Allowed As String
    Flags As Long
End Type

Public Type typUserRCON
    IP As String
    Port As Long
    Command As String
End Type

Public Type typConnectUsers
    Name As String
    PassWord As String
    DePass As String
    
    Index As Integer
    IP As String
    LogLevel As Integer
    UserNum As Integer
    SendingFile As Boolean
    FileRecieveMode As Boolean
    FileSavePath As String
    FileSize As Long
    FileNum As Long
    BytesTransferred As Long
    FileStop As Boolean 'flag to stop sending
    Version As String
    HiddenMode As Boolean
    Port As Long
    RemoveMe As Boolean
    
    'away mode stuff
    AwayMode As Integer 'Defines: 0 - not away, 1 - Away, 2 - N/A, 3 - Sleeping, 4 - InGame, 5 - Eating
    AwayMsg As String
    LastUserUpdate As Double
    LastUpdate As Double
    
    EncryptedMode As Boolean
    
    IdleTime As Long
    
End Type

Public Type typRconMonitor  'Used for ppl waiting for RCON results
    IsTCP As Boolean
    Index As Integer
    Port As Long
    IP As String
    Command As String
    TimeSent As Long
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

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    UID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Type typUserVars
    Name As String
    Value As String
End Type

Public Type typParams
    Value As String
End Type

Public Type typRealPlayerOLD
    RealName As String
    UniqueID As String
    LastName As String
    Points As String ' numba of points they got
    TimesSeen As String ' Number of times this player has been seen (once a day)
    LastTime As Double    'last time player was seen
    Flags As Long
End Type

Public Type typRealPlayer
    'this type is used to store information about certain players
    'they are identified by uniqueid, and can be found regardless of
    'their current name.
    RealName As String
    UniqueID As String
    LastName As String
    Points As String ' numba of points they got
    TimesSeen As String ' Number of times this player has been seen (once a day)
    LastTime As Double    'last time player was seen
    Flags As Long
    'UnUsed1 As Integer
    'UnUsed2 As Integer
    'UnUsed3 As String
    'Message As String
End Type

Public Type typRGB 'color set
    r As Byte
    G As Byte
    B As Byte
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
    CurrHTML As String          'The current HTML file
    Colors(1 To 21) As typRGB   'List of colours
End Type

Public Type typEvent
    mde As Integer '0 - once, 1 - once every xxx, 2 - once on these days
    Times As Integer 'the number of (every) to run
    Every As Integer '0 - weeks, 1 - days, 2 - hours, 3 - minutes, 4 - seconds
    Days(0 To 6) As Boolean '0 - monday, 6 - sunday, true - selected, false - not selected
    FirstCheck As Date 'the first time to run or check on this event (in other words, the first time the sub has to look at it)
    WhatToDo As Integer '0 - Script, 1 - RCON
    ScriptName As String 'The script to run
    ComPara As String 'Commands or parameters (depending on what was selected)
    Name As String      'what is it CALLED!?
End Type

Public Type typSvMessageBuffer 'bufferes sa_message settings
    Red1 As Byte
    Green1 As Byte
    Blue1 As Byte
    Red2 As Byte
    Green2 As Byte
    Blue2 As Byte
    Effect As Integer
    HoldTime As Single
    Channel As Integer
    FxTime As Single
    FadeInTime As Single
    FadeOutTime As Single
    X As Single
    Y As Single
    Dynamic As Integer
End Type

Public Type typOldSwearWords
    BadWord As String
    Flags As Long
End Type

Public Type typSwearWords
    BadWord As String
    Replacement As String
    Flags As Long
    Unused1 As String
    unused2 As String
    unused3 As Integer
    unused4 As Integer
End Type

Public Type typDirList ' Directory Listing
    Name As String      'File/Dir name
    FullPath As String  'Full path (on server) to this file
    Type As Byte        '0 - file, 1 - dir
    DateTime As Date    'guess what this is?
    Size As String      'size of file
End Type

Public Type typLogFound
    LogFile As String
    LogLine As String
End Type

Public Type typAdminLast    'Record of the last x people
    UserID As Integer
End Type

Public Type typTeleportExit
    X As Integer
    Y As Integer
    Z As Integer
    Angle As Integer
    Name As String
End Type

Public Type typMessages     'Kinda like email
    MsgText As String
    MsgFor As String
    MsgFrom As String
    MsgTimeSent As Date
    MsgSubj As String
    MsgId As Integer
    Flags As Integer        '0 - request reciept, 1 - unread
End Type

Public Type SHFILEOPSTRUCT
   hwnd As Long
   wFunc As Long
   pFrom As String
   pTo As String
   fFlags As Long
   fAnyOperationsAborted As Long
   hNameMappings As Long
   lpszProgressTitle As String
End Type

Public Type typPointSystem
    DoubleKickVotesAt As Integer        'When points are => this, votes on kickvotes are doubled!
    DoubleMapVotesAt As Integer         'When points are => this, votes on mapvotes are doubled!
    KickVotesCost As Integer            'Starting a kickvote costs this much
    SpamKickCosts As Integer            'Getting kicked for spamming costs this much
    KickedCosts As Integer              'Getting kicked (kickvoted) costs this much
    JoiningAdds As Integer              'Joining the game adds this many
    
End Type

Public Type typLastLines
    Line As String
    Name As String
    TimeSent As String
    Team As Integer
End Type

Public Type typScriptData
    LastIf As Integer
    UserIP As String
    UserPort As Long
    UserName As String
    UserIsTCP As Boolean
    Index As Integer
    TimeStarted As Double
    VarNames() As String
    VarValues() As String
    NumVars As Integer
    NoAutoCreate As Boolean
    LogToFile As Boolean
    MaxTime As Byte
    IsRCON As Boolean
    ExitNow As Boolean
    StartedName As String
    LastRunningCheck As Date
    
End Type

Public Type typWhiteboard 'an object on the whiteboard

    ObjType  As Integer             ' Types:
                                    ' 1 = Line, data contains co-ords of 2 points
                                    ' 2 = Box, data contains corner co-ords
                                    ' 3 = circle/oval, data contains corner co-ords
                                    ' 4 = Rounded Rectangle, data contains corner co-ords
                                    ' 5 = Pencil drawing, data contains all the coords
                                    ' 6 = Image, data contains the image data and location
                                    ' 7 = Text, data contains text
                                    
                                    
    LineColour As Long
    ' co-ords
    Pos1X As Integer
    Pos1Y As Integer
    Pos2X As Integer
    Pos2Y As Integer
    LineWidth As Integer
    FillColour As Long
    
    ExtraData As String
    ShapeID As Integer
    
    Creator As String
    
    
End Type


Public Type typServerStart
    ' info for autorunning of server
    HLDSPath As String
    CommandLine As String
    HLDSDir As String
    
    
    AutoRestart As Boolean
    UseFeature As Boolean
    
    Unused1 As String
    unused2 As String
    unused3 As Integer
    unused4 As Integer


End Type


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'GLOBAL VARS
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

'Whiteboard
Public Shapes() As typWhiteboard
Public NumShapes As Integer

'server stopstart
Public ServerStart As typServerStart

'Messages
Public Messages() As typMessages
Public NumMessages As Integer

Public AdminBMP() As typAdminBMP
Public NumAdminBMP As Integer

'GetBans
Public GetBanList As Boolean

Public AlertedBan As Boolean

'Teleporters
Public Tele() As typTeleportExit
Public NumTele As Integer

'Log Searching
Public NumLogFound As Long
Public LogFound() As typLogFound
Public LogSearchString As String

'Files
Public DirList() As typDirList
Public NumDirs As Long
Public DirFullPath As String     'path

Public NextTime As Integer
Public PointData As typPointSystem

'Scheduler (events)
Public Events() As typEvent
Public NumEvents As Integer
Public EventTimer As Integer

'Fancy Message
Public SvMes As typSvMessageBuffer
Public MesLastY As Single
Public LastTalk As Byte

'Talking window
Public LastLines(1 To 20) As typLastLines

Public PVVersion As String

Public AdminICQId As Integer
Public AdminICQTime As Integer
Public AdminIcqNum As Integer

'debug crap
Public LastCalled As String
Public DebugMode As Boolean
Public DebugTime As Boolean

Public GetBansIndex As Integer
Public LastChats(1 To 20) As String
Public LastChatsName(1 To 20) As String
Public LastChatsTime(1 To 20) As String
Public LastChatsCol(1 To 20) As Long

'rcon
Public WaitingCommands() As String

'team names
Public TeamNames(1 To 4) As String


'crash detection
Public LastCrashCall As Integer
Public CrashTimer As Integer

'log sorting timer
Public LogSortTimer As Long
Public StopAllScripts As Boolean
Public ICQURL As String
Public SendingICQ As Boolean


'etc
Public TimeLeftSetting

'stuff involving new DLL
Public DLLEnabled As Boolean

'web
Public Web As typWebInfo

'auto unbanlast
Public UnBanLast As Boolean
Public BanScriptData As typScriptData

'real players
Public RealPlayers() As typRealPlayer
Public NumRealPlayers As Integer

'auto kicks/bans
Public NumKickBans As Integer
Public KickBans(1 To 200) As typKickBans

'enabling
Public ProgramON As Boolean

'Swearing
Public Swears() As typSwearWords
Public NumSwears As Integer

'Rcon watcher (startlog)
Public RconWatchersIP As New Collection
Public RconWatchersPort As New Collection

Public TempStop As Boolean

'speech
Public Speech() As typSpeech
Public NumSpeech As Integer
Public LastSpeech(1 To 20) As typAdminLast
Public AdminSpeechBuffer As Integer

'kills
Public NumKills As Integer
Public KillList() As typKills

'Scripts
Public Commands() As typScripts
Public NumCommands As Integer
Public Server As serv
Public LastCommand As String
Public Vars As Varss
Public General As typGeneral
Public LogPath As String
Public CurrLastLog As String
Public NewLastLog As String
Public TimerVars() As String
Public NumTimerVars As Integer

Public Alpha(1 To 26)
Public ChallengeNum As String
Public DoAfterChallenge As Boolean


'players list
Public Players(1 To 400) As Playersss
Public NumPlayers As Integer

'time calculation
Public LogWatchTimer As Integer
Public AnnTime As Boolean
Public TimeRemainAskCount As Integer

'clans list
Public Clans(1 To 20) As typClan
Public NumClans As Integer

' running
Public FindRunningScripts As Boolean
Public FindRunningScriptsTime As Date
Public RunningScripts As String

'map vote
Public Votes(1 To 200) As Votess
Public NumVotes As Integer
Public AvailMaps As New Collection
Public MapVoteTimer As Integer
Public LastMaps As New Collection

Public LastUpdateCheck As Double

'Map Storage Info
Public MapArray(0 To 64, 0 To 64) As Integer ' Stores Z coord at this location, used for keeping map data

'Map Data Format:
' -4096   to   4096 -> more than one team / old format
'  4097   to  12288 -> blue team   (norm: -8192)
' 12289   to  20480 -> red team    (norm: -16384)
' -4097   to -12288 -> yellow team (norm: +8192)
'-12289   to -20480 -> green team  (norm: +16384)


'Kickvotes
Public KickVoteTimer As Integer
Public KickVoteUser As Integer
Public KickVoteStarterName As String
Public KickVoteStarterUIN As String
Public LastKickVotes(1 To 20) As String         'Stores uniqueIDs of the last 20 people kick-voted

'ChooseVote
Public ChooseVoteTimer As Integer
Public ChooseVote() As String
Public ChooseVoteQuestion As String
Public ChooseVoteTime As Integer

'connected users
Public Users(1 To 200) As typUsers
Public NumUsers As Integer

'constants
Public TimerVR As Integer
Public TimerVar2 As Integer
Public NumParams As Integer
Public DataFile As String
Public DataFile2 As String
Public DataFile3 As String
Public DataFileNew As String
Public DataFileOld As String
Public DataFileOlder As String
Public Data(1 To 20) As String

'more script stuff
Public NumScriptParams As Integer
Public Params(1 To 200) As String
Public UserVars() As typUserVars
Public NumUserVars As Integer

'rcon stuff
Public LastRCON As String
Public ScriptParams(1 To 200) As String
Public Settings As OtherSettings
Public TimeElapsed As Integer
Public LogBuffer As String
Public Indent As String
Public LastIf As Integer
Public Reload As Boolean

'map
Public MapCounter As Integer

'users
Public UserEditNum As Integer
Public UserRCONBuffer() As typUserRCON
Public RconMonitors() As typRconMonitor
Public NumRconMonitors As Integer
Public LastUser As Integer

'connected users
Public ConnectUsers(1 To 400) As typConnectUsers
Public NumConnectUsers As Integer
Public TCPCreated(0 To 400) As Boolean
Public RecData(0 To 30) As String

'map processing
Public MapProcess() As typMapRecord
Public NumMapProcess As Integer

Public CurrBans() As typBans
Public NumCurrBans As Integer

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'API DECLARATIONS
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const TH32CS_SNAPPROCESS As Long = 2&
Public Const MAX_PATH As Integer = 260
Public Const SW_SHOW As Integer = 5

Public Const PROCESS_TERMINATE = &H1

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SW_SHOWNORMAL = 1
Public Const SE_ERR_FNF = 2
Public Const SE_ERR_NOASSOC = 31

Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function SendMessageWindow Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (lpExecInfo As SHELLEXECUTEINFO) As Long

   
'Public Declare Function ReportEnts Lib "assist2.dll" Alias "reportents" (BSPPath As String, BankData As String) As Long
Public Declare Function InitEnts Lib "assist2.dll" Alias "init_ents" (BSPPath As String) As Long
Public Declare Function EntData Lib "assist2.dll" Alias "entdata" (BankData As String) As Boolean
Public Declare Function ImportEnts Lib "assist2.dll" Alias "writeents" (BSPPath As String, BankData As String) As Long

Public Declare Function DllCanUnloadNow Lib "assist2.dll" () As Integer
Public Const BankSize = 512000
Public Const WM_CLOSE = &H10

Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long

'Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
'Declare Function GetCurrentProcess Lib "kernel32" () As Long
'Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
'Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'Global Const PROCESS_TERMINATE = &H1&
'Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Declare Function GetVolumeSerialNumber Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'Constants
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const SA_CHECK = "checktime"
Public Const DLLFile = "assist.dll"

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
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONDOWN = &H204


Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    End Type
    Public Const VER_PLATFORM_WIN32s = 0
    Public Const VER_PLATFORM_WIN32_WINDOWS = 1
    Public Const VER_PLATFORM_WIN32_NT = 2


Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long


Public Type PROCESS_BASIC_INFORMATION
    ExitStatus As Long
    PebBaseAddress As Long
    AffinityMask As Long
    BasePriority As Long
    UniqueProcessId As Long
    InheritedFromUniqueProcessId As Long 'ParentProcessID'
    End Type


Public Declare Function NtQueryInformationProcess _
    Lib "ntdll" (ByVal ProcessHandle As Long, _
    ByVal ProcessInformationClass As Long, _
    ByRef ProcessInformation As PROCESS_BASIC_INFORMATION, _
    ByVal lProcessInformationLength As Long, _
    ByRef lReturnLength As Long) As Long


Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
    Public Const PROCESS_VM_READ = 16

Public HLDSProcID As Long

Public SupposedToBeRunning As Boolean
Public ManualStart As Boolean

'process crap

Global Const TH32CS_SNAPHEAPLIST = &H1
'Global Const TH32CS_SNAPPROCESS = &H2
Global Const TH32CS_SNAPTHREAD = &H4
Global Const TH32CS_SNAPMODULE = &H8
Global Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Global Const TH32CS_INHERIT = &H80000000

Global Const NORMAL_PRIORITY_CLASS = 8
Global Const IDLE_PRIORITY_CLASS = 4
Global Const HIGH_PRIORITY_CLASS = 13
Global Const REALTIME_PRIORITY_CLASS = 24



Type THREAD32ENTRY

  dwSize As Long
  cntUsage As Long
  th32ThreadID As Long
  th32OwnerProcessID As Long
  tpBasePri As Long
  tpDeltaPri As Long
  dwFlags As Long
  
End Type

Type MODULEENTRY32

  dwSize As Long
  th32ModuleID As Long
  th32ProcessID As Long
  GlblcntUsage As Long
  ProccntUsage As Long
  modBaseAddr As Long
  modBaseSize As Long
  hModule As Long
  szModule As String * 256
  szExePath As String * 260
  
End Type


Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal dwIdProc As Long) As Long
Declare Function Process32First Lib "kernel32" (ByVal hndl As Long, ByRef pstru As PROCESSENTRY32) As Boolean
Declare Function Process32Next Lib "kernel32" (ByVal hndl As Long, ByRef pstru As PROCESSENTRY32) As Boolean
Declare Function Thread32First Lib "kernel32" (ByVal hndl As Long, ByRef pstru As THREAD32ENTRY) As Boolean
Declare Function Thread32Next Lib "kernel32" (ByVal hndl As Long, ByRef pstru As THREAD32ENTRY) As Boolean
Declare Function Module32First Lib "kernel32" (ByVal hndl As Long, ByRef pstru As MODULEENTRY32) As Boolean
Declare Function Module32Next Lib "kernel32" (ByVal hndl As Long, ByRef pstru As MODULEENTRY32) As Boolean
Declare Function GetLastError Lib "kernel32" () As Long
Declare Sub Sleep Lib "kernel32" (ByVal msec As Long)



Global vgProcInd As Integer
Global vgTotalThrd As Integer

Type MODULE32LST

  Index As Integer
  moduleFullPath As String
  
End Type

Type THRD32LST

  Index  As Integer
  thrdId As Long
  BasePriority As Long
  countUsage As Long

End Type

Type PROC32LIST

  count  As Integer
  Index  As Integer
  procID As Long
  threadCount As Long
  countUsage As Long
  defaultHeapID As Long
  moduleID As Long
  parentProcessID As Long
  pcPriClassBase As Long
  Flags As Long
  procName As String
  exeFile As String
   
  thrdInfo(256) As THRD32LST
  
  countModules As Integer
  moduleInfo(256) As MODULE32LST

End Type

Global vgProc32(1024) As PROC32LIST

Global LastTimer2 As Date
Global LastTimer2What As String

