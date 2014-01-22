Attribute VB_Name = "a_mains"
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
'      FILE: udp1.bas
'      PURPOSE: Main code file for the Client of SA.
'
'
'
' ===========================================================================
' ---------------------------------------------------------------------------

Public UnloadTime As Boolean

Public ImGone As Boolean
Public LastMsg As String

Public Type typFormList
    
    'does it snap?
    Snaps As Boolean
    
    'dimensions
    MaxWidth As Integer
    MaxHeight As Integer
    MinWidth As Integer
    MinHeight As Integer
    
    'window ID
    hwnd As Long
    FormObj As Object
    FormName As String
    
End Type

Public FormList() As typFormList

Public Type typButtons
    ButtonName As String
    ButtonText As String
    OptionOn As String
    OptionOff As String
    Type As Integer
End Type

Public Type typMapRecord
    MapName As String
    TimesPlayed As Long
    LastTimePlayed As Date
    Unused1 As String
    unused2 As String
    unused3 As String
    Perc As Double
    
End Type

Public Type typGameData
    GameForm As Object
    Index As Integer
    Opponent As String
End Type

Public Type typMenuScripts
    AskForQuestion As String            'ask this when script starts.
    MenuName As String
    ScriptName As String
End Type

Public Type typConnectUsers
    Name As String
    IP As String
    Version As String
    AwayMode As Integer
    AwayMsg As String
    IdleTime As Long
    
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
    Changed As Boolean
End Type

Public Type typToolBar
    Description As String
    IconID As Integer
    Type As Integer
    Tag As String
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

Public Type typPlayerPos
    X As Integer
    Y As Integer
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
    pos As typPlayerPos
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
    PassWord As String
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
    Replacement As String
    Flags As Long
    Unused1 As String
    unused2 As String
    unused3 As Integer
    unused4 As Integer
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
    CustomFlag1 As String
    CustomFlag2 As String
    CustomFlag3 As String
    CustomFlag4 As String
    MaxFileSend As Integer
    NewestClient As String
    SendToDisco As Integer
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
    X As Integer
    Y As Integer
    Z As Integer
    Name As String
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
    fillColour As Long
    
    ExtraData As String
    ShapeID As Integer
    Creator As String
    
End Type


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'GLOBAL VARS
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

'away mode
Public MyAwayMode As Integer
Public MyAwayMsg As String

Public SetAway10Min As Boolean
Public SetNA20Min As Boolean
Public AutoAwayReturn As Boolean

Public SecondsAway As Long
Public LastMouseX As Integer
Public LastMouseY As Integer

Public AutoSet As Boolean

Public PrevConnection As Boolean


'Whiteboard
Public Shapes() As typWhiteboard
Public NumShapes As Integer


Public AdminBMP() As typAdminBMP
Public NumAdminBMP As Integer

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

Public ButtonShowMode As Integer

Public SendSize As Long
Public FindReal As String

Public Swears() As typSwearWords
Public NumSwears As Integer

Public DllEnabled As Boolean
Public GameMode As Integer

Public ServVersion As String

Public LastKnownState As Integer

Public NewestClient As String

'server stopstart
Public ServerStart As typServerStart
'flags

Public CustomFlag1 As String
Public CustomFlag2 As String
Public CustomFlag3 As String
Public CustomFlag4 As String

'Menu Scripts
Public NumMenuScripts As Integer
Public MenuScripts() As typMenuScripts


'Files and File Transfer Vars
Public DirList0() As typDirList
Public DirList1() As typDirList
Public NumDirs(0 To 1) As Long
Public DirFullPath(0 To 1) As String     'path to the current window
Public FileBuffer As String
Public FileMode As Integer               '0 - unset, 1 - edit this file, 2 - download this file
Public FilePath As String                'Path to download to for mode 2 above
Public FileLocalPath As String           'Where to save this file LOCALLY
Public DownloadingNew As Boolean           'Where to save this file LOCALLY
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

Public ScriptButtonName As String

'team names
Public TeamNames(1 To 4) As String

'chat freeze
Public ChatFrozen As Boolean
Public FreezeText() As String

'Scheduler :)
Public Events() As typEvent
Public NumEvents As Integer

'time/update
Public SecondsLeft As Long
Public MapName As String
Public PlayersOn As String

Public ShowPlayers As Boolean
Public ShowUsers As Boolean
Public ShowMap As Boolean
Public ShowChat As Boolean
Public ShowWhiteBoard As Boolean

Public ServerBans() As typBans
Public NumCurrBans As Integer

Public EmailCheckCounter As Integer

Public HLEXEPath As String
Public HLExtraArgs As String
Public HLQuitSA As Integer
Public HLSetAway As Integer
Public HLIP As String
Public HLPort As String
Public HLGame As String

Public EncryptedMode As Boolean


'web
Public Web As typWebInfo

'real players
Public RealPlayers() As typRealPlayer
Public NumRealPlayers As Integer

Public Server As serv
Public Presets(1 To 200) As typPreset
Public NumPresets As Integer

Public HiddenNow As Boolean


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

Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Public Type POINTAPI
        X As Long
        Y As Long
End Type


Public Const ERR_TEXT_CONTROL_EXPECTED = 20000
Public Const ERR_TEXT_CONTROL_EXPECTED_MSG = "Expected: Text Box or other edit control"

Public Const EM_GETLINECOUNT = &HBA

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
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As WINDOWPOS) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = (-4)

Type WINDOWPOS
    hwnd As Long
    hWndInsertAfter As Long
    X As Long
    Y As Long
    cx As Long
    cy As Long
    Flags As Long
End Type
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WINDOWPOSCHANGED = &H47

Private Declare Function GetOpenFileName Lib "comdlg32.dll" _
        Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) _
        As Long
        
Private Declare Function GetSaveFileName Lib "comdlg32.dll" _
        Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) _
        As Long



Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHOWHELP = &H10




Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Public Const SND_SYNC = &H0        ' Play synchronously (default).
Public Const SND_ASYNC = &H1      ' Play asynchronously (see note below).
Public Const SND_NODEFAULT = &H2   ' Do not use default sound.
Public Const SND_MEMORY = &H4      ' lpszSoundName points to a  memory file.
Public Const SND_LOOP = &H8        ' Loop the sound until next sndPlaySound.
Public Const SND_NOSTOP = &H10     ' Do not stop any currently
                                   ' playing sound.
Public bytSound() As Byte ' Always store binary data in byte arrays!

Public Const FO_MOVE As Long = &H1
Public Const FO_COPY As Long = &H2
Public Const FO_DELETE As Long = &H3
Public Const FO_RENAME As Long = &H4

Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

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

Public Commands() As typScripts
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

Public CurrentToolBar() As typToolBar
Public DefaultToolBar() As typToolBar

'map processing
Public MapProcess() As typMapRecord
Public NumMapProcess As Integer

'games
Public GameData() As typGameData
Public NumGames As Integer


'Map
Public MapArray(0 To 64, 0 To 64) As Integer ' Stores Z coord at this location


Public Sub PlayWaveRes(vntResourceID As Variant, Optional vntFlags)
'-----------------------------------------------------------------
' WARNING:  If you want to play sound files asynchronously in
'           Win32, then you MUST change bytSound() from a local
'           variable to a module-level or static variable. Doing
'           this prevents your array from being destroyed before
'           sndPlaySound is complete. If you fail to do this, you
'           will pass an invalid memory pointer, which will cause
'           a GPF in the Multimedia Control Interface (MCI).
'-----------------------------------------------------------------

bytSound = LoadResData(vntResourceID, "WAVE")

If IsMissing(vntFlags) Then
   vntFlags = SND_NODEFAULT Or SND_ASYNC Or SND_MEMORY
End If

If (vntFlags And SND_MEMORY) = 0 Then
   vntFlags = vntFlags Or SND_MEMORY
End If

sndPlaySound bytSound(0), vntFlags
End Sub

Function MessBox(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String, Optional ShowMode As Boolean, Optional TimeToShow As Integer) As Long

Dim MessageBox As New frmMessageBox

MessageBox.Prompt = Prompt
MessageBox.Buttons = Buttons
MessageBox.Title = Title

MessageBox.Display
MessageBox.ReturnValue = -1
MessageBox.ShowMode = ShowMode

If ShowMode = False Then
    ttm = Timer
    Do
        DoEvents
        ttr = Timer - ttm
        If TimeToShow > 0 Then ttd = ttr
        
    Loop Until MessageBox.ReturnValue <> -1 Or ttd > TimeToShow

    MessBox = MessageBox.ReturnValue
    Unload MessageBox

ElseIf TimeToShow > 0 Then
    
    MessageBox.TimeToShow = TimeToShow
    
End If

End Function

Function SelIcon(CurrIc As Long) As Long

Dim MessageBox As New frmIconSelect

MessageBox.CurrIcon = CurrIc
MessageBox.DoneDone = False
MessageBox.Show

    Do
        DoEvents
    Loop Until MessageBox.DoneDone

    n = MessageBox.Selected
    
    Unload MessageBox
    
    SelIcon = n


End Function


Function InBox(Prompt As String, Optional Title As String, Optional Default As String) As String

Dim MessageBox As New frmInputBox

MessageBox.Prompt = Prompt
MessageBox.Title = Title


MessageBox.ReturnValue = ""
MessageBox.Default = Default
MessageBox.Display
Do
    DoEvents
Loop Until MessageBox.Finito <> 0

InBox = MessageBox.ReturnValue
Unload MessageBox

End Function


Function CalenBox(InitialDate As Date, Caption As String) As Date

Dim DateBox As New frmCalendar

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
SendPacket "M1", a$

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

'Debug.Print Timer - strt, Len(String2), Len(EndString)

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


frmMain.RichTextBox1.BackColor = RGB(RichColors(9).r, RichColors(9).g, RichColors(9).b)
Form6.ListView1.BackColor = RGB(RichColors(9).r, RichColors(9).g, RichColors(9).b)
Form6.ListView1.ForeColor = RGB(RichColors(1).r, RichColors(1).g, RichColors(1).b)


End Sub

Function BracketCount(CharNum, String1 As String) As Integer
If DebugMode Then LastCalled = "BracketCount"

'Counts the bracket level of this character

Dim Brck As Integer

e = 0
fff = 0
Brck = 0

Do
    e2 = InStrQuote(e + 1, String1, "(")
    e4 = InStrQuote(e + 1, String1, ")")
    
    flg = 0
    If e2 > 0 Then flg = 1
    e = e2
    If e = 0 Then e = 100000000
    If e4 < e And e4 > 0 Then e = e4: flg = 2
    
    If e <= CharNum Then
        If flg = 1 Then Brck = Brck + 1
        If flg = 2 Then Brck = Brck - 1
    End If
    
    
    If StopAllScripts = True Then Exit Function
    DoEvents

Loop Until e = 100000000 Or e >= CharNum

BracketCount = Brck

End Function

Sub AddForm(Snaps As Boolean, MinW As Integer, MinH As Integer, MaxW As Integer, MaxH As Integer, ByRef FormObj As Object)

'Exit Sub
If App.EXEName = "udp" Then Exit Sub


n = UBound(FormList)
b = 0
For i = 1 To n
    If FormObj.Name = FormList(i).FormName Then b = i: Exit For
Next i

If b = 0 Then
    n = n + 1
    ReDim Preserve FormList(0 To n)
    b = n
End If

With FormList(b)
    .FormName = FormObj.Name
    Set .FormObj = FormObj
    .hwnd = FormObj.hwnd
    .MaxHeight = MaxH
    .MaxWidth = MaxW
    .MinHeight = MinH
    .MinWidth = MinW
    .Snaps = Snaps
End With

'now, call the function

FormObj.OldWindowProc = SetWindowLong(FormObj.hwnd, GWL_WNDPROC, AddressOf NewWindowProc)


End Sub

Sub Main()



Load MDIForm1
MDIForm1.Visible = False

DllEnabled = True

ReDim FreezeText(0 To 0)
ReDim FormList(0 To 0)
ReDim Commands(0 To 0)

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
LoadDefaultToolbar
ReDim CurrentToolBar(0 To 0)
'CopyDefaultToCurrent

DataFile = App.Path + "\client.dat"
DataFile2 = App.Path + "\recentip.dat"
DataFile3 = App.Path + "\lastconn.dat"

'ReDim Commands(1 To 200)
LoadCommands

ApplyToToolbar


'MDIForm1.Toolbar1.RestoreToolbar "Software", "Server Assistant Client", "ToolBar"

MDIForm1.Caption = "Server Assistant Client - Copyright 2001 CyberWyre"
'MDIForm1.Show

EditFileTemp = App.Path + "\temp1.txt"

MDIForm1.StatusBar1.Panels(1).Text = "Client: " + Ts(App.Major) + "." + Ts(App.Minor) + "." + Ts(App.Revision)

MDIForm1.Toolbar1.Enabled = False
a4 = Val(GetSetting("Server Assistant", "Settings", "FirstTime", "0"))

If a4 = 0 Then
    a5 = MessBox("Welcome to Server Assistant Client." + vbCrLf _
+ "This is the first time you are running this program" + vbCrLf + "with the new toolbar in place." + vbCrLf + vbCrLf _
+ "Please note that by default, there are a large number" + vbCrLf + "of buttons on the toolbar, not all of them needed." + vbCrLf + vbCrLf _
+ "It it suggested that you immediatly open the Customize" + vbCrLf + "Toolbar window and remove the buttons you will not" + vbCrLf + "be using." + vbCrLf _
+ vbCrLf + "This feature is available from the Settings menu." + vbCrLf _
+ "Would you like to open this window now?" + vbCrLf, vbQuestion + vbYesNo, "Welcome to Server Assistant Client")

    If a5 = vbYes Then frmCustomize.Show
End If


    MDIForm1.Height = GetSetting("Server Assistant Client", "Window", "winh", 800 * Screen.TwipsPerPixelX)
    MDIForm1.Top = GetSetting("Server Assistant Client", "Window", "wint", 50 * Screen.TwipsPerPixelX)
    MDIForm1.Left = GetSetting("Server Assistant Client", "Window", "winl", 50 * Screen.TwipsPerPixelX)
    MDIForm1.Width = GetSetting("Server Assistant Client", "Window", "winw", 800 * Screen.TwipsPerPixelX)
'End If

MDIForm1.WindowState = Val(GetSetting("Server Assistant Client", "Window", "winmd", 2))
MDIForm1.Timer2.Enabled = False

MDIForm1.Show

'Load Form1
Form1.Show
'Load frmMain
frmMain.Show
'frmConnect.Show

a1 = Val(GetSetting("Server Assistant", "Settings", "NamesInColor", "-1"))
a2 = Val(GetSetting("Server Assistant", "Settings", "TimeStamp", "0"))
a3 = Val(GetSetting("Server Assistant", "Settings", "ShowMessages", "0"))
a4 = Val(GetSetting("Server Assistant", "Settings", "PopUpAdmin", "-1"))
a5 = Val(GetSetting("Server Assistant", "Settings", "AutoReconnect", "-1"))
a6 = Val(GetSetting("Server Assistant", "Settings", "SnapWindows", "-1"))
a7 = Val(GetSetting("Server Assistant", "Settings", "EnableBing", "-1"))
SaveSetting "Server Assistant", "Settings", "FirstTime", "1"

If a1 <> MDIForm1.mnuSettingsIn(1).Checked Then
    MDIForm1.mnuSettingsIn_Click 1
End If

If a2 <> MDIForm1.mnuSettingsIn(2).Checked Then
    MDIForm1.mnuSettingsIn_Click 2
End If

If a3 <> MDIForm1.mnuSettingsIn(3).Checked Then
    MDIForm1.mnuSettingsIn_Click 3
End If

If a4 <> MDIForm1.mnuSettingsIn(4).Checked Then
    MDIForm1.mnuSettingsIn_Click 4
End If

If a5 <> MDIForm1.mnuSettingsIn(5).Checked Then
    MDIForm1.mnuSettingsIn_Click 5
End If
If a6 <> MDIForm1.mnuSettingsIn(6).Checked Then
    MDIForm1.mnuSettingsIn_Click 6
End If
If a7 <> MDIForm1.mnuSettingsIn(7).Checked Then
    MDIForm1.mnuSettingsIn_Click 7
End If
' away mode stuff

min10timer = Val(GetSetting("Server Assistant", "Settings", "AwayMode10Min", "-1"))
min20timer = Val(GetSetting("Server Assistant", "Settings", "AwayMode20Min", "-1"))
autoreturn = Val(GetSetting("Server Assistant", "Settings", "AutoReturn", "-1"))

If min10timer = -1 Then SetAway10Min = True
If min10timer = 1 Then SetAway10Min = True

If min20timer = -1 Then SetNA20Min = True
If min20timer = 1 Then SetNA20Min = True

If autoreturn = -1 Then AutoAwayReturn = True
If autoreturn = 1 Then AutoAwayReturn = True


End Sub

Sub Swap(a As Variant, b As Variant)
Dim c As Variant

c = a
a = b
b = c




End Sub


Sub SendExe()

'prepare to send EXE

'a$ = inbox("New EXE path?", "Path?", App.Path + "\server.exe")
b$ = CompileEXE(a$)
SendPacket "EX", b$

End Sub

Function Ts(a) As String
    Ts = Trim(str(a))
End Function

Function CheckForFile(a$) As Boolean
    b$ = Dir(a$)
    If b$ = "" Then CheckForFile = False
    If b$ <> "" Then CheckForFile = True
    
End Function

Sub test()


'Open "tes.txt" For Append As #1
'F 'or I = 1 To 50000
'p 'rint #1, Ts(I)
'n 'ext I
'c 'lose #1

'mdiform1.mnuFunctionsIn(14).


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
b$ = ""

For i = 1 To NumRealPlayers
    a$ = a$ + Chr(251)
    a$ = a$ + Ts(RealPlayers(i).Flags) + Chr(250)
    a$ = a$ + RealPlayers(i).RealName + Chr(250)
    a$ = a$ + RealPlayers(i).UniqueID + Chr(250)
    a$ = a$ + Ts(CDbl(RealPlayers(i).LastTime)) + Chr(250)
    a$ = a$ + RealPlayers(i).LastName + Chr(250)
    a$ = a$ + RealPlayers(i).Points + Chr(250)
    a$ = a$ + RealPlayers(i).TimesSeen + Chr(250)
    a$ = a$ + Chr(251)
    
    If Len(a$) > 1000 Then b$ = b$ & a$: a$ = ""
    
Next i

b$ = b$ & a$

'all set, send it
SendPacket "RR", b$

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
SendPacket "RA", a$

End Sub

Sub UpdatePlayerList()
'Form6.Show
bbc = -1

If ShowPlayers = False Then Exit Sub
Load Form6 'Form6.Show

'Form6.Visible = ShowPlayers

For i = 1 To Form6.ListView1.ListItems.Count
    If Form6.ListView1.ListItems.Item(i).Selected = True Then bbc = Val(Form6.ListView1.ListItems.Item(i).SubItems(2))
Next i

With Form6.ListView1.ListItems

For i = 1 To .Count
    .Item(i).Tag = "0"
Next i

For i = 1 To NumPlayers
    bc = Len(Ts(Players(i).UserID))
    If bc > maxnum Then maxnum = bc
Next i
maxnum = maxnum + 1

For i = 1 To NumPlayers

    img = 0
    'find
    j = 0
    For k = 1 To .Count
        If Val(.Item(k).SubItems(2)) = Players(i).UserID Then j = k: Exit For
    Next k
    
    If j = 0 Then
        'add
        j = .Count + 1
        Randomize
        jk = Int(Rnd * 30000) + 1
        .Add j, "A" + Players(i).Name + Ts(jk)
        
        For k = 1 To .Count
            If .Item(k).Key = "A" + Players(i).Name + Ts(jk) Then jj = k: Exit For
        Next k
        Form6.ListView1.Sorted = False
        
        j = jj
        
    End If
    
    If .Item(j).Text <> Players(i).Name Then .Item(j).Text = Players(i).Name
    If .Item(j).SubItems(1) <> Players(i).RealName Then .Item(j).SubItems(1) = Players(i).RealName
    
    us$ = Ts(Players(i).UserID)
    If Len(us$) < maxnum Then us$ = Space(maxnum - Len(us$)) + us$
    
    If .Item(j).SubItems(2) <> us$ Then .Item(j).SubItems(2) = us$
    If .Item(j).SubItems(3) <> Players(i).UniqueID Then .Item(j).SubItems(3) = Players(i).UniqueID
    
    tt1 = 0
    img = Players(i).Team + 2
    If Players(i).Team = 1 And GameMode <> 2 Then t$ = "Blue": tt1 = 1
    If Players(i).Team = 2 And GameMode <> 2 Then t$ = "Red": tt1 = 2
    If Players(i).Team = 1 And GameMode = 2 Then t$ = "Terrorists"
    If Players(i).Team = 2 And GameMode = 2 Then t$ = "CT"
    If Players(i).Team = 3 Then t$ = "Yellow": tt1 = 3
    If Players(i).Team = 4 Then t$ = "Green": tt1 = 4
    If Players(i).Team = 6 Then t$ = "Spectator": img = 7
    If Players(i).Team = 0 Then t$ = " None": img = 1
   
    If tt1 > 0 Then
        
        If TeamNames(tt1) <> "" Then t$ = TeamNames(tt1)
    
    End If
    
   
   
    cc = Players(i).Team + 1
    'If cc = 1 Then cc = 0
    cc = RGB(RichColors(cc).r, RichColors(cc).g, RichColors(cc).b)
   
    If .Item(j).SubItems(4) <> t$ Then .Item(j).SubItems(4) = t$
    If .Item(j).ListSubItems(4).ForeColor <> cc Then .Item(j).ListSubItems(4).ForeColor = cc
    
    If GameMode <> 2 Then
        If Players(i).Class = 1 Then r$ = "Scout"
        If Players(i).Class = 2 Then r$ = "Sniper"
        If Players(i).Class = 3 Then r$ = "Soldier"
        If Players(i).Class = 4 Then r$ = "Demoman"
        If Players(i).Class = 5 Then r$ = "Medic"
        If Players(i).Class = 6 Then r$ = "HWGuy"
        If Players(i).Class = 7 Then r$ = "Pyro"
        If Players(i).Class = 8 Then r$ = "Spy"
        If Players(i).Class = 9 Then r$ = "Engineer"
        If Players(i).Class = 0 Then r$ = "N/A": img = 1
        If Players(i).Class = -1 Then r$ = "Civilian"
        If Players(i).Class = -2 Then
            r$ = "Undecided": img = 1
            If Players(i).Team <> 0 Then img = 8
            If Players(i).Team = 6 Then img = 7
            
        End If
    Else
        'r$ = "N/A"
    End If
    
    If .Item(j).SubItems(5) <> r$ Then .Item(j).SubItems(5) = r$
    If .Item(j).SubItems(6) <> Players(i).IP Then .Item(j).SubItems(6) = Players(i).IP
    
    If Players(i).Status = True Then r$ = "Connected": img = 2
    If Players(i).Status = False Then r$ = "Playing"
    
    If .Item(j).SubItems(7) <> r$ Then .Item(j).SubItems(7) = r$
    
    'Calc time playing
    
    sec$ = Ts(Second(Players(i).JoinTime))
    mn$ = Ts(Minute(Players(i).JoinTime))
    hr$ = Ts(Hour(Players(i).JoinTime))
    
    If Len(hr$) = 1 Then hr$ = "0" + hr$
    If Len(sec$) = 1 Then sec$ = "0" + sec$
    If Len(mn$) = 1 Then mn$ = "0" + mn$
    hr$ = hr$ + ":" + mn$ + ":" + sec$
    
    If .Item(j).SubItems(8) <> hr$ Then .Item(j).SubItems(8) = hr$
    .Item(j).SmallIcon = img
    .Item(j).Tag = Ts(i)
    
Next i

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
For i = 1 To .Count
    If .Item(i).Tag = "0" Or .Item(i).Text = "" Then
        .Remove i: GoTo aggg
    End If
Next i

End With

End Sub

Sub UpdateUsersList()


MDIForm1.StatusBar1.Panels(5).Text = "Users: " + Ts(NumConnectUsers)

If ShowUsers = False Then Exit Sub


For i = 1 To frmConnectUsers.ListView1.ListItems.Count
    'If frmConnectUsers.ListView1.ListItems.Item(I).Selected = True Then bbc = Val(Form6.ListView1.ListItems.Item(I).SubItems(2))
Next i

With frmConnectUsers.ListView1.ListItems

For i = 1 To .Count
    .Item(i).Tag = "0"
Next i


For i = 1 To NumConnectUsers
    
    picc = 0
    For k = 1 To frmConnectUsers.ImageList1.ListImages.Count
        If frmConnectUsers.ImageList1.ListImages.Item(k).Tag = ConnectUsers(i).Name Then
            picc = k
        End If
    Next k
    
    If picc = 0 Then
        ' load
        k = 0
        For Ii = 1 To NumAdminBMP
            If AdminBMP(Ii).AdminName = ConnectUsers(i).Name Then k = Ii: Exit For
        Next Ii
        
        If k > 0 Then
            If CheckForFile(App.Path + "\apics\" + AdminBMP(k).BMPFile) Then
                
                On Error Resume Next
                
                frmConnectUsers.Picture1.Picture = LoadPicture(App.Path + "\apics\" + AdminBMP(k).BMPFile)
                
                ' assign picture
                n = frmConnectUsers.ImageList1.ListImages.Count
                frmConnectUsers.ImageList1.ListImages.Add n + 1, , frmConnectUsers.Picture1.Picture
                frmConnectUsers.ImageList1.ListImages.Item(n + 1).Tag = ConnectUsers(i).Name
                picc = n + 1
            End If
        End If
        
    End If
    
    
    
    j = 0
    For k = 1 To .Count
        If .Item(k).SubItems(1) = ConnectUsers(i).IP Then j = k: Exit For
    Next k
    
    If j = 0 Then
        'add
        j = .Count + 1
        Randomize
        jk = Int(Rnd * 30000) + 1
        .Add j, "A" + ConnectUsers(i).IP + Ts(jk)
        
        For k = 1 To .Count
            If .Item(k).Key = "A" + ConnectUsers(i).IP + Ts(jk) Then jj = k: Exit For
        Next k
        frmConnectUsers.ListView1.Sorted = False
        
        j = jj
        
    End If
    
    
    
    .Item(j).Text = ConnectUsers(i).Name
    
    .Item(j).SubItems(1) = ConnectUsers(i).IP
    .Item(j).SubItems(2) = ConnectUsers(i).Version
    .Item(j).SubItems(3) = GetAwayName(ConnectUsers(i).AwayMode)
    .Item(j).Icon = picc
    .Item(j).SmallIcon = picc
    .Item(j).Tag = Ts(i)
'
'
'
'    b$ = ConnectUsers(I).Name
'    If Len(ConnectUsers(I).Name) < 20 Then
'        b$ = b$ + Space(24 - Len(ConnectUsers(I).Name)) + ConnectUsers(I).IP
'    Else
'        b$ = b$ + Space(5) + ConnectUsers(I).IP
'    End If
'
'    If Len(ConnectUsers(I).IP) < 20 Then
'        b$ = b$ + Space(30 - Len(ConnectUsers(I).IP)) + ConnectUsers(I).Version
'    Else
'        b$ = b$ + Space(5) + ConnectUsers(I).Version
'    End If
'
'    If Len(b$) > 50 Then
'        b$ = b$ + Space(5) + GetAwayName(ConnectUsers(I).AwayMode)
'    Else
'        b$ = b$ + Space(55 - Len(b$)) + GetAwayName(ConnectUsers(I).AwayMode)
'    End If
'
'    frmConnectUsers.List1.AddItem b$
'    e = frmConnectUsers.List1.NewIndex
'    frmConnectUsers.List1.ItemData(e) = I
Next i

frmConnectUsers.Caption = "Connected Users - " + Ts(NumConnectUsers)
frmConnectUsers.ListView1.Sorted = True

aggg:
For i = 1 To .Count
    If .Item(i).Tag = "0" Or .Item(i).Text = "" Then
        .Remove i: GoTo aggg
    End If
Next i

End With
End Sub

Function GetAwayName(Mde As Integer) As String

If Mde = 0 Then GetAwayName = "Available"
If Mde = 1 Then GetAwayName = "Away"
If Mde = 2 Then GetAwayName = "N/A"
If Mde = 3 Then GetAwayName = "Sleeping"
If Mde = 4 Then GetAwayName = "In Game"
If Mde = 5 Then GetAwayName = "Eating"

End Function

Function ReplaceString(ByVal Txt As String, ByVal from_str As String, ByVal to_str As String)
'Dim new_txt As String
'Dim pos As Integer
'
'    Do While Len(Txt) > 0
'        pos = InStr(Txt, from_str)
'        If pos = 0 Then
'            ' No more occurrences.
'            new_txt = new_txt & Txt
'            Txt = ""
'        Else
'            ' Found it.
'            new_txt = new_txt & Left$(Txt, pos - 1) & to_str
'            Txt = Mid$(Txt, pos + Len(from_str))
'        End If
'    Loop
'
'    ReplaceString = new_txt

ReplaceString = Replace(Txt, from_str, to_str)
End Function

Sub SaveCommands()

'Saves commands to file

If CheckForFile(DataFile) Then Kill DataFile

Open DataFile For Binary As #1
    Put #1, , Presets
    Put #1, , NumPresets
    Put #1, , RichColors
    
Close #1

a$ = App.Path + "\assisttool.dat"

h = FreeFile
Dim ard As Integer
ard = UBound(CurrentToolBar)
If CheckForFile(a$) Then Kill a$
Open a$ For Binary As h
    Put #h, , ard
    Put #h, , CurrentToolBar
Close h

End Sub

Function LoadCommands() As Boolean

'Loads commands from file

If CheckForFile(DataFile) Then
    Open DataFile For Binary As #1
        
        Get #1, , Presets
        Get #1, , NumPresets
        Get #1, , RichColors
        
    Close #1
    LoadCommands = True
End If

Dim erg(1 To 10) As String

a$ = App.Path + "\assisthl.dat"

h = FreeFile
If CheckForFile(a$) Then
    Open a$ For Binary As h
        Get #h, , erg
    Close h

    HLEXEPath = erg(1)
    HLExtraArgs = erg(2)
    HLQuitSA = Val(erg(3))
    HLSetAway = Val(erg(4))

End If

a$ = App.Path + "\assisttool.dat"

Dim ard1 As Integer

On Error Resume Next

h = FreeFile
If CheckForFile(a$) Then
    Open a$ For Binary As h
        Get #h, , ard1
        ReDim CurrentToolBar(0 To ard1)
        Get #h, , CurrentToolBar
    Close h
Else
    CopyDefaultToCurrent
End If


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

rr$ = Form1.Caption
f$ = ""

mn = Len(e$)
'code it into hex

For i = 1 To Len(e$)
    
    g$ = Hex$(Asc(Mid(e$, i, 1)))
    If Len(g$) = 1 Then g$ = "0" + g$
       
    ff$ = ff$ + g$
    
    If i Mod 2000 = 0 Then
        Form1.Caption = "Encoding EXE: " + Ts(Int((i / mn) * 100)) + "%"
        f$ = f$ + ff$
        ff$ = ""
        DoEvents
        
    End If
Next i
Form1.Caption = rr$
CompileEXE = f$

End Function

Function MakeHex(str As String) As String

For i = 1 To Len(str)
    
    g$ = Hex$(Asc(Mid(str, i, 1)))
    If Len(g$) = 1 Then g$ = "0" + g$
    ff$ = ff$ + g$

Next i

MakeHex = ff$

End Function

Function DeHex(str As String) As String

For i = 1 To Len(str) Step 2
    
    g$ = Chr(Hex2Dec(Mid(str, i, 2)))
    
    ff$ = ff$ + g$

Next i

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

'If a$ = "X1" Then 'welcome!
'    SendPacket "X2", LoginPass
'End If

'AddEvent "Got packet: " & a$ & ", Params: " & p$

If a$ = "X2" Then 'welcome!
    EncryptedMode = True
    SendPacket "HL", ""
End If

If a$ = "IC" Then 'incorrect password
    MessBox "Incorrect password!", vbCritical, "Incorrect Password"
    'Form1.TCP1.Close
End If

If a$ = "HI" Then 'welcome!
    AddEvent "**** Logged in."
    DownloadingNew = False
    LoadWindows
    MessBox p$, , "Welcome!", True, 3
    UpdateLogDetail
    PackageConnectPacket
    PrevConnection = True
End If

If a$ = "MS" Then 'message
    If MDIForm1.mnuSettingsIn(3).Checked = False Then
        MessBox p$, , "Server Message", , 4
    Else
        AddMsg "----------" + vbCrLf + "Server Message:" + vbCrLf + p$ + vbCrLf + "----------"
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
    If ChatFrozen = False Then
        UnPackageMessage p$
    Else
        n = UBound(FreezeText) + 1
        ReDim Preserve FreezeText(0 To n)
        FreezeText(n) = p$
        Form1.Command10.Caption = "Continue: " + Ts(n)
    End If
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
    If ButtonShowMode = 0 Then frmSelectBut.Show
    If ButtonShowMode = 1 Then frmCustomize.AddScripts
    If ButtonShowMode = 2 Then DoTheScript
    
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

If a$ = "GB" Then 'Connect Packet
    UnPackageGetBans p$
End If

If a$ = "NM" Then 'notice of mapchange
    NotifyMapChange
End If

If a$ = "AC" Then 'admin chat

    UnPackageAdminChat p$
    'AddChat p$
End If

If a$ = "MP" Then 'admin chat
    UnPackageMapProcess p$
End If

If a$ = "MM" Then 'admin chat
    UnPackageMenuScripts p$
End If

If a$ = "G1" Then 'admin chat
    UnPackageGameRequest p$
End If

If a$ = "G2" Then 'admin chat
    UnPackageGamePacket p$
End If

If a$ = "AL" Then 'admin BMP list
    UnPackageAdminBMPList p$
End If

If a$ = "AM" Then 'admin BMP file
    UnPackageAdminBMP p$
End If

If a$ = "AD" Then 'admin BMP file delete
    fn$ = App.Path + "\apics\" + p$
    If CheckForFile(fn$) Then Kill fn$
    
    
End If

If a$ = "U1" Then 'sendpacket
    SendPacket "U2", ""
End If

If a$ = "BE" Then 'beep
    PlayWaveRes 105
    AddChat "User " & p$ & " has beeped you.", "*", RGB(255, 255, 255), Time$
End If

If ShowWhiteBoard Then
    
    If a$ = "NS" Then 'new shape
        frmWhiteBoard.NewShapeLoad p$
    End If
    
    If a$ = "AS" Then 'all shapes
        frmWhiteBoard.NewShapeLoad p$
    End If
    
    If a$ = "CB" Then 'all shapes
        frmWhiteBoard.ClearBoard p$
    End If
    
    If a$ = "TC" Then 'all shapes
        frmWhiteBoard.ChangeText p$
    End If
    
    If a$ = "SM" Then 'move
        frmWhiteBoard.MoveObject p$
    End If
    If a$ = "DS" Then 'move
        frmWhiteBoard.DeleteObject p$
    End If
    
End If

If a$ = "Z3" Then 'status
    MDIForm1.StatusBar1.Panels(6).Text = "HLDS: " & p$
End If

If a$ = "A2" Then

    If ShowChat = True Then
        frmAdminChat.clearChat
    End If

    
End If


End Sub




Sub AddChat(Msg$, Nme$, col, tm$)




If Not ShowChat And MDIForm1.mnuSettingsIn(4).Checked Then
    SendPacket "SC", ""
    ShowChat = True
    Exit Sub
End If

If Not MDIForm1.mnuSettingsIn(4).Checked And Not ShowChat Then Exit Sub

frmAdminChat.AddChat Msg$, Nme$, col, tm$

'Dim Txt As String
'
'If ShowChat Or MDIForm1.mnuSettingsIn(4).Checked Then
'
'    Txt = "[" + tm$ + "] " + Nme$ + ": " + Msg$
'
'
'
'    frmAdminChat.RT2.Text = ""
'    frmAdminChat.RT2.SelStart = 0
'    frmAdminChat.RT2.SelLength = Len(Txt)
'    frmAdminChat.RT2.SelColor = RGB(255, 255, 255)
'
'    frmAdminChat.RT2.SelText = Txt
'    frmAdminChat.RT2.SelStart = Len("[" + tm$ + "] ")
'    frmAdminChat.RT2.SelLength = Len(Nme)
'    frmAdminChat.RT2.SelColor = col
'
'    frmAdminChat.RT2.SelStart = 0
'    frmAdminChat.RT2.SelLength = Len(frmAdminChat.RT2.Text)
'
'    frmAdminChat.Text1.SelStart = Len(frmAdminChat.Text1.Text)
'    frmAdminChat.Text1.SelText = vbCrLf
'    frmAdminChat.Text1.SelRTF = frmAdminChat.RT2.SelRTF
'    frmAdminChat.Text1.SelStart = Len(frmAdminChat.Text1.Text)
'
'    If Len(frmAdminChat.Text1.Text) > 5000 Then
'        frmAdminChat.Text1.SelStart = 0
'        frmAdminChat.Text1.SelLength = 500
'        frmAdminChat.Text1.SelText = ""
'        frmAdminChat.Text1.SelStart = Len(frmAdminChat.Text1.Text)
'    End If
'
''    frmAdminChat.Text1 = frmAdminChat.Text1 + vbCrLf + Msg$
''
''
''    If Len(frmAdminChat.Text1.Text) > 5000 Then
''        frmAdminChat.Text1 = Right(frmAdminChat.Text1, 4500)
''
''    End If
''
''    frmAdminChat.Text1.SelStart = Len(frmAdminChat.Text1)
'End If

End Sub

Sub DoTheScript()

For i = 1 To NumCommands
    
    If Commands(i).ScriptName = ScriptButtonName Then
    
    
        frmControlFill.ButIndex = i
        frmControlFill.Draw
        
        
        Exit For
    End If
Next i


End Sub

Sub NotifyMapChange()

If ShowMap Then

    'request new map data
    SendPacket "TE", ""
    SendPacket "MD", ""
    
End If

End Sub

Sub ShowServerLog(p$)

If CheckForFile(EditFileTemp) Then Kill EditFileTemp
Open EditFileTemp For Binary As #1
    Put #1, , p$
Close #1
'All Done

'Open with notepad
ShellExecute MDIForm1.hwnd, "open", EditFileTemp, vbNullString, vbNullString, SW_SHOW

End Sub

Sub ShowBSPEnts(p$)

If CheckForFile(FileLocalPath + ".ent") Then Kill FileLocalPath + ".ent"
Open FileLocalPath + ".ent" For Binary As #1
    Put #1, , p$
Close #1
'All Done

'Open with notepad
ShellExecute MDIForm1.hwnd, "open", FileLocalPath + ".ent", vbNullString, vbNullString, SW_SHOW

frmBSPEdit.Show

End Sub

Sub AddMsg(Txt As String, Optional ColTxt As String, Optional r As Byte, Optional g As Byte, Optional b As Byte, Optional TimeStamp As String)

If MDIForm1.mnuSettingsIn(2).Checked = True Then
    Txt = "[" + TimeStamp + "] " + Txt
    ColTxt = "[" + TimeStamp + "] " + ColTxt
End If
'add text to console

Txt = ReplaceString(Txt, vbCrLf, Chr(10))
Txt = ReplaceString(Txt, Chr(10), vbCrLf)


'Form1.Text1 = Form1.Text1 + Txt + vbCrLf
'If Len(Form1.Text1) > 5000 Then Form1.Text1 = Right(Form1.Text1, 4500)
'Form1.Text1.SelStart = Len(Form1.Text1)

If r = 0 And g = 0 And b = 0 Then r = 255: b = 255: g = 255


Form1.RT2.SelText = Txt
Form1.RT2.SelStart = 0
Form1.RT2.SelLength = Len(Txt)
Form1.RT2.SelColor = RGB(RichColors(1).r, RichColors(1).g, RichColors(1).b)
'frmmain.RichTextBox1.SelColor = RGB(RichColors(1).r, RichColors(1).g, RichColors(1).b)

If Len(ColTxt) > 0 Then Form1.RT2.SelLength = Len(ColTxt)
If Len(ColTxt) = 0 Then Form1.RT2.SelLength = Len(Txt)

If MDIForm1.mnuSettingsIn(1).Checked = True Then
    If r = 0 And g = 0 And b = 0 Then
    Else
        Form1.RT2.SelColor = RGB(r, g, b)
    End If
End If

Form1.RT2.SelLength = Len(Form1.RT2.Text)

frmMain.RichTextBox1.SelStart = Len(frmMain.RichTextBox1.Text)
frmMain.RichTextBox1.SelText = vbCrLf
frmMain.RichTextBox1.SelRTF = Form1.RT2.SelRTF
frmMain.RichTextBox1.SelStart = Len(frmMain.RichTextBox1.Text)


If Len(frmMain.RichTextBox1.Text) > 5000 Then
    frmMain.RichTextBox1.SelStart = 0
    frmMain.RichTextBox1.SelLength = 500
    frmMain.RichTextBox1.SelText = ""
    frmMain.RichTextBox1.SelStart = Len(frmMain.RichTextBox1.Text)
End If

End Sub

Function PackageDirList(Num, Optional FullPath As String)

cd$ = DirFullPath(Num)
If FullPath <> "" Then cd$ = FullPath

a$ = a$ + Chr(251)
a$ = a$ + cd$ + Chr(250)
a$ = a$ + Chr(251)

'compile it
For i = 1 To NumDirs(Num)
    a$ = a$ + Chr(251)
    If Num = 0 Then
        a$ = a$ + Ts(DirList0(i).DateTime) + Chr(250)
        a$ = a$ + DirList0(i).FullPath + Chr(250)
        a$ = a$ + DirList0(i).Name + Chr(250)
        a$ = a$ + DirList0(i).Size + Chr(250)
        a$ = a$ + Ts(DirList0(i).Type) + Chr(250)
    Else
        a$ = a$ + Ts(DirList1(i).DateTime) + Chr(250)
        a$ = a$ + DirList1(i).FullPath + Chr(250)
        a$ = a$ + DirList1(i).Name + Chr(250)
        a$ = a$ + DirList1(i).Size + Chr(250)
        a$ = a$ + Ts(DirList1(i).Type) + Chr(250)
    End If
    a$ = a$ + Chr(251)
Next i

'Return
PackageDirList = a$

End Function

Function UnPackageDirList(p$) As Integer


'extracts directory listing from the sent string
f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                m$ = Mid(a$, g, h - g)
                
                If i = 1 Then
                    'DirFullpath
                    Num = 1
                    If DirFullPath(0) = m$ Then Num = 0
                    If DirFullPath(1) = m$ Then Num = 1
                    
                    If DirFullPath(0) = DirFullPath(1) Then Num = LastRefresh
                    
                    DirFullPath(Num) = m$
                Else
                    Ii = i - 1
                    If Num = 0 Then
                        ReDim Preserve DirList0(0 To Ii)
                        If j = 1 Then DirList0(Ii).DateTime = CDate(m$)
                        If j = 2 Then DirList0(Ii).FullPath = m$
                        If j = 3 Then DirList0(Ii).Name = m$
                        If j = 4 Then DirList0(Ii).Size = m$
                        If j = 5 Then DirList0(Ii).Type = Val(m$)
                    Else
                        ReDim Preserve DirList1(0 To Ii)
                        If j = 1 Then DirList1(Ii).DateTime = CDate(m$)
                        If j = 2 Then DirList1(Ii).FullPath = m$
                        If j = 3 Then DirList1(Ii).Name = m$
                        If j = 4 Then DirList1(Ii).Size = m$
                        If j = 5 Then DirList1(Ii).Type = Val(m$)
                    End If
                End If
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

NumDirs(Num) = i - 1
UnPackageDirList = Num

End Function

Sub InterpritUsers(p$)
'extracts user info from the sent string

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then Users(i).Allowed = m$
                If j = 2 Then Users(i).Flags = Val(m$)
                If j = 3 Then Users(i).Name = m$
                If j = 4 Then Users(i).PassWord = m$
                If j = 5 Then Users(i).ICQ = m$
                If j = 6 Then Users(i).Email = m$
                
            End If
        Loop Until h = 0
    
    End If

Loop Until f = 0 Or e = 0
NumUsers = i

frmUserList.Show


End Sub

Sub SendUserEdit()

'compile it

For i = 1 To NumUsers
    a$ = a$ + Chr(251)
    a$ = a$ + Users(i).Allowed + Chr(250)
    a$ = a$ + Ts(Users(i).Flags) + Chr(250)
    a$ = a$ + Users(i).Name + Chr(250)
    a$ = a$ + Users(i).PassWord + Chr(250)
    a$ = a$ + Users(i).ICQ + Chr(250)
    a$ = a$ + Users(i).Email + Chr(250)
    a$ = a$ + Chr(251)
Next i

'all set, send it
SendPacket "EY", a$

End Sub

Sub UpdateLogDetail()

ll = 0

For i = 0 To Form1.LogDetail.Count - 1
    If Form1.LogDetail(i).Value = 1 Then ll = ll + (2 ^ i)
Next i

If ShowPlayers Then ll = ll + 2 ^ 9
If ShowUsers Then ll = ll + 2 ^ 10
If ShowWhiteBoard Then ll = ll + 2 ^ 11



SendPacket "LL", Ts(ll)


End Sub


Public Sub SendPacket(Cde As String, Params As String)

If SendingFile = True Then Exit Sub
    
'AddEvent "Sending: " & Cde & " - " & Params & ", Enc: " & EncryptedMode
    
Msg$ = Chr(255) + Cde + Chr(255) + LoginName + Chr(255) + LoginPass + Chr(255) + Params + Chr(255)
' encrypt

If EncryptedMode Then
    Msg$ = Encrypt(Msg$, LoginPass)
End If

a$ = Chr(254) + Chr(254) + Chr(254) + Msg$ + Chr(253) + Chr(253) + Chr(253)


If Form1.TCP1.State = sckConnected Then
    'send it in increments of 65000 bytes
    If Len(a$) <= 65000 Then
        Form1.TCP1.SendData a$
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
            
            Form1.TCP1.SendData b$
            'DoEvents
        Loop Until Len(b$) < 65000
    End If
End If





End Sub

Public Sub AttemptConnect(IP As String, Port As String, UserName As String, PassWord As String)

LoginName = UserName
LoginPass = PassWord
Form1.TCP1.RemoteHost = IP
Form1.TCP1.RemotePort = Val(Port)
Form1.TCP1.Connect

End Sub

Public Sub AddEvent(Txt As String)

'Form1.Text1 = Form1.Text1 + Txt + vbCrLf
'If Len(Form1.Text1) > 3000 Then Form1.Text1 = Right(Form1.Text1, 2990)
'Form1.Text1.SelStart = Len(Form1.Text1)
AddMsg Txt

End Sub

Sub UnPackageScripts(p$, Mde)
'extracts scripts from the sent string

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g Then
                m$ = Mid(a$, g, h - g)
                
                On Error Resume Next
                n = UBound(Commands)
                If n < i Then ReDim Preserve Commands(0 To i)
                
                If j = 1 And m$ <> "" Then Commands(i).Exec = m$
                If j = 2 Then Commands(i).MustHave = Val(m$)
                If j = 3 Then Commands(i).Name = m$
                If j = 4 Then Commands(i).NumParams = Val(m$)
                If j = 5 Then Commands(i).ScriptName = m$
                If j = 6 Then Commands(i).AutoMakeVars = CBool(m$)
                If j = 7 Then Commands(i).Group = m$
                If j = 8 Then Commands(i).LogExec = CBool(m$)
                If j = 9 Then Commands(i).Unused1 = m$
                If j = 10 Then Commands(i).unused2 = m$
                If j = 11 Then Commands(i).unused3 = m$
                If j = 12 Then
                    Commands(i).ScriptID = Val(m$)
                    
                    If Commands(i).ScriptID = 0 Then
                        Do
                            Randomize
                            newid = Int(Rnd * 30000) + 1
                            usd = 0
                            For k = 1 To i
                                If newid = Commands(k).ScriptID Then usd = 1: Exit For
                            Next k
                        Loop Until usd = 0
                        Commands(i).ScriptID = newid
                    End If
                    
                End If
                If j = 13 Then Commands(i).Unused5 = Val(m$)
                
                If j = 14 Then
                    Commands(i).NumButtons = Val(m$)
                    ReDim Commands(i).Buttons(0 To Val(m$))
                End If
                If j > 14 Then 'clan member list
                    
                    kk = (j - 15) Mod 5
                    k = ((j - 10) \ 5)
                    
                    If kk = 0 Then Commands(i).Buttons(k).ButtonName = m$
                    If kk = 1 Then Commands(i).Buttons(k).ButtonText = m$
                    If kk = 2 Then Commands(i).Buttons(k).OptionOff = m$
                    If kk = 3 Then Commands(i).Buttons(k).OptionOn = m$
                    If kk = 4 Then Commands(i).Buttons(k).Type = Val(m$)
                
                End If
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0
NumCommands = i

'If NumCommands < UBound(Commands) Then MsgBox "fuchy"

If Mde = 0 Then Form3.Show

End Sub

Sub UnPackageClans(p$)
'extracts clans from the sent string

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then Clans(i).Clan = m$
                If j = 2 Then Clans(i).JoinPass = m$
                If j = 3 Then Clans(i).Flags = Val(m$)
                If j = 4 Then Clans(i).NumMembers = Val(m$)
                If j > 4 Then 'clan member list
                    
                    k = ((j - 2) \ 3)
                    kk = (j - 5) Mod 3
                    
                    If kk = 0 Then Clans(i).Members(k).UIN = m$
                    If kk = 1 Then Clans(i).Members(k).LastIP = m$
                    If kk = 2 Then Clans(i).Members(k).Name = m$
                
                End If
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0
NumClans = i

frmClans.Show

End Sub


Sub UnPackageFilePacket(p$)
'extracts clans from the sent string

f = 0
i = 0
e = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStrRev(p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
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
    
    ShellExecute MDIForm1.hwnd, "open", EditFileTemp, vbNullString, vbNullString, SW_SHOW
    
    EditMode = True
    TheEditFile = FilePath

End If

If DownloadingNew = False Then
    frmFileBrowser.RefreshDir 0, DirFullPath(0)
    frmFileBrowser.RefreshDir 1, DirFullPath(1)
Else
    ShellExecute MDIForm1.hwnd, "open", FileLocalPath, vbNullString, vbNullString, SW_SHOW
    DownloadingNew = False
End If

Unload frmTransferProgress

End Sub

Sub PackageFileSend(LocFle As String, Fle As String)

'get the file
startimer = Timer



If CheckForFile(LocFle) Then

    mn = FileLen(LocFle)
    a$ = ""
    FileSize = mn
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
            
            SendPacket "F9", a$
            

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


Sub PackageAdminBMP(Fle As String)

'Send


If CheckForFile(Fle) Then

    mn = FileLen(Fle)
    a$ = ""
    h = FreeFile
    Open Fle For Binary As h
        
        ' read the file
        ret$ = Input(50000, #h)
        ret$ = Convert255(ret$)
        
        a$ = Chr(251)
        a$ = a$ + ret$ + Chr(250)
        a$ = a$ + Chr(251)
        
        ' send our BMP to the server.
        SendPacket "AM", a$
            
    Close h
End If

End Sub

Sub UnPackageAdminBMP(p$)

'extracts the BMP file

If p$ = "" Then Exit Sub


d$ = App.Path + "\apics"
If Dir(d$, vbDirectory) = "" Then MkDir d$
d$ = ""

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then AdminName$ = m$
                If j = 2 Then BMPFil$ = m$
                If j = 3 Then bmp$ = m$
                
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

bmp$ = DeCode255(bmp$)

'got the file

If CheckForFile(App.Path + "\apics\" + BMPFil$) Then Kill App.Path + "\apics\" + BMPFil$

If Dir(App.Path + "\apics", vbDirectory) = "" Then
    MkDir App.Path + "\apics"
End If

h = FreeFile
Open App.Path + "\apics\" + BMPFil$ For Binary As h
    Put #h, , bmp$
Close h

' At this point, reload the file in the chat or soemthing.

'Do this by deleting all the pictureboxes

If ShowChat Then

    n = frmAdminChat.Picture5.Count - 1
    
    For i = 1 To n
        Unload frmAdminChat.Picture5(i)
    Next i

End If

End Sub

Function Hex2Dec(InputData As String) As Double
If DebugMode Then LastCalled = "Hex2Dec"

''
''  Converts Hexadecimal to Decimal
''
Dim i As Integer
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
For i = Lenhex To 1 Step -1

HexStep = HexStep * 16
If HexStep = 0 Then
  HexStep = 1
End If

 If Mid(InputData, i, 1) = "0" Then
   DecOut = DecOut + (0 * HexStep)
 ElseIf Mid(InputData, i, 1) = "1" Then
   DecOut = DecOut + (1 * HexStep)
 ElseIf Mid(InputData, i, 1) = "2" Then
   DecOut = DecOut + (2 * HexStep)
 ElseIf Mid(InputData, i, 1) = "3" Then
   DecOut = DecOut + (3 * HexStep)
 ElseIf Mid(InputData, i, 1) = "4" Then
   DecOut = DecOut + (4 * HexStep)
 ElseIf Mid(InputData, i, 1) = "5" Then
   DecOut = DecOut + (5 * HexStep)
 ElseIf Mid(InputData, i, 1) = "6" Then
   DecOut = DecOut + (6 * HexStep)
 ElseIf Mid(InputData, i, 1) = "7" Then
   DecOut = DecOut + (7 * HexStep)
 ElseIf Mid(InputData, i, 1) = "8" Then
   DecOut = DecOut + (8 * HexStep)
 ElseIf Mid(InputData, i, 1) = "9" Then
   DecOut = DecOut + (9 * HexStep)
 ElseIf Mid(InputData, i, 1) = "A" Then
   DecOut = DecOut + (10 * HexStep)
 ElseIf Mid(InputData, i, 1) = "B" Then
   DecOut = DecOut + (11 * HexStep)
 ElseIf Mid(InputData, i, 1) = "C" Then
   DecOut = DecOut + (12 * HexStep)
 ElseIf Mid(InputData, i, 1) = "D" Then
   DecOut = DecOut + (13 * HexStep)
 ElseIf Mid(InputData, i, 1) = "E" Then
   DecOut = DecOut + (14 * HexStep)
 ElseIf Mid(InputData, i, 1) = "F" Then
   DecOut = DecOut + (15 * HexStep)
 Else
 End If

Next i

Hex2Dec = DecOut

eds:
End Function
Sub UnPackageUpdate(p$)
'extracts time, map, and players from the string

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
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
                If j = 4 Then TeamNames(1) = m$
                If j = 5 Then TeamNames(2) = m$
                If j = 6 Then TeamNames(3) = m$
                If j = 7 Then TeamNames(4) = m$
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

UpdateLabel

If ShowMap = True Then frmMap.Caption = "Map - " + MapName

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

Form1.lblUpdate = g$

End Sub

Sub UnPackageSpeech(p$)
'extracts clans from the sent string

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                m$ = Mid(a$, g, h - g)
                ReDim Preserve Speech(0 To i)
                
                If j = 1 Then Speech(i).ClientText = m$
                If j = 2 Then
                    Speech(i).NumAnswers = Val(m$)
                    ReDim Speech(i).Answers(0 To Val(m$))
                End If
                If j > 2 Then 'speech answer list
                    k = j - 2
                    Speech(i).Answers(k) = m$
                End If
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0
NumSpeech = i

frmSpeech.Show

End Sub

Sub UnPackageRealPlayers(p$)
'extracts real players from the sent string

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                m$ = Mid(a$, g, h - g)
                ReDim Preserve RealPlayers(0 To i)
                
                If j = 1 Then RealPlayers(i).LastName = m$
                If j = 2 Then RealPlayers(i).RealName = m$
                If j = 3 Then RealPlayers(i).UniqueID = m$
                If j = 4 Then RealPlayers(i).LastTime = CDate(m$)
                If j = 5 Then RealPlayers(i).Flags = Val(m$)
                If j = 6 Then RealPlayers(i).Points = Ts(Val(m$))
                If j = 7 Then RealPlayers(i).TimesSeen = m$
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

NumRealPlayers = i
frmReal.Show

End Sub

Sub UnPackagePlayers(p$)
'extracts scripts from the sent string

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
        Players(i).RealName = ""
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then Players(i).Class = Val(m$)
                If j = 2 Then Players(i).IP = m$
                If j = 3 Then Players(i).Name = m$
                If j = 4 Then Players(i).Team = Val(m$)
                If j = 5 Then Players(i).UniqueID = m$
                If j = 6 Then Players(i).UserID = Val(m$)
                If j = 7 Then Players(i).Status = Val(m$)
                If j = 8 Then
                    Players(i).RealName = m$
                End If
                If j = 9 Then Players(i).pos.X = Val(m$)
                If j = 10 Then Players(i).pos.Y = Val(m$)
                If j = 11 Then Players(i).pos.Z = Val(m$)
                If j = 12 Then Players(i).JoinTime = CDate(m$)
                If j = 13 Then Players(i).EntryName = m$
                If j = 14 Then Players(i).NumKickVotes = Val(m$)
                If j = 15 Then Players(i).Port = Val(m$)
                If j = 16 Then Players(i).ShutUp = Val(m$)
                If j = 17 Then Players(i).Warn = Val(m$)
                If j = 18 Then Players(i).Points = Val(m$)
                If j = 19 Then Players(i).LastEvent = CDate(m$)
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0
NumPlayers = i
UpdatePlayerList

If ShowMap Then frmMap.Update2



'Form6.Show

End Sub

Sub UnPackageConnectUsers(p$)
'extracts connected users from the sent string

For i = 1 To NumConnectUsers
    ConnectUsers(i).IP = ""
    ConnectUsers(i).Name = ""
    ConnectUsers(i).Version = ""
Next i
NumConnectUsers = 0

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then ConnectUsers(i).IP = m$: NumConnectUsers = NumConnectUsers + 1
                If j = 2 Then ConnectUsers(i).Name = m$
                If j = 3 Then ConnectUsers(i).Version = m$
                If j = 4 Then ConnectUsers(i).AwayMode = Val(m$)
                If j = 5 Then ConnectUsers(i).AwayMsg = m$
                If j = 6 Then ConnectUsers(i).IdleTime = Val(m$)
                
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

UpdateUsersList

End Sub

Sub UnPackageTeleporters(p$)
'extracts connected users from the sent string

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                ReDim Preserve Tele(0 To i)
                m$ = Mid(a$, g, h - g)
                    
                If j = 1 Then Tele(i).Name = m$
                If j = 2 Then Tele(i).X = Val(m$)
                If j = 3 Then Tele(i).Y = Val(m$)
                If j = 4 Then Tele(i).Z = Val(m$)
                
                
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

NumTele = i
If ShowMap Then frmMap.Update2

End Sub

Sub UnPackageMessages(p$)
'extracts connected users from the sent string
NumMessages = 0

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                NumMessages = i
                ReDim Preserve Messages(0 To i)
                                
                If j = 1 Then Messages(i).Flags = Val(m$)
                If j = 2 Then Messages(i).MsgFor = m$
                If j = 3 Then Messages(i).MsgFrom = m$
                If j = 4 Then Messages(i).MsgID = Val(m$)
                If j = 5 Then Messages(i).MsgSubj = m$
                If j = 6 Then Messages(i).MsgText = m$
                If j = 7 Then Messages(i).MsgTimeSent = CDate(m$)
                
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0
frmMessageList.Show
frmMessageList.RefreshMessageList


End Sub
Sub LoadWindows()

showi = CBool(GetSetting("Server Assistant Client", "Window", "showchat", 0))
If showi Then
    SendPacket "SC", ""
    frmAdminChat.Show
End If

showi = CBool(GetSetting("Server Assistant Client", "Window", "showplayers", 0))
If showi Then
    ShowPlayers = True
    Form6.Show
End If

showi = CBool(GetSetting("Server Assistant Client", "Window", "showmap", 0))
If showi Then
    frmMap.Show
End If

showi = CBool(GetSetting("Server Assistant Client", "Window", "showusers", 0))
If showi Then
    SendPacket "CU", "": ShowUsers = True: frmConnectUsers.Show
End If

End Sub
Sub UnPackageWebColors(p$)
'extracts web colors from the sent string

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
           
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then Web.Colors(i).r = Val(m$)
                If j = 2 Then Web.Colors(i).g = Val(m$)
                If j = 3 Then Web.Colors(i).b = Val(m$)
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

frmWebColor.Show

End Sub

Sub PackageScripts(Optional ChangedMode As Boolean)

'compiles and sends the script info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'users has:
'Commands.Exec
'Commands.MustHave
'Commands.Name
'Commands.NumParams

'compile it

For i = 1 To NumCommands
    
    doit = 1
    If ChangedMode = True And Commands(i).Changed = False Then doit = 0
    
    If doit = 1 Then
        a$ = a$ + Chr(251)
        a$ = a$ + Commands(i).Exec + Chr(250)
        a$ = a$ + Ts(Commands(i).MustHave) + Chr(250)
        a$ = a$ + Commands(i).Name + Chr(250)
        a$ = a$ + Ts(Commands(i).NumParams) + Chr(250)
        a$ = a$ + Commands(i).ScriptName + Chr(250)
        a$ = a$ + Ts(CInt(Commands(i).AutoMakeVars)) + Chr(250)
        a$ = a$ + Commands(i).Group + Chr(250)
        a$ = a$ + Ts(CInt(Commands(i).LogExec)) + Chr(250)
        a$ = a$ + Commands(i).Unused1 + Chr(250)
        a$ = a$ + Commands(i).unused2 + Chr(250)
        a$ = a$ + Commands(i).unused3 + Chr(250)
        a$ = a$ + Ts(Commands(i).ScriptID) + Chr(250)
        a$ = a$ + Ts(Commands(i).Unused5) + Chr(250)
        
        a$ = a$ + Ts(Commands(i).NumButtons) + Chr(250)
        For j = 1 To Commands(i).NumButtons
            a$ = a$ + Commands(i).Buttons(j).ButtonName + Chr(250)
            a$ = a$ + Commands(i).Buttons(j).ButtonText + Chr(250)
            a$ = a$ + Commands(i).Buttons(j).OptionOff + Chr(250)
            a$ = a$ + Commands(i).Buttons(j).OptionOn + Chr(250)
            a$ = a$ + Ts(Commands(i).Buttons(j).Type) + Chr(250)
        Next j
        a$ = a$ + Chr(251)
    End If
    Commands(i).Changed = False
Next i

'all set, send it
If ChangedMode = False Then SendPacket "ED", a$
If ChangedMode = True And a$ <> "" Then SendPacket "E1", a$

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
SendPacket "O1", a$

End Sub

Sub UnPackageKickBans(p$)
'extracts kickbans from the sent string

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then KickBans(i).Clan = m$
                If j = 2 Then KickBans(i).Name = m$
                If j = 3 Then KickBans(i).Type = Val(m$)
                If j = 4 Then KickBans(i).UID = m$
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0
NumKickBans = i

frmKickBan.Show


End Sub

Sub UnPackageAdminChat(p$)
'extracts kickbans from the sent string

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then tm$ = m$
                If j = 2 Then Msg$ = m$
                If j = 3 Then col = Val(m$)
                If j = 4 Then Nme$ = m$
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

AddChat Msg$, Nme$, col, tm$

End Sub



Sub UnPackageMapProcess(p$)
'extracts kickbans from the sent string

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                    
                ReDim Preserve MapProcess(0 To i)
                If j = 1 Then MapProcess(i).LastTimePlayed = CDate(m$)
                If j = 2 Then MapProcess(i).MapName = m$
                If j = 3 Then MapProcess(i).TimesPlayed = Val(m$)
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0
NumMapProcess = i

frmMapProcess.Show
frmMapProcess.DrawIt

End Sub


Sub UnPackageMessageUsers(p$)
'extracts kickbans from the sent string
'Load frmNewMessage
frmNewMessage.Combo1.Clear
frmNewMessage.Combo1.AddItem "(ALL)"
f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
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
                If m$ = frmNewMessage.ReplyTo Then kk = i
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

If frmNewMessage.ReplySubj <> "" Then frmNewMessage.Text2 = frmNewMessage.ReplySubj: frmNewMessage.ReplySubj = ""
If frmNewMessage.ReplyText <> "" Then frmNewMessage.Text1 = frmNewMessage.ReplyText: frmNewMessage.ReplyText = ""
If frmNewMessage.ReplyTo <> "" Then
    'frmNewMessage.Combo1.ListIndex = kk
    
    For i = 0 To frmNewMessage.Combo1.ListCount - 1
        If frmNewMessage.Combo1.List(i) = frmNewMessage.ReplyTo Then
            frmNewMessage.Combo1.ListIndex = i
            frmNewMessage.ReplyTo = ""
            Exit For
        End If
    Next i
    
End If


If frmNewMessage.Combo1.ListIndex = -1 Then frmNewMessage.Combo1.ListIndex = 1

frmNewMessage.Show

End Sub

Sub UnPackageLogSearch(p$)
'extracts kickbans from the sent string
NumLogFound = 0

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                NumLogFound = i
                ReDim Preserve LogFound(0 To i)
                If j = 1 Then LogFound(i).LogFile = m$
                If j = 2 Then LogFound(i).LogLine = m$
            
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
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
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
                
                ' new stuff
                If j = 7 Then ServerStart.AutoRestart = CBool(m$)
                If j = 8 Then ServerStart.UseFeature = CBool(m$)
                If j = 9 Then ServerStart.CommandLine = m$
                If j = 10 Then ServerStart.HLDSDir = m$
                If j = 11 Then ServerStart.HLDSPath = m$
            
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
SendPacket "WI", a$

End Sub

Sub UnPackageWebInfo(p$)
'extracts web info from the sent string

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
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
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
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
                If j = 24 Then General.CustomFlag1 = m$
                If j = 25 Then General.CustomFlag2 = m$
                If j = 26 Then General.CustomFlag3 = m$
                If j = 27 Then General.CustomFlag4 = m$
                If j = 28 Then General.MaxFileSend = Val(m$)
                If j = 29 Then General.NewestClient = m$
                If j = 30 Then General.SendToDisco = Val(m$)
                
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
        For i = 1 To NumPlayers
            If LCase(Players(i).Name) = LCase(nm2$) Then
                cc = Players(i).Team
                Exit For
            End If
        Next i
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
    i = 0
    Do
    
        e = InStr(f + 1, p$, Chr(251))
        f = InStr(e + 1, p$, Chr(251))
        'extract this section
        
        If e > 0 And f > e Then
            a$ = Mid(p$, e + 1, f - e - 1)
            i = i + 1
                
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
                    If j = 4 Then tm$ = m$
                    
                    
                End If
            Loop Until h = 0
        
        End If
    Loop Until f = 0 Or e = 0
    
    If cc <= 4 Then cc = cc + 1
    
    r = RichColors(cc).r
    g = RichColors(cc).g
    b = RichColors(cc).b
    
    LastMsg = nm$ + Txt


    AddMsg Txt, nm$, CByte(r), CByte(g), CByte(b), tm$
End If

End Sub

Sub UnPackageConnectPacket(p$)
'extracts general stuff from the sent string
Dim jjj As Boolean

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
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
                If j = 5 Then HLIP = m$
                If j = 6 Then HLPort = m$
                If j = 7 Then HLGame = m$
                If j = 8 Then CustomFlag1 = m$
                If j = 9 Then CustomFlag2 = m$
                If j = 10 Then CustomFlag3 = m$
                If j = 11 Then CustomFlag4 = m$
                If j = 12 Then TeamNames(1) = m$
                If j = 13 Then TeamNames(2) = m$
                If j = 14 Then TeamNames(3) = m$
                If j = 15 Then TeamNames(4) = m$
                If j = 16 Then NewestClient = m$

            End If
        Loop Until h = 0 Or h >= Len(a$)
    
    End If
Loop Until f = 0 Or e = 0


MDIForm1.mnuAdminIn(30).Visible = jjj
MDIForm1.mnuAdminIn(20).Visible = jjj
ChangeButton 17, 2, jjj
ChangeButton 17, 3, jjj


MDIForm1.mnuFunctionsIn(4).Visible = DllEnabled
MDIForm1.mnuFunctionsIn(5).Visible = DllEnabled
MDIForm1.mnuFunctionsIn(11).Visible = DllEnabled
MDIForm1.mnuFunctionsIn(13).Visible = DllEnabled
MDIForm1.mnuFunctionsIn(6).Visible = DllEnabled
MDIForm1.mnuFunctionsIn(7).Visible = DllEnabled

MDIForm1.mnuFunctionsMore(4).Visible = DllEnabled
MDIForm1.mnuWindowsIn(2).Visible = DllEnabled


If MDIForm1.mnuFunctionsIn(6).Enabled = True Then
    If GameMode = 2 Then MDIForm1.mnuFunctionsIn(6).Enabled = False
End If

MDIForm1.StatusBar1.Panels(2).Text = "Server: " + ServVersion

CheckNewestClient

End Sub

Sub ChangeButton(Bt, Tp, Vl As Variant)

With MDIForm1.Toolbar1.Buttons

For i = 1 To .Count
    If Val(.Item(i).Tag) = Bt Then
        
        If Tp = 1 Then .Item(i).Value = Vl
        If Tp = 2 Then .Item(i).Enabled = Vl
        If Tp = 3 Then .Item(i).Visible = Vl
        
    
    End If
Next i
End With

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

For i = 1 To 21
        a$ = a$ + Chr(251)
        a$ = a$ + Ts(Web.Colors(i).r) + Chr(250)
        a$ = a$ + Ts(Web.Colors(i).g) + Chr(250)
        a$ = a$ + Ts(Web.Colors(i).b) + Chr(250)
        a$ = a$ + Chr(251)
Next i

'all set, send it
SendPacket "WC", a$

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

' new stuff
a$ = a$ + Ts(CInt(ServerStart.AutoRestart)) + Chr(250)
a$ = a$ + Ts(CInt(ServerStart.UseFeature)) + Chr(250)
a$ = a$ + ServerStart.CommandLine + Chr(250)
a$ = a$ + ServerStart.HLDSDir + Chr(250)
a$ = a$ + ServerStart.HLDSPath + Chr(250)

a$ = a$ + Chr(251)

'all set, send it
SendPacket "ER", a$

End Sub

Sub PackageSearchStart(Text As String, Check1 As Integer, FromDay As Date, ToDay As Date, _
SearchSubs As Integer, SearchPath As String, ExactPhrase As Boolean, AllWords As Integer, SearchSaysOnly As Integer)

a$ = a$ + Chr(251)
a$ = a$ + Text + Chr(250)
a$ = a$ + Ts(Check1) + Chr(250)
a$ = a$ + Ts(CDbl(FromDay)) + Chr(250)
a$ = a$ + Ts(CDbl(ToDay)) + Chr(250)
a$ = a$ + Ts(SearchSubs) + Chr(250)
a$ = a$ + SearchPath + Chr(250)
a$ = a$ + Ts(CInt(ExactPhrase)) + Chr(250)
a$ = a$ + Ts(AllWords) + Chr(250)
a$ = a$ + Ts(SearchSaysOnly) + Chr(250)
a$ = a$ + Chr(251)

'all set, send it
SendPacket "LS", a$

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

For i = 1 To NumSpeech
    a$ = a$ + Chr(251)
    a$ = a$ + Speech(i).ClientText + Chr(250)
    a$ = a$ + Ts(Speech(i).NumAnswers) + Chr(250)
    For j = 1 To Speech(i).NumAnswers
        a$ = a$ + Speech(i).Answers(j) + Chr(250)
    Next j
    a$ = a$ + Chr(251)
    If Len(a$) > 1000 Then b$ = b$ & a$: a$ = ""
Next i

b$ = b$ & a$: a$ = ""
'all set, send it
SendPacket "SP", b$

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

For i = 1 To NumKickBans
    a$ = a$ + Chr(251)
    a$ = a$ + KickBans(i).Clan + Chr(250)
    a$ = a$ + KickBans(i).Name + Chr(250)
    a$ = a$ + Ts(KickBans(i).Type) + Chr(250)
    a$ = a$ + KickBans(i).UID + Chr(250)
    a$ = a$ + Chr(251)
Next i

'all set, send it
SendPacket "KD", a$

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

For i = 1 To NumClans
    a$ = a$ + Chr(251)
    a$ = a$ + Clans(i).Clan + Chr(250)
    a$ = a$ + Clans(i).JoinPass + Chr(250)
    a$ = a$ + Ts(Clans(i).Flags) + Chr(250)
    a$ = a$ + Ts(Clans(i).NumMembers) + Chr(250)
    For j = 1 To Clans(i).NumMembers
        a$ = a$ + Clans(i).Members(j).UIN + Chr(250)
        a$ = a$ + Clans(i).Members(j).LastIP + Chr(250)
        a$ = a$ + Clans(i).Members(j).Name + Chr(250)
    Next j
    a$ = a$ + Chr(251)
Next i

'all set, send it
SendPacket "CM", a$

End Sub

Sub PackageSwears()

'compile it

For i = 1 To NumSwears
    a$ = a$ + Chr(251)
    a$ = a$ + Swears(i).BadWord + Chr(250)
    a$ = a$ + Ts(Swears(i).Flags) + Chr(250)
    a$ = a$ + Swears(i).Replacement + Chr(250)
    a$ = a$ + Chr(251)
Next i

'all set, send it
SendPacket "SW", a$

End Sub

Sub UnPackageSwears(p$)

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                m$ = Mid(a$, g, h - g)
                ReDim Preserve Swears(0 To i)
                
                If j = 1 Then Swears(i).BadWord = m$
                If j = 2 Then Swears(i).Flags = Val(m$)
                If j = 3 Then Swears(i).Replacement = m$
                
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

NumSwears = i

frmSwear.Show

End Sub


Sub UnPackageGetBans(p$)

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                m$ = Mid(a$, g, h - g)
                
                ReDim Preserve ServerBans(0 To i)
                
                If j = 1 Then ServerBans(i).BannedAt = m$
                If j = 2 Then ServerBans(i).BanTime = m$
                If j = 3 Then ServerBans(i).EntryName = m$
                If j = 4 Then ServerBans(i).IP = m$
                If j = 5 Then ServerBans(i).Map = m$
                If j = 6 Then ServerBans(i).Name = m$
                If j = 7 Then ServerBans(i).RealName = m$
                If j = 8 Then ServerBans(i).Reason = m$
                If j = 9 Then ServerBans(i).UIDs = m$
                
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

NumCurrBans = i

frmBans.Show

End Sub

Sub UnPackageAdminBMPList(p$)

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                m$ = Mid(a$, g, h - g)
                
                ReDim Preserve AdminBMP(0 To i)
                
                If j = 1 Then AdminBMP(i).AdminName = m$
                If j = 2 Then AdminBMP(i).BMPFile = m$
                
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

NumAdminBMP = i

' Do something here, like maybe decide which files we DONT have and then ask the server for them

'ok, check which of the files we need.


For i = 1 To NumAdminBMP

    'see if we have this file
    
    If CheckForFile(App.Path + "\apics\" + AdminBMP(i).BMPFile) Then
    
    Else
        
        'request the file for this admin
        SendPacket "AR", AdminBMP(i).AdminName
        DoEvents
    End If
Next i


End Sub

Sub UnPackageMenuScripts(p$)

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                m$ = Mid(a$, g, h - g)
                
                ReDim Preserve MenuScripts(0 To i)
                
                If j = 1 Then MenuScripts(i).AskForQuestion = m$
                If j = 2 Then MenuScripts(i).MenuName = m$
                If j = 3 Then MenuScripts(i).ScriptName = m$
                
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

NumMenuScripts = i

FillScriptsMenu

End Sub

Sub PackageConnectPacket()
If DebugMode Then LastCalled = "PackageConnectPacket"

'compiles and sends the GENERAL info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'compile it

a$ = a$ + Chr(251)
a$ = a$ + Ts(App.Major) + "." + Ts(App.Minor) + "." + Format(Ts(App.Revision), "0000") + Chr(250)
a$ = a$ + Ts(HiddenNow) + Chr(250)
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
For i = 0 To 6
    a$ = a$ + Ts(CInt(NewEvent.Days(i))) + Chr(250)
Next i
a$ = a$ + Ts(NewEvent.Every) + Chr(250)
a$ = a$ + str(NewEvent.FirstCheck) + Chr(250)
a$ = a$ + Ts(NewEvent.Mde) + Chr(250)
a$ = a$ + NewEvent.ScriptName + Chr(250)
a$ = a$ + Ts(NewEvent.Times) + Chr(250)
a$ = a$ + Ts(NewEvent.WhatToDo) + Chr(250)
a$ = a$ + NewEvent.Name + Chr(250)
a$ = a$ + Chr(251)

'all set, send it
SendPacket "AE", a$

End Sub


Sub UnPackageEvents(p$)
'extracts event list from the sent string

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                
                m$ = Mid(a$, g, h - g)
                
                ReDim Preserve Events(0 To i)
                
                If j = 1 Then Events(i).ComPara = m$
                If j >= 2 And j <= 8 Then Events(i).Days(j - 2) = CBool(m$)
                If j = 9 Then Events(i).Every = Val(m$)
                If j = 10 Then Events(i).FirstCheck = CDate(m$)
                If j = 11 Then Events(i).Mde = Val(m$)
                If j = 12 Then Events(i).ScriptName = m$
                If j = 13 Then Events(i).Times = Val(m$)
                If j = 14 Then Events(i).WhatToDo = Val(m$)
                If j = 15 Then Events(i).Name = m$
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

NumEvents = i

frmEvents.Show
frmEvents.UpdateList


End Sub

Sub UnPackageMapData(p$)
'extracts event list from the sent string

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
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
                
                If i - 1 < 64 And j - 1 < 64 Then MapArray(i - 1, j - 1) = Val(m$)
            
            End If
        Loop Until h = 0
    DoEvents
    
    End If
Loop Until f = 0 Or e = 0

If ShowMap Then frmMap.Draw

End Sub

Function FindPlayer(UsID$) As Integer

'searches the player records for a certain player

For i = 1 To NumPlayers
    If Players(i).UserID = UsID$ Then j = i: Exit For
Next i

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
a$ = a$ + General.CustomFlag1 + Chr(250)
a$ = a$ + General.CustomFlag2 + Chr(250)
a$ = a$ + General.CustomFlag3 + Chr(250)
a$ = a$ + General.CustomFlag4 + Chr(250)
a$ = a$ + Ts(General.MaxFileSend) + Chr(250)
a$ = a$ + General.NewestClient + Chr(250)
a$ = a$ + Ts(General.SendToDisco) + Chr(250)
a$ = a$ + Chr(251)

'all set, send it
SendPacket "GI", a$

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

CurrWnd = GetWindow(MDIForm1.hwnd, GW_HWNDFIRST)

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
    For i = 1 To e
        OutStr = "0" + OutStr
    Next i
Else
    OutStr = Ts(-(Val(OutStr)))
    For i = 1 To e
        OutStr = "0" + OutStr
    Next i
    OutStr = " -" + OutStr

End If


Numberize = OutStr


End Function

Sub LoadDefaultToolbar()
ReDim DefaultToolBar(1 To 36)

DefaultToolBar(1).Description = "Start Script"
DefaultToolBar(1).Tag = "1"
DefaultToolBar(1).IconID = 10
DefaultToolBar(1).Type = 0

DefaultToolBar(2).Description = "Map"
DefaultToolBar(2).Tag = "2"
DefaultToolBar(2).IconID = 6
DefaultToolBar(2).Type = 0

DefaultToolBar(3).Description = "File Manager"
DefaultToolBar(3).Tag = "3"
DefaultToolBar(3).IconID = 2
DefaultToolBar(3).Type = 0

DefaultToolBar(4).Description = "Join Server Now"
DefaultToolBar(4).Tag = "4"
DefaultToolBar(4).IconID = 7
DefaultToolBar(4).Type = 0

DefaultToolBar(5).Description = "Exit Program"
DefaultToolBar(5).Tag = "5"
DefaultToolBar(5).IconID = 8
DefaultToolBar(5).Type = 0

DefaultToolBar(6).Description = "Edit Scripts"
DefaultToolBar(6).Tag = "6"
DefaultToolBar(6).IconID = 3
DefaultToolBar(6).Type = 0

DefaultToolBar(7).Description = "Edit Users"
DefaultToolBar(7).Tag = "7"
DefaultToolBar(7).IconID = 21
DefaultToolBar(7).Type = 0

DefaultToolBar(8).Description = "Edit Kick/Ban List"
DefaultToolBar(8).Tag = "8"
DefaultToolBar(8).IconID = 19
DefaultToolBar(8).Type = 0

DefaultToolBar(9).Description = "Edit Server Info"
DefaultToolBar(9).Tag = "9"
DefaultToolBar(9).IconID = 27
DefaultToolBar(9).Type = 0

DefaultToolBar(10).Description = "Edit Clans"
DefaultToolBar(10).Tag = "10"
DefaultToolBar(10).IconID = 18
DefaultToolBar(10).Type = 0

DefaultToolBar(11).Description = "Edit Admin Speech"
DefaultToolBar(11).Tag = "11"
DefaultToolBar(11).IconID = 5
DefaultToolBar(11).Type = 0

DefaultToolBar(12).Description = "Edit RealPlayers"
DefaultToolBar(12).Tag = "12"
DefaultToolBar(12).IconID = 17
DefaultToolBar(12).Type = 0

DefaultToolBar(13).Description = "Edit Web Info"
DefaultToolBar(13).Tag = "13"
DefaultToolBar(13).IconID = 9
DefaultToolBar(13).Type = 0

DefaultToolBar(14).Description = "Edit General Info"
DefaultToolBar(14).Tag = "14"
DefaultToolBar(14).IconID = 15
DefaultToolBar(14).Type = 0

DefaultToolBar(15).Description = "Edit Events"
DefaultToolBar(15).Tag = "15"
DefaultToolBar(15).IconID = 12
DefaultToolBar(15).Type = 0

DefaultToolBar(16).Description = "Edit Bad Word List"
DefaultToolBar(16).Tag = "16"
DefaultToolBar(16).IconID = 16
DefaultToolBar(16).Type = 0

DefaultToolBar(17).Description = "Hidden Mode"
DefaultToolBar(17).Tag = "17"
DefaultToolBar(17).IconID = 14
DefaultToolBar(17).Type = 1

DefaultToolBar(18).Description = "Change Password"
DefaultToolBar(18).Tag = "18"
DefaultToolBar(18).IconID = 20
DefaultToolBar(18).Type = 0

DefaultToolBar(19).Description = "Show Names in Colour"
DefaultToolBar(19).Tag = "19"
DefaultToolBar(19).IconID = 11
DefaultToolBar(19).Type = 1

DefaultToolBar(20).Description = "Timestamp Console Messages"
DefaultToolBar(20).Tag = "20"
DefaultToolBar(20).IconID = 13
DefaultToolBar(20).Type = 1

DefaultToolBar(21).Description = "Show Messages in Console"
DefaultToolBar(21).Tag = "21"
DefaultToolBar(21).IconID = 25
DefaultToolBar(21).Type = 1

DefaultToolBar(22).Description = "Interface Colours"
DefaultToolBar(22).Tag = "22"
DefaultToolBar(22).IconID = 1
DefaultToolBar(22).Type = 0

DefaultToolBar(23).Description = "Configure Half-Life"
DefaultToolBar(23).Tag = "23"
DefaultToolBar(23).IconID = 30
DefaultToolBar(23).Type = 0

DefaultToolBar(24).Description = "Players Window"
DefaultToolBar(24).Tag = "24"
DefaultToolBar(24).IconID = 32
DefaultToolBar(24).Type = 0

DefaultToolBar(25).Description = "Logged In Users"
DefaultToolBar(25).Tag = "25"
DefaultToolBar(25).IconID = 28
DefaultToolBar(25).Type = 0

DefaultToolBar(26).Description = "Open Mailbox"
DefaultToolBar(26).Tag = "26"
DefaultToolBar(26).IconID = 29
DefaultToolBar(26).Type = 0

DefaultToolBar(27).Description = "Compose Message"
DefaultToolBar(27).Tag = "27"
DefaultToolBar(27).IconID = 26
DefaultToolBar(27).Type = 0

DefaultToolBar(28).Description = "View Server Log"
DefaultToolBar(28).Tag = "28"
DefaultToolBar(28).IconID = 24
DefaultToolBar(28).Type = 0

DefaultToolBar(29).Description = "View App Log"
DefaultToolBar(29).Tag = "29"
DefaultToolBar(29).IconID = 23
DefaultToolBar(29).Type = 0

DefaultToolBar(30).Description = "Log Search"
DefaultToolBar(30).Tag = "30"
DefaultToolBar(30).IconID = 4
DefaultToolBar(30).Type = 0

DefaultToolBar(31).Description = "Admin Chat"
DefaultToolBar(31).Tag = "31"
DefaultToolBar(31).IconID = 31
DefaultToolBar(31).Type = 0

DefaultToolBar(32).Description = "Whiteboard"
DefaultToolBar(32).Tag = "32"
DefaultToolBar(32).IconID = 33
DefaultToolBar(32).Type = 0

DefaultToolBar(33).Description = "Set Away Mode"
DefaultToolBar(33).Tag = "33"
DefaultToolBar(33).IconID = 34
DefaultToolBar(33).Type = 0

DefaultToolBar(34).Description = "Cancel Away Mode"
DefaultToolBar(34).Tag = "34"
DefaultToolBar(34).IconID = 35
DefaultToolBar(34).Type = 0

DefaultToolBar(35).Description = "Start Game Server"
DefaultToolBar(35).Tag = "35"
DefaultToolBar(35).IconID = 36
DefaultToolBar(35).Type = 0

DefaultToolBar(36).Description = "Stop Game Server"
DefaultToolBar(36).Tag = "36"
DefaultToolBar(36).IconID = 37
DefaultToolBar(36).Type = 0

End Sub

Sub CopyDefaultToCurrent()

ReDim CurrentToolBar(0 To UBound(DefaultToolBar))

For i = 1 To UBound(DefaultToolBar)
    
    CurrentToolBar(i).Description = DefaultToolBar(i).Description
    CurrentToolBar(i).Tag = DefaultToolBar(i).Tag
    CurrentToolBar(i).IconID = DefaultToolBar(i).IconID
    CurrentToolBar(i).Type = DefaultToolBar(i).Type
Next i



End Sub

Sub ApplyToToolbar()

With MDIForm1.Toolbar1.Buttons

    .Clear
    
    For i = 1 To UBound(CurrentToolBar)
    
        Typ = CurrentToolBar(i).Type
        If CurrentToolBar(i).Type = 10 Then Typ = 0
        If CurrentToolBar(i).IconID > MDIForm1.ImageList1.ListImages.Count Then CurrentToolBar(i).IconID = MDIForm1.ImageList1.ListImages.Count
        .Add i, , , Typ, CurrentToolBar(i).IconID
        .Item(i).ToolTipText = CurrentToolBar(i).Description
        .Item(i).Tag = CurrentToolBar(i).Tag
    Next i

    If (UBound(CurrentToolBar)) = 0 Then
        MDIForm1.Toolbar1.Visible = False
    Else
        MDIForm1.Toolbar1.Visible = True
    End If

End With

End Sub

Sub FillScriptsMenu()

'fills the script menu with the various available scripts

ClearScriptsMenu

For i = 1 To NumMenuScripts
    
    Load MDIForm1.mnuFuncScriptsIn(i)
    
    nm$ = MenuScripts(i).MenuName
    If Len(MenuScripts(i).AskForQuestion) > 0 Then nm$ = nm$ + "..."
    MDIForm1.mnuFuncScriptsIn(i).Caption = nm$
    MDIForm1.mnuFuncScriptsIn(i).Enabled = True

Next i


If NumMenuScripts = 0 Then MDIForm1.mnuFuncScripts.Visible = False
If NumMenuScripts <> 0 Then MDIForm1.mnuFuncScripts.Visible = True

If ShowPlayers Then Form6.FillCombo

End Sub

Sub ClearScriptsMenu()
On Error Resume Next

Num = MDIForm1.mnuFuncScriptsIn.UBound

Do Until Num = 0
    Unload MDIForm1.mnuFuncScriptsIn(Num)
    Num = MDIForm1.mnuFuncScriptsIn.UBound
Loop


End Sub

Sub StartFuncScript(Index)

'Run
'Dim Params() As String
'ReDim Params(0 To Commands(ButIndex).NumButtons)

If Len(MenuScripts(Index).AskForQuestion) > 0 Then
    q$ = InBox(MenuScripts(Index).AskForQuestion, MenuScripts(Index).MenuName)
End If

'userid:
b$ = "0"

For i = 1 To Form6.ListView1.ListItems.Count
    If Form6.ListView1.ListItems(i).Selected = True Then b$ = Form6.ListView1.ListItems(i).SubItems(2)
Next i

'If a$ = "" Then Exit Sub
'For I = 1 To Form6.ListView1.ListItems.Count
'    If Form6.ListView1.ListItems.Item(I).Text = a$ Then j = I: Exit For
'Next I
'b$ = Trim(Form6.ListView1.ListItems.Item(j).SubItems(2)) 'userid

a$ = ""
a$ = a$ + Chr(251)
a$ = a$ + MenuScripts(Index).ScriptName + Chr(250)
a$ = a$ + b$ + Chr(250)
If q$ <> "" Then a$ = a$ + q$ + Chr(250)
a$ = a$ + Chr(251)

SendPacket "SS", a$

End Sub

Sub SendGameRequestPacket(GameName As String, ToWho As String, GameID)

'compiles and sends the clan info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'compile it

a$ = a$ + Chr(251)
a$ = a$ + GameName + Chr(250)
a$ = a$ + ToWho + Chr(250)
a$ = a$ + Ts(GameID) + Chr(250)
a$ = a$ + Chr(251)

'all set, send it
SendPacket "G1", a$

End Sub


Sub UnPackageGamePacket(p$)

'extracts the crap on who to send it to, and then just passes it on.


f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then IndexFrom = Val(m$)
                If j = 2 Then WhoFrom$ = m$
                If j = 3 Then gmid = Val(m$)
                If j = 4 Then mygmid = Val(m$)
                If j = 5 Then datatopass$ = m$
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

'UnPackageActualGamePacket datatopass$, IndexFrom
For i = 1 To NumGames
    On Error Resume Next
    If GameData(i).GameForm.GameID = mygmid Then
        If gmid <> 0 Then GameData(i).GameForm.RemoteGameID = gmid
        If GameData(i).GameForm.OppIndex = 0 Then GameData(i).GameForm.OppIndex = IndexFrom
        If GameData(i).GameForm.Opponent = "" Then GameData(i).GameForm.Opponent = WhoFrom$
        
        GameData(i).GameForm.GameInterprit datatopass$, CInt(IndexFrom), WhoFrom$
        Exit Sub
    End If
Next i
'Or GameData(I).GameForm.RemoteGameID = gmid
'For I = 1 To NumGames
'    If GameData(I).RemoteGameID = 0 And GameData(I).Index = IndexFrom Then
'        GameData(I).RemoteGameID = gmid
'        GameData(I).GameForm.GameInterprit datatopass$, CInt(IndexFrom), WhoFrom$
'        Exit For
'    End If
'Next I

End Sub


Sub UnPackageGameRequest(p$)

'extracts the crap on who to send it to, and then just passes it on.


f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then GameName$ = m$
                If j = 2 Then WhoFrom$ = m$
                If j = 3 Then Index = Val(m$)
                If j = 4 Then gmid = Val(m$)
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

hn = MessBox("Got a request for " + GameName$ + " from " + WhoFrom$ + " ..." + vbCrLf + "Do you want to join?", vbYesNo, "Request")

If hn = vbYes Then
    
    NumGames = NumGames + 1
    ReDim Preserve GameData(0 To NumGames)
    
    Randomize
    t = Int(Rnd * 30000) + 1

    GameData(NumGames).Index = Index
    GameData(NumGames).Opponent = WhoFrom$
    
    If GameName$ = "Tic Tac Toe v1" Then Set GameData(NumGames).GameForm = New frmGame1
    If GameName$ = "Hangman v1" Then Set GameData(NumGames).GameForm = New frmGame2
    If GameName$ = "Battleship v1" Then Set GameData(NumGames).GameForm = New frmGame3
    If GameName$ = "Private Chat" Then Set GameData(NumGames).GameForm = New frmFreakChat
    If GameName$ = "Private Chat SZ" Then Set GameData(NumGames).GameForm = New frmSZChat
    If GameName$ = "Car Race" Then Set GameData(NumGames).GameForm = New frmCarRace
   

    GameData(NumGames).GameForm.Opponent = WhoFrom$
    GameData(NumGames).GameForm.OppIndex = Index
    GameData(NumGames).GameForm.GameID = t
    GameData(NumGames).GameForm.RemoteGameID = gmid
    GameData(NumGames).GameForm.IsOpponent = True
    GameData(NumGames).GameForm.GameDataNum = NumGames
    GameData(NumGames).GameForm.Show
    GameData(NumGames).GameForm.StartGame

Else

    'tell opponent we wont be playing
    SendGamePacket "N!", "", CInt(Index), 0, gmid

End If

End Sub

Sub RemoveGameData(Num)

For i = Num To NumGames - 1

    Set GameData(i).GameForm = GameData(i + 1).GameForm
    GameData(i).Index = GameData(i + 1).Index
    GameData(i).Opponent = GameData(i + 1).Opponent
    GameData(i).GameForm.GameDataNum = i

Next i
NumGames = NumGames - 1

ReDim Preserve GameData(0 To NumGames)

End Sub

Sub SendActualGamePacket(PacketData As String, ToWho As Integer, GameID As Integer, HisGameID)
'compiles and sends the clan info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'compile it

a$ = a$ + Chr(251)
a$ = a$ + Ts(ToWho) + Chr(250)
a$ = a$ + Ts(GameID) + Chr(250)
a$ = a$ + Ts(HisGameID) + Chr(250)
a$ = a$ + PacketData + Chr(250)
a$ = a$ + Chr(251)

'all set, send it
SendPacket "G2", a$

End Sub

Sub SendGamePacket(Cde As String, Params As String, ToWho As Integer, GameID As Integer, HisGameID)

a$ = Chr(244) + Chr(244) + Chr(244) + Chr(245) + Cde + Chr(245) + Params + Chr(245) + Chr(243) + Chr(243) + Chr(243)

SendActualGamePacket a$, ToWho, GameID, HisGameID

End Sub

Sub ReadAwayMenu()

On Error Resume Next
e = Val(frmConnectUsers.ListView1.SelectedItem.Tag)
If e = 0 Then Exit Sub

b$ = ConnectUsers(e).Name
Msg$ = ConnectUsers(e).AwayMsg



a = ConnectUsers(e).IdleTime

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



Msg$ = "Idle for: " & c$ & "." & vbCrLf & Msg$

'
frmAwayMsg.Label1 = b$ + " is currently in " + GetAwayName(ConnectUsers(e).AwayMode) + " mode."
frmAwayMsg.Text1 = Msg$
frmAwayMsg.Caption = b$ + " is " + GetAwayName(ConnectUsers(e).AwayMode)

'
'a = frmConnectUsers.List1.ListIndex
'If a = -1 Then Exit Sub
'
'e = frmConnectUsers.List1.ItemData(a)
'
'b$ = ConnectUsers(e).Name
'Msg$ = ConnectUsers(e).AwayMsg
'
'frmAwayMsg.Label1 = b$ + " is currently in " + GetAwayName(ConnectUsers(e).AwayMode) + " mode."
'frmAwayMsg.Text1 = Msg$
'frmAwayMsg.Caption = b$ + " is " + GetAwayName(ConnectUsers(e).AwayMode)

End Sub

Sub GamesMenu(Index As Integer)

On Error Resume Next
e = Val(frmConnectUsers.ListView1.SelectedItem.Tag)
If e = 0 Then Exit Sub

b$ = ConnectUsers(e).Name
If Index <> 2 Then
    NumGames = NumGames + 1
    ReDim Preserve GameData(0 To NumGames)
    Randomize
    t = Int(Rnd * 30000) + 1

    If ConnectUsers(e).AwayMode > 0 Then
        MessBox "This person is currently in " + GetAwayName(ConnectUsers(e).AwayMode) + " mode, and cannot answer this request."
        Exit Sub
    End If
End If

If Index = 2 Then
    ' beep
    SendPacket "BE", b$
End If

If Index = 1 Then 'privchat

        Set GameData(NumGames).GameForm = New frmFreakChat
        GameData(NumGames).GameForm.GameID = t
        GameData(NumGames).GameForm.GameDataNum = NumGames
        GameData(NumGames).GameForm.Opponent = b$
        GameData(NumGames).GameForm.Show
        GameData(NumGames).GameForm.lblStatus = "Waiting for " + b$ + "'s response..."
        SendGameRequestPacket "Private Chat", b$, t


End If

If Index = 11 Then 'tic tac toe

        Set GameData(NumGames).GameForm = New frmGame1
        GameData(NumGames).GameForm.GameID = t
        GameData(NumGames).GameForm.GameDataNum = NumGames
        GameData(NumGames).GameForm.Opponent = b$
        GameData(NumGames).GameForm.Show
        GameData(NumGames).GameForm.lblStatus = "Waiting for " + b$ + "'s response..."
        SendGameRequestPacket "Tic Tac Toe v1", b$, t


End If

If Index = 12 Then 'hangman


        Set GameData(NumGames).GameForm = New frmGame2
        GameData(NumGames).GameForm.GameID = t
        GameData(NumGames).GameForm.GameDataNum = NumGames
        GameData(NumGames).GameForm.Opponent = b$
        GameData(NumGames).GameForm.Show
        GameData(NumGames).GameForm.lblStatus = "Waiting for " + b$ + "'s response..."
        SendGameRequestPacket "Hangman v1", b$, t


End If

If Index = 13 Then 'battleship


        Set GameData(NumGames).GameForm = New frmGame3
        GameData(NumGames).GameForm.GameID = t
        GameData(NumGames).GameForm.GameDataNum = NumGames
        GameData(NumGames).GameForm.Opponent = b$
        GameData(NumGames).GameForm.Show
        GameData(NumGames).GameForm.lblStatus = "Waiting for " + b$ + "'s response..."
        SendGameRequestPacket "Battleship v1", b$, t


End If

If Index = 14 Then 'privchat2

        Set GameData(NumGames).GameForm = New frmSZChat
        GameData(NumGames).GameForm.GameID = t
        GameData(NumGames).GameForm.GameDataNum = NumGames
        GameData(NumGames).GameForm.Opponent = b$
        GameData(NumGames).GameForm.Show
        GameData(NumGames).GameForm.lblStatus = "Waiting for " + b$ + "'s response..."
        SendGameRequestPacket "Private Chat SZ", b$, t


End If

If Index = 15 Then 'car race

        Set GameData(NumGames).GameForm = New frmCarRace
        GameData(NumGames).GameForm.GameID = t
        GameData(NumGames).GameForm.GameDataNum = NumGames
        GameData(NumGames).GameForm.Opponent = b$
        GameData(NumGames).GameForm.Show
        GameData(NumGames).GameForm.lblStatus = "Waiting for " + b$ + "'s response..."
        SendGameRequestPacket "Car Race", b$, t


End If

'frmSZChat

End Sub


Sub CheckNewestClient()

a = Val(ReplaceString(NewestClient, ".", ""))
b = Val(Ts(App.Major) + Ts(App.Minor) + Format(Ts(App.Revision), "0000"))

If a > b Then
    
    rq = MessBox("The server has a Server Assistant Client which is newer" + vbCrLf + "then the one that you have. Would you like to download now?" + vbCrLf + "Server version: " + NewestClient, vbYesNo, "New Client Download")
    
    If rq = vbYes Then
    
        'FileLocalPath
        
        With Form1.Dlg1
            
            .FileName = "client.zip"
            .Filter = "Zip Files (*.zip)|*.zip|All Files (*.*)|*.*"
            .DialogTitle = "Select Save Location for New Client"
            .InitDir = App.Path
            .ShowSave
    
            sv$ = .FileName
            
            FileLocalPath = sv$
            
            DownloadingNew = True
            SendPacket "F8", "newclient\\\"
        End With
    End If
End If


End Sub


Public Function ShowOpen(Filter$, Flags&, hwnd) As String
  Dim Buffer$, Result&
  Dim ComDlgOpenFileName As OPENFILENAME
  
    Buffer = String$(128, 0)
   
    With ComDlgOpenFileName
      .lStructSize = Len(ComDlgOpenFileName)
       .hwndOwner = hwnd
      .Flags = Flags
      .nFilterIndex = 1
      .nMaxFile = Len(Buffer)
      .lpstrFile = Buffer
      .lpstrFilter = Filter
    End With
    
    Result = GetOpenFileName(ComDlgOpenFileName)
    
    If Result <> 0 Then
      ShowOpen = Left$(ComDlgOpenFileName.lpstrFile, _
                 InStr(ComDlgOpenFileName.lpstrFile, _
                 Chr$(0)) - 1)
    End If
End Function

Public Function ShowSave(Filter$, Flags&, _
                           hwnd, FileName$) As String
                           
  Dim Buffer$, Result&
  Dim ComDlgOpenFileName As OPENFILENAME
  
    Buffer = FileName & String$(128 - Len(FileName), 0)
    
    With ComDlgOpenFileName
      .lStructSize = Len(ComDlgOpenFileName)
      .hwndOwner = hwnd
      .Flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST
      .nFilterIndex = 1
      .nMaxFile = Len(Buffer)
      .lpstrFile = Buffer
      .lpstrFilter = Filter
    End With

    Result = GetSaveFileName(ComDlgOpenFileName)
    
    If Result <> 0 Then
      ShowSave = Left$(ComDlgOpenFileName.lpstrFile, _
                 InStr(ComDlgOpenFileName.lpstrFile, _
                 Chr$(0)) - 1)
    End If
End Function

Sub UpdateAwayMode()

' send server new away mode

a$ = Chr(251)
a$ = a$ + Ts(MyAwayMode) + Chr(250)
a$ = a$ + MyAwayMsg + Chr(250)
a$ = a$ + Chr(251)

SendPacket "AW", a$

If MyAwayMode > 0 Then
    MDIForm1.Caption = "Server Assistant Client - Copyright 2001 CyberWyre - " + GetAwayName(MyAwayMode)
Else
    MDIForm1.Caption = "Server Assistant Client - Copyright 2001 CyberWyre"
End If


End Sub

Public Function NewWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As WINDOWPOS) As Long
' Size bounds in pixels.


    Dim OrigForm As Object
    
    For k = 1 To UBound(FormList)
    
        If FormList(k).hwnd = hwnd Then
            Set OrigForm = FormList(k).FormObj
            Exit For
        End If
    Next k
        
    If k = 0 Then
        MsgBox "PROBLEM!"
        End
    End If
            
    
    ' Keep the dimensions in bounds.
    If Msg = WM_WINDOWPOSCHANGING Then
    
        If lParam.cx < FormList(k).MinWidth Then lParam.cx = FormList(k).MinWidth
        If lParam.cx > FormList(k).MaxWidth And FormList(k).MaxWidth > 0 Then lParam.cx = FormList(k).MaxWidth
        If lParam.cy < FormList(k).MinHeight Then lParam.cy = FormList(k).MinHeight
        If lParam.cy > FormList(k).MaxHeight And FormList(k).MaxHeight > 0 Then lParam.cy = FormList(k).MaxHeight
    
        On Error GoTo 0
        
        Dim xb As Object
        
        If MDIForm1.mnuSettingsIn(6).Checked Then
            
            For i = Forms.Count - 1 To 1 Step -1
               
                Set xb = Forms(i)
                doit = 0
                For j = 1 To UBound(FormList)
                
                    If FormList(j).hwnd = xb.hwnd Then doit = 1: Exit For
                Next j
                If j = k Then doit = 0
                If doit Then
                
                    snaptop = 0
                    snapleft = 0
                    
                    
                
                    If Abs(lParam.X - ((xb.Left + xb.Width) / Screen.TwipsPerPixelX)) < 10 _
                        And (( _
                        lParam.Y >= xb.Top / Screen.TwipsPerPixelY And lParam.Y <= (xb.Top + xb.Height) / Screen.TwipsPerPixelY _
                        ) Or ( _
                        lParam.cy + lParam.Y >= xb.Top / Screen.TwipsPerPixelY And lParam.cy + lParam.Y <= (xb.Top + xb.Height) / Screen.TwipsPerPixelY _
                        ) Or ( _
                        lParam.Y < xb.Top / Screen.TwipsPerPixelY And lParam.cy + lParam.Y > (xb.Top + xb.Height) / Screen.TwipsPerPixelY _
                        )) Then
                        
                        newone = (xb.Left + xb.Width) / Screen.TwipsPerPixelX
                        diff = lParam.X - newone
                        lParam.X = newone
                        lParam.cx = lParam.cx + diff
                        snaptop = 1
                    End If
                    
                    If Abs((lParam.X + lParam.cx) - (xb.Left / Screen.TwipsPerPixelX)) < 10 _
                        And (( _
                        lParam.Y >= xb.Top / Screen.TwipsPerPixelY And lParam.Y <= (xb.Top + xb.Height) / Screen.TwipsPerPixelY _
                        ) Or ( _
                        lParam.cy + lParam.Y >= xb.Top / Screen.TwipsPerPixelY And lParam.cy + lParam.Y <= (xb.Top + xb.Height) / Screen.TwipsPerPixelY _
                        ) Or ( _
                        lParam.Y < xb.Top / Screen.TwipsPerPixelY And lParam.cy + lParam.Y > (xb.Top + xb.Height) / Screen.TwipsPerPixelY _
                        )) Then
                    
                        
                        newone = Abs((xb.Left / Screen.TwipsPerPixelX) - lParam.X)
                        diff = lParam.cx - newone
                        lParam.cx = newone
                        snaptop = 1
                    End If
                    
                    If Abs(lParam.Y - ((xb.Top + xb.Height) / Screen.TwipsPerPixelY)) < 10 _
                        And (( _
                        lParam.X >= xb.Left / Screen.TwipsPerPixelX And lParam.X <= (xb.Left + xb.Width) / Screen.TwipsPerPixelX _
                        ) Or ( _
                        lParam.cx + lParam.X >= xb.Left / Screen.TwipsPerPixelX And lParam.cx + lParam.X <= (xb.Left + xb.Width) / Screen.TwipsPerPixelX _
                        ) Or ( _
                        lParam.X < xb.Left / Screen.TwipsPerPixelX And lParam.cx + lParam.X > (xb.Left + xb.Width) / Screen.TwipsPerPixelX _
                        )) Then
                        
                        newone = (xb.Top + xb.Height) / Screen.TwipsPerPixelY
                        diff = lParam.Y - newone
                        lParam.Y = newone
                        lParam.cy = lParam.cy + diff
                        snapleft = 1
                    End If
                    
                    If Abs((lParam.Y + lParam.cy) - (xb.Top / Screen.TwipsPerPixelY)) < 10 _
                        And (( _
                        lParam.X >= xb.Left / Screen.TwipsPerPixelX And lParam.X <= (xb.Left + xb.Width) / Screen.TwipsPerPixelX _
                        ) Or ( _
                        lParam.cx + lParam.X >= xb.Left / Screen.TwipsPerPixelX And lParam.cx + lParam.X <= (xb.Left + xb.Width) / Screen.TwipsPerPixelX _
                        ) Or ( _
                        lParam.X < xb.Left / Screen.TwipsPerPixelX And lParam.cx + lParam.X > (xb.Left + xb.Width) / Screen.TwipsPerPixelX _
                        )) Then
                        
                        newone = Abs((xb.Top / Screen.TwipsPerPixelY) - lParam.Y)
                        diff = lParam.cy - newone
                        lParam.cy = newone
                        snapleft = 1
                    End If
                    
                    If Abs(lParam.Y - (xb.Top / Screen.TwipsPerPixelY)) < 10 And snaptop = 1 Then
                        lParam.Y = xb.Top / Screen.TwipsPerPixelY
                    End If
                    If Abs((lParam.Y + lParam.cy) - ((xb.Top + xb.Height) / Screen.TwipsPerPixelY)) < 10 And snaptop = 1 Then
                        lParam.cy = (xb.Top + xb.Height) / Screen.TwipsPerPixelY - lParam.Y
                    End If
                    
                    If Abs(lParam.X - (xb.Left / Screen.TwipsPerPixelX)) < 10 And snapleft = 1 Then
                        lParam.X = xb.Left / Screen.TwipsPerPixelX
                    End If
                    If Abs((lParam.X + lParam.cx) - ((xb.Left + xb.Width) / Screen.TwipsPerPixelX)) < 10 And snapleft = 1 Then
                        lParam.cx = (xb.Left + xb.Width) / Screen.TwipsPerPixelX - lParam.X
                    End If
                    
                End If
            Next
        End If
    End If
    
    ' Continue normal processing. VERY IMPORTANT!
    NewWindowProc = CallWindowProc( _
        OrigForm.OldWindowProc, hwnd, Msg, wParam, _
        lParam)
End Function

Sub UnloadReady()
SaveCommands
Static alreadydone
If alreadydone = 1 Then Exit Sub
alreadydone = 1


SaveSetting "Server Assistant Client", "Window", "winmd", MDIForm1.WindowState
SaveSetting "Server Assistant Client", "Window", "winh", MDIForm1.Height
SaveSetting "Server Assistant Client", "Window", "wint", MDIForm1.Top
SaveSetting "Server Assistant Client", "Window", "winl", MDIForm1.Left
SaveSetting "Server Assistant Client", "Window", "winw", MDIForm1.Width

On Error Resume Next
SaveSetting "Server Assistant Client", "Window", "showchat", ShowChat
SaveSetting "Server Assistant Client", "Window", "showplayers", ShowPlayers
SaveSetting "Server Assistant Client", "Window", "showmap", ShowMap
SaveSetting "Server Assistant Client", "Window", "showusers", ShowUsers

Dim Form As Form
   
For Each Form In Forms
    If Form.Name <> "MDIForm1" And Form.Name <> "Form1" Then
        Unload Form
        Set Form = Nothing
        
    End If
    DoEvents
    a = a + 1
    If a > 50 Then Exit For
Next Form

d$ = App.Path + "\wbdata\*.*"
If Dir(d$) <> "" Then Kill d$

'Dim I As Integer
'For I = Forms.Count - 1 To 1 Step -1
'   Unload Forms(I)
'Next

DoEvents


End Sub
