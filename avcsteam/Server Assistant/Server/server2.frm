VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Assistant - Version Alpha"
   ClientHeight    =   1815
   ClientLeft      =   150
   ClientTop       =   750
   ClientWidth     =   5700
   Icon            =   "server2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer5 
      Interval        =   1000
      Left            =   4380
      Top             =   480
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2940
      Top             =   0
   End
   Begin MSWinsockLib.Winsock TCP2 
      Left            =   4380
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "wwp.icq.com"
      RemotePort      =   80
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   2520
      Picture         =   "server2.frx":548A
      ScaleHeight     =   555
      ScaleWidth      =   615
      TabIndex        =   0
      Top             =   720
      Width           =   675
   End
   Begin MSWinsockLib.Winsock TCP1 
      Index           =   0
      Left            =   1740
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3900
      Top             =   0
   End
   Begin MSWinsockLib.Winsock RconUDP 
      Left            =   1020
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "24.64.165.49"
      RemotePort      =   28000
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   3420
      Top             =   0
   End
   Begin MSWinsockLib.Winsock UDP1 
      Left            =   60
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "209.166.22.5"
      RemotePort      =   27015
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2460
      Top             =   0
   End
   Begin MSWinsockLib.Winsock UDP3 
      Left            =   540
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "24.64.165.49"
      RemotePort      =   27005
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Menu"
      Begin VB.Menu mnuPopIn 
         Caption         =   "Server Assistant"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuPopIn 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuPopIn 
         Caption         =   "&Stop"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'      FILE: server2.frm
'      PURPOSE: Just a hidden form which contains Winsock Controls, timers, the icon,
'      and the menu.
'
'
' ===========================================================================
' ---------------------------------------------------------------------------

Private Sub Form_Unload(Cancel As Integer)

EndProgram

End Sub


Private Sub mnuPopIn_Click(Index As Integer)
MenuClick Index

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

X = X / Screen.TwipsPerPixelX

TrayIcon X

End Sub

Private Sub RconUDP_DataArrival(ByVal bytesTotal As Long)
If DebugMode Then LastCalled = "RconUDP_DataArrival"
On Error GoTo errocc

RconUDP.GetData a$
'a$ = Chr(254) + a$ + Chr(255) + Chr(255) + RconUDP.RemoteHostIP + Chr(255) + Chr(255) + Trim(Str(RconUDP.RemotePort)) + Chr(254)
'Text1 = Text1 + a$ + vbCrLf

IP$ = RconUDP.RemoteHostIP
Prt = RconUDP.RemotePort

'check buffer

k = UBound(UserRCONBuffer)

For i = 0 To k

    If IP$ = UserRCONBuffer(i).IP And Prt = UserRCONBuffer(i).Port Then j = i: Exit For
Next i

If j = 0 Then 'add an entry
    ReDim Preserve UserRCONBuffer(0 To k + 1)
    j = k + 1
End If

UserRCONBuffer(j).Command = UserRCONBuffer(j).Command + a$
UserRCONBuffer(j).IP = IP$
UserRCONBuffer(j).Port = Prt

Exit Sub
errocc:
ErrorReport Err.Number, Err.Description + ", " + Err.Source

End Sub

Private Sub TCP1_SendComplete(Index As Integer)

'SEE if he need to be kicked
For i = 1 To NumConnectUsers
    If ConnectUsers(i).Index = Index Then j = i: Exit For
Next i

If j > 0 Then
    If ConnectUsers(j).RemoveMe = True Then TCP1_Close Index
End If



End Sub

Private Sub TCP2_Connect()

Dim strCommand As String
Dim strWebpage As String
strWebpage = ICQURL
strCommand = "GET " + strWebpage + " HTTP/1.0" + vbCrLf
strCommand = strCommand + "Accept: *.*" + vbCrLf
strCommand = strCommand + "Accept: text/html" + vbCrLf
strCommand = strCommand + vbCrLf

TCP2.SendData strCommand

End Sub

Private Sub TCP2_DataArrival(ByVal bytesTotal As Long)

TCP2.Close
SendingICQ = False

End Sub

Private Sub Timer1_Timer()
TimerVR = TimerVR - 1
End Sub

Public Sub TimerGo()

End Sub



Private Sub Timer2_Timer()
Timer2.Enabled = False

On Error GoTo errocc
If DebugMode Then LastCalled = "Timer2_Timer"

LastTimer2 = Now
LastTimer2What = "Location 1"

If Len(LogBuffer) > 0 Then
    HandleEntry
End If

LastTimer2 = Now
LastTimer2What = "Location 2"


If Len(LastRCON) > 0 Then
    HandleLastRcon
End If

LastTimer2 = Now
LastTimer2What = "Location 3"

If DebugMode Then LastCalled = "Timer2_Timer - After HandleLast"

LastTimer2 = Now
LastTimer2What = "Location 4"

k = UBound(UserRCONBuffer)
If k > 0 Then
    DoEvents
    HandleUserRcon False, "", "", 0, 0
End If

LastTimer2 = Now
LastTimer2What = "Location 5"

Timer2.Enabled = True
Exit Sub
errocc:
Timer2.Enabled = True
ErrorReport Err.Number, Err.Description + ", " + Err.Source

End Sub

Private Sub Timer3_Timer()
If DebugMode Then LastCalled = "Timer3_Timer"

Dim SecElap As Variant

DoEvents

'If Len(Text1) > 4000 Then Text1 = Right(Text1, 3990): Text1.SelStart = Len(Text1)
'If NumPlayers > 20 Then NumPlayers = 20

On Error GoTo errocc

If DebugMode Then LastCalled = "Timer3_Timer - Part 1"
If General.MaxTime > 0 And General.MaxMsg > 0 And Vars.ClanBattle = False Then
    'spam checking
    
    Vars.TimeCounter = Vars.TimeCounter + 1
    If Vars.TimeCounter > 30000 Then Vars.TimeCounter = 0
        
    Vars.TimeCounter2 = Vars.TimeCounter2 + 1
    If Vars.TimeCounter2 > 30000 Then Vars.TimeCounter2 = 0
        
    If Vars.TimeCounter >= General.MaxTime Then
        Vars.TimeCounter = 0
        
        'reset talk totals
        For i = 1 To NumPlayers
            Players(i).MessInMin = 0
        Next i
    End If

    If Vars.TimeCounter2 >= Val(General.SameSpamTime) Then
        Vars.TimeCounter2 = 0
        'reset talk totals
        For i = 1 To NumPlayers
            For kk = 1 To 10
                Players(i).LastMsgs(kk).Text = ""
                Players(i).LastMsgs(kk).When = 0
            Next kk
        Next i
    End If

End If

'Clear timer if time
If DebugMode Then LastCalled = "Timer3_Timer - Part 2"

AdminSpeechBuffer = AdminSpeechBuffer + 1
If AdminSpeechBuffer > General.MaxSpeechTime And General.MaxSpeechTime > 0 Then
    AdminSpeechBuffer = 0
    For i = 1 To 20
        LastSpeech(i).UserID = 0
    Next i
End If
If AdminSpeechBuffer > 30000 Then AdminSpeechBuffer = 0

If AdminICQTime > 0 Then
    AdminICQTime = AdminICQTime - 1
    If AdminICQTime = 0 Then AdminICQId = 0
End If

If DebugMode Then LastCalled = "Timer3_Timer - Part 3"
'Handle TimerVars
If NumTimerVars > 0 Then

again123:
    For i = 1 To NumTimerVars
        c = Val(GetVar(TimerVars(i), BanScriptData)) - 1
        SetVar TimerVars(i), Ts(c), BanScriptData
    
        If c <= 0 Then 'remove it
        
            For j = i To NumTimerVars - 1
                TimerVars(i) = TimerVars(i + 1)
            Next j
            NumTimerVars = NumTimerVars - 1
            ReDim Preserve TimerVars(0 To NumTimerVars)
            GoTo again123
        End If
    Next i
End If
'Check EVENTS.

If DebugMode Then LastCalled = "Timer3_Timer - Part 4"
EventTimer = EventTimer - 1
If EventTimer <= 0 Then
    EventTimer = 10
    EventHandler
End If

If DebugMode Then LastCalled = "Timer3_Timer - Part 5"
LogWatchTimer = LogWatchTimer + 1
If LogWatchTimer > 180 Then
    're-initiate rcon thing
    LogWatchTimer = 0
    
    'scan for AMA files
    
    ScanForAMAFiles
    
    If ProgramON Then
        If TempStop = False Then
            SendActualRcon "logaddress"
            StartLogWatch
        End If
        AskClanBattle 'check to ensure this isnt a clan battle
        CheckForDLL
        're-check the user list with whats on the server
        SendRCONCommand "users"
        SendRCONCommand "status"
        If LastCrashCall <> 3 Then
            LastCrashCall = 1
            CrashTimer = 20
        End If
    End If
End If

If DebugMode Then LastCalled = "Timer3_Timer - Part 5"
If Vars.MapTimeLeft > 0 Then Vars.MapTimeLeft = Vars.MapTimeLeft - 1

If LastCrashCall > 0 And LastCrashCall < 3 Then
    'deduct timer
    If CrashTimer > 0 Then CrashTimer = CrashTimer - 1
    If CrashTimer = 0 Then CheckIfServerCrashed
End If

If NextTime = 0 Then
    If Vars.ClanBattle = False Then

        If Vars.MapTimeLeft = 960 Then Vars.MapTimeLeft = 959: AnnTime = False: AskTimeRemaining: NextTime = 10
        If Vars.MapTimeLeft = 900 Then Vars.MapTimeLeft = 899: AnnTime = True: AskTimeRemaining: NextTime = 10
        If Vars.MapTimeLeft = 30 Then Vars.MapTimeLeft = 29: AnnTime = True: AskTimeRemaining: NextTime = 10
        If Vars.MapTimeLeft = 60 Then Vars.MapTimeLeft = 59: AnnTime = True: AskTimeRemaining: NextTime = 10
        If Vars.MapTimeLeft = 10 Then Vars.MapTimeLeft = 9: AnnTime = True: AskTimeRemaining: NextTime = 10
    End If
End If

If NextTime > 0 Then NextTime = NextTime - 1

If General.MapVoteMode = "2" And Vars.MapTimeLeft < 3 And Vars.VotedMap <> "" Then
    'do it manually
    SendRCONCommand General.MapChangeMode + " " + Vars.VotedMap
    Vars.VotedMap = ""
End If

If TimeRemainAskCount > 0 Then
    TimeRemainAskCount = TimeRemainAskCount - 1
    If TimeRemainAskCount <= 0 Then
        AnnTime = False
        AskTimeRemaining
        If General.MapVoteMode = "1" Then SendRCONCommand "mapcyclefile " + Vars.MapCycleFile
    End If
End If

If DebugMode Then LastCalled = "Timer3_Timer - Part 6"
'map time remaining calculator
MapCounter = MapCounter + 1
If MapCounter > 60 Then
    MapCounter = 0
    LogSortTimer = LogSortTimer + 1
    If LogSortTimer > 30000 Then LogSortTimer = 0
    If Vars.MapTimeElapsed < 2000 Then Vars.MapTimeElapsed = Vars.MapTimeElapsed + 1
    If Vars.MapTimeElapsed = 1 Then AnnTime = False: AskTimeRemaining
        
    'update the lists on the clients
    SendUpdate
End If
If MapCounter = 30 Then SendUpdate

'If Vars.MapTimeLeft = 0 And TimeRemainAskCount = 0 Then AskTimeRemaining: TimeRemainAskCount = 15: AnnTime = False: Vars.MapTimeLeft = -1
If DebugMode Then LastCalled = "Timer3_Timer - Part 7"
'start an automatic mapvote

If General.MapVoteStartTime = 0 Then General.MapVoteStartTime = 300

 
Dim NewScriptData As typScriptData

If General.MapVoteStartTimeMode = 0 Then
    If Vars.MapTimeLeft < General.MapVoteStartTime And Vars.MapTimeLeft > General.MapVoteStartTime - 20 Then
        NewScriptData.TimeStarted = Timer
        NewScriptData.UserName = "<SERVER>"

        StartAutoMapVote NewScriptData
    End If
End If

If General.MapVoteStartTimeMode = 1 Then
    SecElap = (Vars.MapTimeElapsed * 60) + MapCounter
    If SecElap > General.MapVoteStartTime _
    And SecElap < General.MapVoteStartTime + 60 Then
        NewScriptData.TimeStarted = Timer
        NewScriptData.UserName = "<SERVER>"
        StartAutoMapVote NewScriptData
    End If
End If

DoEvents

If DebugMode Then LastCalled = "Timer3_Timer - Part 8"
'see if its time to set the map cycle file
If General.MapVoteMode = "1" Then
    If Vars.MapCycleFile = "" Then ' get the current file
        SendRCONCommand "mapcyclefile"
        Vars.MapCycleFile = "-1"
    ElseIf Vars.MapTimeLeft < 60 And Vars.TempFileMade = False And Vars.VotedMap <> "" And Vars.ClanBattle = False Then
        'Write file
        
        If CheckForFile(Server.BothPath + "\" + Vars.TempMapCycleFile) Then Kill Server.BothPath + "\" + Vars.TempMapCycleFile
        
        Randomize
        
        Vars.TempMapCycleFile = "assistmap" + Ts(Int(Rnd * 100) + 1) + ".txt"
        
        If CheckForFile(Server.BothPath + "\assistmap*.txt") Then Kill Server.BothPath + "\assistmap*.txt"
        
        h = FreeFile
        Open Server.BothPath + "\" + Vars.TempMapCycleFile For Append As h
            Print #h, Vars.VotedMap
        Close h
          
        Vars.TempFileMade = True
        SendRCONCommand "mapcyclefile " + Vars.TempMapCycleFile
    End If
End If



If DebugMode Then LastCalled = "Timer3_Timer - Part 9"
'sort logs every 5 hours
If LogSortTimer > 300 Then
    LogSortTimer = 0
    If CheckBit2(General.Flags, 0) Then SortLogs
    If CheckBit2(General.Flags, 5) Then RemoveOldMessages
    
    'Remove any bad temporary real players
    ScanForOldReals
End If

If DebugMode Then LastCalled = "Timer3_Timer - Part 10"
If MapVoteTimer > 0 Then
    MapVoteTimer = MapVoteTimer - 1
    If MapVoteTimer Mod 15 = 0 Then  'time for a status report!
        a$ = TotalVotes(0)
        SendRCONCommand "say " + a$
    End If
    If MapVoteTimer = 0 Then
        SendRCONCommand "say " + TotalVotes(0)
        B$ = TotalVotes(1)
        
        'finish the vote
        'FinishVote b$
    End If
    
    If MapVoteTimer = 119 Then SendRCONCommand "say Map Vote: 2 min left!"
    If MapVoteTimer = 59 Then SendRCONCommand "say Map Vote: 1 min left!"
    If MapVoteTimer = 29 Then SendRCONCommand "say Map Vote: 30 sec left!"
End If


If DebugMode Then LastCalled = "Timer3_Timer - Part 11"
If KickVoteTimer > 0 Then
    KickVoteTimer = KickVoteTimer - 1
    If KickVoteTimer = 0 Then
        TotalKickVotes
    End If
End If

If ChooseVoteTimer > 0 Then
    ChooseVoteTimer = ChooseVoteTimer - 1
    If ChooseVoteTimer = 0 Then
        TotalChooseVotes
    End If
End If

If DebugMode Then LastCalled = "Timer3_Timer - Done"

Exit Sub
errocc:
ErrorReport Err.Number, Err.Description + ", " + Err.Source

End Sub


Private Sub Timer4_Timer()
If DebugMode Then LastCalled = "Timer4_Timer"

On Error Resume Next

'check to ensure RconMonitor's get timed out after 1 second
If NumRconMonitors > 0 Then
    Do
        If Timer - RconMonitors(1).TimeSent >= 1 Then RemoveFirstMonitor
        If NumRconMonitors = 0 Then
            ad = 1
        Else
            If Timer - RconMonitors(1).TimeSent < 1 Then ad = 1
        End If
    Loop Until ad = 1
    
End If

End Sub

Private Sub Timer5_Timer()

CheckHLDS

End Sub

Private Sub UDP1_DataArrival(ByVal bytesTotal As Long)

    On Error Resume Next
    
    tc = UDP1.State
    
'   If tc = 1 Then Exit Sub
       
           
        
    UDP1.GetData a$
    
    B$ = Chr(255) + Chr(255) + Chr(255) + Chr(255) + "challenge rcon"
    
    If LeftR(a$, Len(B$)) = B$ Then
            
        ' get challenge num
        'ÿÿÿÿchallenge rcon 3135120253
        
        e = InStrRev(a$, " ")
        f = InStr(1, a$, Chr(10))
        If e > 0 And f > 0 Then
            
            
            ch$ = Mid(a$, e + 1, f - e - 1)
    
            ChallengeNum = ch$
            AfterChallenge
    
        End If
    
    Else
        LogBuffer = LogBuffer + a$
    End If
    
End Sub

Private Sub TCP1_Close(Index As Integer)
On Error Resume Next
TCP1(Index).Close
TCPCreated(Index) = False
Unload TCP1(Index)

'remove the user

For i = 1 To NumConnectUsers
    If ConnectUsers(i).Index = Index Then
        AddToLogFile "LOGOUT: " + ConnectUsers(i).Name + " logged out."
        ExecFunctionScript "spec_adminlogout", 1, ConnectUsers(i).Name
        RemoveUser i
        Exit For
    End If
Next i
UpdateUsersList

End Sub

Private Sub TCP1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If DebugMode Then LastCalled = "TCP1_ConnectionRequest"
On Error GoTo errclr


Dim IP As String
If Index <> 0 Then Exit Sub

'First, see what we can remove.

doagain:

For i = 1 To NumConnectUsers
    j = ConnectUsers(i).Index
    If TCPCreated(j) = True Then
        'See if this is still active.
        If TCP1(j).State <> sckConnected Then
            'Its There, but not connected... remove it!
            Unload TCP1(j)
            TCPCreated(j) = False
            
            'This means the connectuser is also using a defunct tcp... remove the user
            RemoveUser i
            'restart
            GoTo doagain
        End If
    End If
Next i

'now ensure we dont have 2 users with the same index

doagain2:
For i = 1 To NumConnectUsers
    j = ConnectUsers(i).Index
    
    For k = 1 To NumConnectUsers
        If k <> i Then
            If ConnectUsers(k).Index = j Then
                RemoveUser k
                GoTo doagain2
            End If
        End If
    Next k
Next i


IP = TCP1(Index).RemoteHostIP

j = 0
For i = 1 To TCP1.count - 1

    If TCPCreated(i) = True Then
        If j = 0 And TCP1(i).State <> sckConnected Then
            If TCP1(i).State <> sckClosed Then TCP1(i).Close
            j = i
        End If
    Else
        j = i
        Exit For
    End If
Next i

If j = 0 Then
    Index = TCP1.count - 1

    If TCPCreated(Index + 1) = False Then
        Load TCP1(Index + 1)
        TCPCreated(Index + 1) = True
        
    End If
    TCP1(Index + 1).LocalPort = 0
    
Else
    Index = j - 1
    
    If TCPCreated(Index + 1) = False Then
        Load TCP1(Index + 1)
        TCPCreated(Index + 1) = True
    End If
    TCP1(Index + 1).LocalPort = 0
End If

TCP1(Index + 1).LocalPort = 0
TCP1(Index + 1).RemotePort = TCP1(Index).RemotePort
TCP1(Index + 1).Accept requestID
RecData(Index + 1) = ""

'add the user to the list

again:
'now add im

NumConnectUsers = NumConnectUsers + 1

ConnectUsers(NumConnectUsers).IP = IP$
ConnectUsers(NumConnectUsers).Port = TCP1(Index).RemotePort
ConnectUsers(NumConnectUsers).Index = Index + 1
ConnectUsers(NumConnectUsers).Name = "Retrieving..."
ConnectUsers(NumConnectUsers).FileRecieveMode = False
ConnectUsers(NumConnectUsers).SendingFile = False
ConnectUsers(NumConnectUsers).FileStop = False
ConnectUsers(NumConnectUsers).FileNum = 0
ConnectUsers(NumConnectUsers).HiddenMode = False
ConnectUsers(NumConnectUsers).RemoveMe = False
ConnectUsers(NumConnectUsers).AwayMode = 0
ConnectUsers(NumConnectUsers).AwayMsg = ""
ConnectUsers(NumConnectUsers).EncryptedMode = False
ConnectUsers(NumConnectUsers).FileSavePath = ""
ConnectUsers(NumConnectUsers).IdleTime = 0

'UpdateUsersList

Exit Sub
errclr:
ErrorReport Err.Number, Err.Description + ", " + Err.Source

End Sub

Private Sub TCP1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
If DebugMode Then LastCalled = "TCP1_DataArrival"
On Error GoTo errocc
'format

'(254)(254)(254)(255)[CODE](255)[NAME](255)[PASSWORD](255)[PARAMS](255)(253)(253)(253)
TCP1(Index).GetData a$

For i = 1 To NumConnectUsers
    If ConnectUsers(i).Index = Index Then j = i: Exit For
Next i
  


st = Timer

If j > 0 Then

    startstr$ = Chr(254) + Chr(254) + Chr(254)
    endstr$ = Chr(253) + Chr(253) + Chr(253)
    
    RecData(Index) = RecData(Index) + a$
    
    Do
        e = InStr(1, RecData(Index), startstr$)
        ee = InStr(e + 1, RecData(Index), endstr$)
        
        pss$ = ConnectUsers(j).PassWord
        
        If e And ee > e Then 'there is a whole line
        
            If e > 1 Then 'not at beginning
                RecData(Index) = Right(RecData(Index), Len(RecData(Index)) - e + 1)
                e = InStr(1, RecData(Index), startstr$)
            End If
        
            'extract
            f = InStr(e + 1, RecData(Index), endstr$)
            
            If e > 0 And f > e And f > 0 Then
                a$ = Mid(RecData(Index), e, f - e + 3)
                        
                If Len(RecData(Index)) - Len(a$) > 0 Then
                    RecData(Index) = Right(RecData(Index), Len(RecData(Index)) - Len(a$))
                Else
                    RecData(Index) = ""
                End If
                
                If ConnectUsers(j).EncryptedMode = True Then
                    
                    a$ = Right(a$, Len(a$) - 3)
                    a$ = Left(a$, Len(a$) - 3)
                    
                    a$ = Encrypt(a$, pss$)
                    Debug.Print something
                End If
            
                Interprit a$, Index
                        
            End If
        End If

nxtpacket:
         
    Loop Until e = 0 Or ee = 0

End If

'tm = Timer - st
'
'h = FreeFile
'Open App.Path + "\recvlog.log" For Append As h
'    Print #h, Date$ + " " + Time$ + " : " + p$ + ", took: " + Ts(Round(tm, 10))
'Close h

Exit Sub

errocc:
ErrorReport Err.Number, Err.Description + ", " + Err.Source

End Sub

Private Sub RemoveUser(Num)

If ConnectUsers(Num).FileNum > 0 Then Close ConnectUsers(Num).FileNum


NumConnectUsers = NumConnectUsers - 1
For i = Num To NumConnectUsers
    ConnectUsers(i).IP = ConnectUsers(i + 1).IP
    ConnectUsers(i).Index = ConnectUsers(i + 1).Index
    ConnectUsers(i).Name = ConnectUsers(i + 1).Name
    ConnectUsers(i).PassWord = ConnectUsers(i + 1).PassWord
    ConnectUsers(i).LogLevel = ConnectUsers(i + 1).LogLevel
    ConnectUsers(i).UserNum = ConnectUsers(i + 1).UserNum
    ConnectUsers(i).SendingFile = ConnectUsers(i + 1).SendingFile
    'ConnectUsers(I).FileRecieveMode = ConnectUsers(I + 1).FileRecieveMode
    ConnectUsers(i).FileSavePath = ConnectUsers(i + 1).FileSavePath
    ConnectUsers(i).FileSize = ConnectUsers(i + 1).FileSize
    ConnectUsers(i).FileNum = ConnectUsers(i + 1).FileNum
    ConnectUsers(i).FileStop = ConnectUsers(i + 1).FileStop
    ConnectUsers(i).Version = ConnectUsers(i + 1).Version
    ConnectUsers(i).BytesTransferred = ConnectUsers(i + 1).BytesTransferred
    ConnectUsers(i).HiddenMode = ConnectUsers(i + 1).HiddenMode
    ConnectUsers(i).Port = ConnectUsers(i + 1).Port
    ConnectUsers(i).RemoveMe = ConnectUsers(i + 1).RemoveMe
    ConnectUsers(i).AwayMode = ConnectUsers(i + 1).AwayMode
    ConnectUsers(i).AwayMsg = ConnectUsers(i + 1).AwayMsg
    
Next i

End Sub


Private Sub UDP1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    CancelDisplay = True
End Sub
