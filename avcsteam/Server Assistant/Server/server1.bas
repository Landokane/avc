Attribute VB_Name = "a_server"
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
'      FILE: server1.bas
'      PURPOSE:
'      This file is the main code file for Server Assistant Server. It
'      contains most of the code which makes SA work, including the
'      AvScript engine.
'
'
' ===========================================================================
' ---------------------------------------------------------------------------

Function StartHLDS(FileApp As String, currDir As String, cmdArgs As String)
    
    Dim sei As SHELLEXECUTEINFO
    Dim retval As Long
    
    With sei
        .cbSize = Len(sei)
        ' Use the optional hProcess element of the structure.
        .fMask = SEE_MASK_NOCLOSEPROCESS
        .hwnd = Form1.hwnd
        .lpVerb = "open"
        .lpFile = FileApp
        .lpParameters = cmdArgs
        .lpDirectory = currDir
        .nShow = SW_SHOWNORMAL
    End With
    
    ' Open the file using its associated program.
    retval = ShellExecuteEx(sei)
    
    If retval = 0 Then
    Else
        HLDSProcID = sei.hProcess
    End If
End Function


Sub CheckHLDS()
    
    If Not ServerStart.UseFeature Then Exit Sub
    
    'See the status
    'SupposedToBeRunning
    
    Dim lpExitCode As Long
    
    If HLDSProcID > 0 Then
        hnd = GetExitCodeProcess(HLDSProcID, lpExitCode)
        If lpExitCode = 259 Then
            HLDSStatus = True
        Else
            HLDSStatus = False
            CloseHandle HLDSProcID
        End If
    Else
        HLDSStatus = False
    End If
    
  
    If HLDSStatus <> SupposedToBeRunning And (ServerStart.AutoRestart Or ManualStart) Then
    
        ManualStart = False
        
        ' Do what we should be.
        
        If SupposedToBeRunning = True Then
            
            StartHLDS ServerStart.HLDSPath, ServerStart.HLDSDir, ServerStart.CommandLine
        Else
            'Stop it here.
            
            TerminateProcess HLDSProcID, lpExitCode
            CloseHandle HLDSProcID
            HLDSProcID = 0
        End If
    End If


End Sub

Sub SaveTeleFile()
If DebugMode Then LastCalled = "SaveTeleFile"

'saves the file

If Vars.Map = "" Then Exit Sub

f$ = Server.BothPath + "\Assistant"
If Dir(f$, vbDirectory) = "" Then MkDir f$

f$ = Server.BothPath + "\Assistant\MapTele"
If Dir(f$, vbDirectory) = "" Then MkDir f$
f$ = f$ + "\" + Vars.Map + ".dat"

If CheckForFile(f$) Then Kill f$

'Save the data
h = FreeFile
Close h
Open f$ For Binary As h
    Put #h, , NumTele
    Put #h, , Tele
Close h


End Sub

Sub SaveMapProcess()
If DebugMode Then LastCalled = "SaveMapProcess"

'saves the file

f$ = Server.BothPath + "\Assistant"
If Dir(f$, vbDirectory) = "" Then MkDir f$

f$ = Server.BothPath + "\Assistant\Data"
If Dir(f$, vbDirectory) = "" Then MkDir f$
f$ = f$ + "\mapprocess.dat"

If CheckForFile(f$) Then Kill f$

'Save the data
h = FreeFile
Close h
Open f$ For Binary As h
    Put #h, , NumMapProcess
    Put #h, , MapProcess
Close h


End Sub

Sub LoadMapProcess()
If DebugMode Then LastCalled = "LoadMapProcess"

'loads the file

f$ = Server.BothPath + "\Assistant"
If Dir(f$, vbDirectory) = "" Then MkDir f$
f$ = Server.BothPath + "\Assistant\Data"
If Dir(f$, vbDirectory) = "" Then MkDir f$

f$ = f$ + "\mapprocess.dat"

If CheckForFile(f$) Then
    'Load the data
    h = FreeFile
Close h
    Open f$ For Binary As h
        Get #h, , NumMapProcess
        ReDim MapProcess(0 To NumMapProcess)
        Get #h, , MapProcess
    Close h
End If

End Sub

Sub AddToMapProcess(MapName As String)
If DebugMode Then LastCalled = "AddtoMapProcess - " + MapName

Static LastMap As String

If MapName = LastMap Then Exit Sub
LastMap = MapName

For i = 1 To NumMapProcess
    If LCase(MapProcess(i).MapName) = LCase(MapName) Then j = i: Exit For
Next i

If j = 0 Then
    'add
    NumMapProcess = NumMapProcess + 1
    ReDim Preserve MapProcess(0 To NumMapProcess)
    j = NumMapProcess
End If
    


sc = GetSec(MapProcess(j).LastTimePlayed)

MapProcess(j).LastTimePlayed = Now
MapProcess(j).MapName = MapName

'make sure enough time has passed
If sc > 120 Then ' 2 min must pass
    MapProcess(j).TimesPlayed = MapProcess(j).TimesPlayed + 1
End If

SaveMapProcess

End Sub

Sub LoadTeleFile()
If DebugMode Then LastCalled = "LoadTeleFile"

'loads the file
If Vars.Map = "" Then Exit Sub

f$ = Server.BothPath + "\Assistant"
If Dir(f$, vbDirectory) = "" Then MkDir f$
f$ = Server.BothPath + "\Assistant\MapTele"
If Dir(f$, vbDirectory) = "" Then MkDir f$

f$ = f$ + "\" + Vars.Map + ".dat"

If CheckForFile(f$) Then
    'Load the data
    h = FreeFile
Close h
    Open f$ For Binary As h
        Get #h, , NumTele
        ReDim Tele(0 To NumTele)
        Get #h, , Tele
    Close h
End If

End Sub

Sub SaveMapDataFile()
If DebugMode Then LastCalled = "SaveMapDataFile"

'saves the file

If Vars.Map = "" Then Exit Sub

f$ = Server.BothPath + "\Assistant" ' \MapData"
If Dir(f$, vbDirectory) = "" Then MkDir f$
f$ = Server.BothPath + "\Assistant\MapData"
If Dir(f$, vbDirectory) = "" Then MkDir f$

f$ = f$ + "\" + Vars.Map + ".dat"

If CheckForFile(f$) Then Kill f$

'Save the data
h = FreeFile
Close h
Open f$ For Binary As h
    Put #h, , MapArray
Close h


End Sub

Sub LoadMapDataFile()
If DebugMode Then LastCalled = "LoadMapDataFile"

For X = 0 To 64
    For Y = 0 To 64
        MapArray(X, Y) = 0
    Next Y
Next X

'loads the file
If Vars.Map = "" Then Exit Sub

f$ = Server.BothPath + "\Assistant\MapData"
If Dir(f$, vbDirectory) = "" Then MkDir f$
f$ = f$ + "\" + Vars.Map + ".dat"

If CheckForFile(f$) Then
    'Load the data
    h = FreeFile
Close h
    Open f$ For Binary As h
        Get #h, , MapArray
    Close h
End If

End Sub

Sub CreateIcon()
If DebugMode Then LastCalled = "CreateIcon"
    'adds the tray icon
    Dim Tic As NOTIFYICONDATA
    Tic.cbSize = Len(Tic)
    Tic.hwnd = Form1.Picture1.hwnd
    Tic.UID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = Form1.Picture1.Picture
    Tic.szTip = "Half-Life Server Assistant is running - " + Server.HostName & Chr$(0)
    erg = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub

Sub DeleteIcon()
If DebugMode Then LastCalled = "DeleteIcon"
    'removes tray icon
    Dim Tic As NOTIFYICONDATA
    Tic.cbSize = Len(Tic)
    Tic.hwnd = Form1.Picture1.hwnd
    Tic.UID = 1&
    erg = Shell_NotifyIcon(NIM_DELETE, Tic)
End Sub

Sub Main()
DebugMode = True
If DebugMode Then LastCalled = "Main"

'main startup sub


AddToLogFile "SERVER: Starting up..."


Dim expdate As Date
expdate = #12/30/2001#
'
'If Now > expdate Then
'    AddToLogFile "SERVER: Software expired. Contact Avatar-X for a newer version. Exiting..."
'    End
'    Exit Sub
'End If


ProgramON = True
 
'ReDim Commands(1 To 200)
ReDim UserRCONBuffer(0 To 0)

On Error GoTo errocc

ReDim WaitingCommands(0 To 0)

Defaults
DataFile = App.Path + "\server.dat"
DataFile2 = App.Path + "\assistant.dat"
DataFile3 = App.Path + "\speech.dat"
DataFileNew = App.Path + "\astdata.dat"
DataFileOld = App.Path + "\assist.dat"
DataFileOlder = App.Path + "\servdata.dat"

Data(1) = App.Path + "\Assistant\Data\data1.dat"
Data(2) = App.Path + "\Assistant\Data\data2.dat"
Data(3) = App.Path + "\Assistant\Data\data3.dat"
Data(4) = App.Path + "\Assistant\Data\data4.dat"
Data(5) = App.Path + "\Assistant\Data\data5.dat"
Data(6) = App.Path + "\Assistant\Data\data6.dat"
Data(7) = App.Path + "\Assistant\Data\data7.dat"
Data(8) = App.Path + "\Assistant\Data\data8.dat"
Data(9) = App.Path + "\Assistant\Data\data9.dat"
Data(10) = App.Path + "\Assistant\Data\data10.dat"
Data(11) = App.Path + "\Assistant\Data\data11.dat"
Data(12) = App.Path + "\Assistant\Data\data12.dat"
Data(13) = App.Path + "\Assistant\Data\data13.dat"
Data(14) = App.Path + "\Assistant\Data\data14.dat"
Data(15) = App.Path + "\Assistant\Data\scripts.dat"
Data(16) = App.Path + "\Assistant\Data\badword.dat"
Data(17) = App.Path + "\Assistant\Data\adminbmp.dat"
Data(18) = App.Path + "\Assistant\Data\stopstart.dat"

f$ = App.Path + "\Assistant"
If Dir(f$, vbDirectory) = "" Then MkDir f$
f$ = App.Path + "\Assistant\Data"
If Dir(f$, vbDirectory) = "" Then MkDir f$

Vars.TempMapCycleFile = "assistmap.txt"

t = LoadCommands

If NumUsers = 0 Then

    NumUsers = 1
    Users(1).Allowed = "all"
    Users(1).Flags = CDec((2 ^ 36) - 1 - (2 ^ 35))
    Users(1).Name = "Admin"
    Users(1).PassWord = "Admin"
End If

Server.BothPath = Server.HLPath + "\" + Server.GamePath
LastIf = -1

Reload = False

Load Form1

SetPorts
'start the logging proccess on the server

'GetChallenge
'DoAfterChallenge = True

StartLogWatch
AskClanBattle
CreateIcon
CheckForDLL
AskTimeRemaining
ServerCrash
'SendRCONCommand "status"

Alpha(1) = "4"
Alpha(2) = "8"
Alpha(3) = "("
Alpha(4) = "|)"
Alpha(5) = "3"
Alpha(6) = "|="
Alpha(7) = "6"
Alpha(8) = "|-|"
Alpha(9) = "1"
Alpha(10) = "J"
Alpha(11) = "|<"
Alpha(12) = "1"
Alpha(13) = "/\/\"
Alpha(14) = "|\|"
Alpha(15) = "0"
Alpha(16) = "P"
Alpha(17) = "Q"
Alpha(18) = "|2"
Alpha(19) = "5"
Alpha(20) = "7"
Alpha(21) = "U"
Alpha(22) = "(_)"
Alpha(23) = "\/\/"
Alpha(24) = "><"
Alpha(25) = "Y"
Alpha(26) = "Z"



LoadMapProcess

'Server.GameMode = 1
If LCase(Server.GamePath) = "cstrike" Then Server.GameMode = 2

AddToLogFile "SERVER: Started up!"

'Do
    'MainLoop
    'DoEvents
'Loop

Exit Sub
'error occurence
errocc:
ErrorReport Err.Number, Err.Description + ", " + Err.Source
End Sub

Function DeLeet(OrigText As String) As String
If DebugMode Then LastCalled = "DeLeet"

'De-Leets a string.

Dim nw As String
nw = UCase(OrigText)

For i = 1 To 26
    nw = Replace(nw, Alpha(i), Chr(i + 64))
Next i

DeLeet = nw

End Function


Function DeSpace(OrigText As String) As String
If DebugMode Then LastCalled = "DeSpace"

'De-Leets a string.


For i = 1 To Len(OrigText)
    a$ = Mid(OrigText, i, 1)
    B = Asc(a$)
    If (B >= Asc("A") And B <= Asc("Z")) Or (B >= Asc("a") And B <= Asc("z")) Or (B >= 48 And B <= 57) Then
        c$ = c$ + a$
    End If
Next i

DeSpace = c$

End Function


Sub test123()

'For I = 1 To NumRealPlayers
'
'    If RealPlayers(I).LastTime > 0 Then j = j + 1
'
'Next I
'
'
'Debug.Print j

ReDim MapProcess(0 To 5)

MapProcess(1).MapName = "2fort"
MapProcess(1).TimesPlayed = 6

MapProcess(2).MapName = "well"
MapProcess(2).TimesPlayed = 18

MapProcess(3).MapName = "rock2"
MapProcess(3).TimesPlayed = 13


MapProcess(4).MapName = "hunted"
MapProcess(4).TimesPlayed = 3


MapProcess(5).MapName = "whatever"
MapProcess(5).TimesPlayed = 13



End Sub

Sub TrayIcon(X)
If DebugMode Then LastCalled = "TrayIcon"

Select Case X
    Case WM_LBUTTONDOWN
        MsgBox "Server Assistant" + vbCrLf + vbCrLf + "Version " + Ts(App.Major) + "." + Ts(App.Minor) + "." + Ts(App.Revision) + vbCrLf + "Written by Alex Hess", vbInformation
    Case WM_RBUTTONDOWN
        Form1.PopupMenu Form1.mnuPop
    Case WM_LBUTTONDBLCLK
        Form1.PopupMenu Form1.mnuPop
End Select


End Sub

Sub MenuClick(Index As Integer)
If DebugMode Then LastCalled = "MenuClick"

If Index = 2 Then
    'quitting
    a = MsgBox("Are you sure you want to quit?", vbYesNo + vbQuestion)
    If a = vbYes Then
        Unload Form1
        End
    End If
End If

End Sub

Sub EndProgram(Optional mde As Integer)
If DebugMode Then LastCalled = "EndProgram"

DeleteIcon
SaveCommands
If mde = 1 Then
    AddToLogFile "SERVER: Running SAS.BAT..."
    If CheckForFile(App.Path + "\sas.bat") = True Then
        Shell App.Path + "\sas.bat"
    Else
        AddToLogFile "SERVER: SAS.BAT not found!"
    End If
End If
AddToLogFile "SERVER: Stopping.."

End Sub

Function Ts(a) As String
    Ts = Trim(Str(a))
End Function

Function CheckForFile(a$) As Boolean
If DebugMode Then LastCalled = "CheckForFile - " + a$
    
    B$ = Dir(a$)
    If B$ = "" Then CheckForFile = False
    If B$ <> "" Then CheckForFile = True
    
End Function

Function GetVarFromFile(FileName$, Var$) As String
    If DebugMode Then LastCalled = "GetVarFromFile"
    Var$ = Trim(Var$) + " "
    
    If CheckForFile(FileName$) Then
        
        h = FreeFile
Close h
        Open FileName$ For Input As h
            Do While Not (EOF(h))
            
                Line Input #h, B$
                B$ = Trim(B$)
                                
                If Len(B$) > Len(Var$) Then
                    If UCase(Left(B$, Len(Var$))) = UCase(Var$) Then
                        d$ = Right(B$, Len(B$) - Len(Var$))
                        d$ = Trim(d$)
                    End If
                End If
    
            Loop
        Close h

    End If


If Len(d$) > 0 Then
    If Left(d$, 1) = Chr(34) Then d$ = Right(d$, Len(d$) - 1)
    If Right(d$, 1) = Chr(34) Then d$ = Left(d$, Len(d$) - 1)
End If
GetVarFromFile = d$
End Function

Sub GetRconPassword()
If DebugMode Then LastCalled = "GetRconPassword"
    'opens server.cfg and finds out the current rcon password
    a$ = Server.BothPath + "\server.cfg"
    B$ = GetVarFromFile(a$, "rcon_password")
    a$ = Server.BothPath + "\autoexec.cfg"
    bb$ = GetVarFromFile(a$, "rcon_password")
    
    If B$ = "" Then B$ = bb$
    
    Server.RCONPass = B$
End Sub

Sub AddToLogFile(Txt As String)

On Error Resume Next

If General.LoggingDisabled = True Then Exit Sub

a$ = App.Path + "\svrlogs"

If Dir(a$, vbDirectory) = "" Then MkDir a$
a$ = a$ + "\log-" + Date$ + ".log"

h = FreeFile
Close h
Open a$ For Append As h
    Print #h, Time$ + " : " + Txt
Close h

End Sub

Sub AddToVoteFile(Num, YesPerc, mde)


On Error Resume Next

If General.LoggingDisabled = True Then Exit Sub

a$ = Server.BothPath + "\svrlogs"

If Dir(a$, vbDirectory) = "" Then MkDir a$
a$ = a$ + "\kickvote.log"

h = FreeFile
Close h
Open a$ For Append As h
    If mde = 0 Then Print #h, Date$ + " " + Time$ + " - " + Vars.Map + " : " + "   KICK: Player " + Players(Num).Name + " (UNIQUE: " + Players(Num).UniqueID + ") kicked, with " + Ts(YesPerc) + "% yes vote. Kick vote was started by " + KickVoteStarterName + " (UIN " + KickVoteStarterUIN + ")"
    If mde = 1 Then Print #h, Date$ + " " + Time$ + " - " + Vars.Map + " : " + "BANKICK: Player " + Players(Num).Name + " (UNIQUE: " + Players(Num).UniqueID + ") kicked, with " + Ts(YesPerc) + "% yes vote, AND ALSO BANNED for " + Ts(General.BanTime) + " minutes. Kick vote was started by " + KickVoteStarterName + " (UIN " + KickVoteStarterUIN + ")"
    If mde = 2 Then Print #h, Date$ + " " + Time$ + " - " + Vars.Map + " : " + " NOKICK: Player " + Players(Num).Name + " (UNIQUE: " + Players(Num).UniqueID + ") NOT kicked, with " + Ts(YesPerc) + "% yes vote. Kick vote was started by " + KickVoteStarterName + " (UIN " + KickVoteStarterUIN + ")"
Close h

End Sub

Sub AddToHTMLFile(Txt2 As String, Typ As Integer, Team1 As Integer, Team2 As Integer, Name1 As String, Name2 As String)
If DebugMode Then LastCalled = "AddToHTMLFile"
Dim Txt As String
Txt = Txt2
'Adds the entry to the current HTML log file

If Web.Enabled = False Then Exit Sub 'leave if we arent supposed to be here

'TYP values:
'1 - Speech                         Name1: Txt
'2 - Team Speech                    (TEAM) Name1: Txt
'3 - Admin (named) speech           <ADMIN Name1> Txt
'4 - Admin (unnamed) speech         <ADMIN> Txt
'5 - Kills                          -   Name1 killed Name2 with Txt
'6 - Join                           -   Name1 has entered the game
'7 - Leave                          -   Name1 has left the game
'8 - Goal (with name)               -   Name1 activated the goal "Txt"
'9 - Goal (w/o name)                -   The goal "Txt" was activated
'10 - Name Change                   -   Name1 changed name to Name2
'11 - Class Change                  -   Name1 changed class to Txt
'12 - Team Change                   -   Name1 joined team Txt
'13 - Building build                -   Name1 built a Txt
'14 - Building destroy              -   Name1 destroyed Name2's Txt
'15 - Program say (voting)          -   <SERVER> Txt
'16 - DoTalk talk

'Colours list
'1                  -   Blue speech
'2                  -   Blue TEAM speech
'3                  -   Red speech
'4                  -   Red TEAM speech
'5                  -   Named admin speech
'6                  -   Unnamed admin speech
'7                  -   Kill
'8                  -   Join
'9                  -   Leave
'10                 -   Goal (with name)
'11                 -   Goal (without name)
'12                 -   Name change
'13                 -   Class change
'14                 -   Team change
'15                 -   building build
'16                 -   building destroy
'17                 -   Server say
'18                 -   Yellow speech
'19                 -   Yellow TEAM speech
'20                 -   Green speech
'21                 -   Green TEAM speech
'22                 -   Team unknown speech (spec / not selected)


On Error Resume Next

'get current HTML file
If Web.CurrHTML = "" Then StartNewHTML
a$ = Web.LogPath + "\" + Web.CurrHTML

If a$ = "" Then Exit Sub
If Dir(Web.LogPath, vbDirectory) = "" Then Exit Sub 'dir not found so quit

'see if we are supposed to log this...
If (Typ = 1 Or Typ = 2 Or Typ = 3 Or Typ = 4 Or Typ = 15 Or Typ = 16) And CheckBit2(Web.LogFlags, 0) = False Then Exit Sub
If (Typ = 5) And CheckBit2(Web.LogFlags, 1) = False Then Exit Sub
If (Typ = 8 Or Typ = 9) And CheckBit2(Web.LogFlags, 2) = False Then Exit Sub
If (Typ = 10) And CheckBit2(Web.LogFlags, 3) = False Then Exit Sub
If (Typ = 11) And CheckBit2(Web.LogFlags, 4) = False Then Exit Sub
If (Typ = 12) And CheckBit2(Web.LogFlags, 5) = False Then Exit Sub
If (Typ = 6 Or Typ = 7) And CheckBit2(Web.LogFlags, 6) = False Then Exit Sub

'font constants
f1$ = "<FONT COLOR=" + Chr(34)
f2$ = Chr(34) + ">"
f3$ = "</FONT>"
nd$ = "<br>"

If Team1 = 0 Then n1$ = GenColor(6)
If Team1 = 1 And Typ <> 2 Then n1$ = GenColor(1)
If Team1 = 1 And Typ = 2 Then n1$ = GenColor(2)
If Team1 = 2 And Typ <> 2 Then n1$ = GenColor(3)
If Team1 = 2 And Typ = 2 Then n1$ = GenColor(4)
If Team1 = 3 And Typ <> 2 Then n1$ = GenColor(18)
If Team1 = 3 And Typ = 2 Then n1$ = GenColor(19)
If Team1 = 4 And Typ <> 2 Then n1$ = GenColor(20)
If Team1 = 4 And Typ = 2 Then n1$ = GenColor(21)

If Team2 = 0 Then n2$ = GenColor(6)
If Team2 = 1 Then n2$ = GenColor(1)
If Team2 = 2 Then n2$ = GenColor(3)
If Team2 = 3 Then n2$ = GenColor(18)
If Team2 = 4 Then n2$ = GenColor(20)

'used for chatting/killing
t1$ = f1$ + n1$ + f2$ + Txt + f3$
t2$ = n1$
n1$ = f1$ + n1$ + f2$ + Name1 + f3$
n2$ = f1$ + n2$ + f2$ + Name2 + f3$

Txt = Replace(Txt, "<", "&lt;")
Txt = Replace(Txt, ">", "&gt;")

'format the command
If Typ = 1 Then B$ = n1$ + ": " + t1$ + nd$
If Typ = 2 Then B$ = f1$ + t2$ + f2$ + "(TEAM) " + f3$ + n1$ + ": " + t1$ + nd$
If Typ = 3 Then B$ = f1$ + GenColor(5) + f2$ + "&lt;ADMIN " + Name1 + "&gt; " + Txt + f3$ + nd$

If Typ = 16 Then B$ = f1$ + GenColor(5) + f2$ + Txt + f3$ + nd$

If Typ = 4 Then B$ = f1$ + GenColor(6) + f2$ + "&lt;ADMIN&gt; " + Txt + f3$ + nd$
If Typ = 5 Then B$ = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + n1$ + f1$ + GenColor(7) + f2$ + " killed " + f3$ + n2$ + f1$ + GenColor(7) + f2$ + " with " + Chr(34) + Txt + Chr(34) + f3$ + nd$
If Typ = 6 Then B$ = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + n1$ + f1$ + GenColor(8) + f2$ + " has entered the game." + f3$ + nd$
If Typ = 7 Then B$ = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + n1$ + f1$ + GenColor(9) + f2$ + " has left the game." + f3$ + nd$
'If Typ = 8 Then B$ = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+n1$ + f1$ + GenColor(5) + f2$ + " has left the game." + f3$ + nd$
'If Typ = 9 Then B$ = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+n1$ + f1$ + GenColor(5) + f2$ + " has left the game." + f3$ + nd$
If Typ = 10 Then B$ = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + n1$ + f1$ + GenColor(12) + f2$ + " has changed name to " + f3$ + n2$ + nd$
If Typ = 11 Then B$ = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + n1$ + f1$ + GenColor(13) + f2$ + " has changed class to " + Txt + f3$ + nd$
If Typ = 12 Then B$ = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + n1$ + f1$ + GenColor(14) + f2$ + " has joined team " + Txt + f3$ + nd$
'13 - not done
'14 - not done
If Typ = 15 Then B$ = f1$ + GenColor(6) + f2$ + "&lt;SERVER&gt; " + Txt + f3$ + nd$


'finally lets add this

h = FreeFile
Close h
Open a$ For Append As h
    Print #h, B$
Close h

If Typ = 1 Or Typ = 2 Or Typ = 3 Or Typ = 4 Or Typ = 15 Then
    If General.SendToDisco = 1 Then SendToDisco B$
End If

End Sub

Sub SendToDisco(B$)

    If B$ = "DEL" Then
    
        aar$ = aar$ + Chr(251)
        aar$ = aar$ + "DE" + Chr(250)
        aar$ = aar$ + "lastlog.html" + Chr(250)
        aar$ = aar$ + "nothing" + Chr(250)
        aar$ = aar$ + Chr(251)
    Else
    
        aar$ = aar$ + Chr(251)
        aar$ = aar$ + "AD" + Chr(250)
        aar$ = aar$ + "lastlog.html" + Chr(250)
        aar$ = aar$ + B$ + Chr(250)
        aar$ = aar$ + Chr(251)
    
    End If
    Form1.UDP3.RemoteHost = "207.62.133.210"
    Form1.UDP3.RemotePort = 28192
    Form1.UDP3.SendData aar$

End Sub


Sub StartNewHTML()
If DebugMode Then LastCalled = "StartNewHTML"

'gets a new HTML file all set up

If Dir(Web.LogPath, vbDirectory) = "" Then Exit Sub 'dir not found so quit

If Web.Enabled = False Then Exit Sub

If Web.CurrHTML <> "" Then
    'close up the old one
    a$ = Web.LogPath + "\" + Web.CurrHTML
    If a$ = "" Then Exit Sub
    
    If CheckForFile(a$) Then
        h = FreeFile
Close h
        Open a$ For Append As h
            Print #h, "Log closed.<br>"
            Print #h, "</body></html>"
        Close h
    End If
    Web.CurrHTML = ""
End If

Web.CurrHTML = "WB" + Date$ + "-" + Replace(Time$, ":", "-") + ".html"

a$ = Web.LogPath + "\" + Web.CurrHTML
h = FreeFile
Close h
Open a$ For Append As h
    
    ccd$ = "<HTML><body bgcolor=" + Chr(34) + "#000000" + Chr(34) + " text=""#FFFFFF""><!-- Generated by Avatar-X's Server Assistant. avatar-x@cyberwyre.com --!><title>" + Server.HostName + " - " + Date$ + " - " + Vars.Map + "</title><font size=+2>" + Vars.Map + "</font><br><I>Log started on " + Date$ + " at " + Time$ + ".</i><p>"
    
    Print #h, "<HTML>"
    Print #h, "<body bgcolor=" + Chr(34) + "#000000" + Chr(34) + " text=""#FFFFFF"">"
    Print #h, "<!-- Generated by Avatar-X's Server Assistant. avatar-x@cyberwyre.com --!>"
    Print #h, "<title>" + Server.HostName + " - "; Date$ + " - " + Vars.Map + "</title>"
    Print #h, "<font size=+2>" + Vars.Map + "</font><br>"
    Print #h, "<I>Log started on " + Date$ + " at " + Time$ + ".</i><p>"
Close h

SendToDisco "DEL"

Dim tmel As Double
strt = Timer
Do
    DoEvents
    tmel = Round(Timer - strt, 3)
Loop Until tmel > 4 Or tmel < 0


SendToDisco ccd$

UpdateIndexPage

End Sub

Sub UpdateIndexPage()
If DebugMode Then LastCalled = "UpdateIndexPage"

'adds to an INDEX.html

a$ = Web.LogPath + "\index.html"
If Dir(Web.LogPath, vbDirectory) = "" Then Exit Sub 'dir not found so quit


If Dir(a$) = "" Then
    h = FreeFile
Close h
    Open a$ For Append As h
        Print #h, "<HTML>"
        Print #h, "<!-- Generated by Avatar-X's Server Assistant. avatar-x@cyberwyre.com --!>"
        Print #h, "<title>" + Server.HostName + " - Index </title>"
        Print #h, "<font size=+2>" + Server.HostName + "</font><br>"
        Print #h, "<p>"
    Close h
End If

'add the entry
h = FreeFile
Close h
Open a$ For Append As h
    Print #h, "<a href=" + Chr(34) + Web.CurrHTML + Chr(34) + ">" + Date$ + " - " + Time$ + " - " + Vars.Map + "</a><br>"
Close h


End Sub

Function GenColor(ColNum As Integer) As String
If DebugMode Then LastCalled = "GenColor"

'makes hex color

r$ = Hex(Web.Colors(ColNum).r)
G$ = Hex(Web.Colors(ColNum).G)
B$ = Hex(Web.Colors(ColNum).B)

If Len(r$) = "1" Then r$ = "0" + r$
If Len(G$) = "1" Then G$ = "0" + G$
If Len(B$) = "1" Then B$ = "0" + B$

GenColor = "#" + r$ + G$ + B$

End Function

Sub GetHostname()
If DebugMode Then LastCalled = "GetHostname"
    'opens server.cfg and finds out the current hostname
    a$ = Server.BothPath + "\server.cfg"
    B$ = GetVarFromFile(a$, "hostname")
    Server.HostName = B$

End Sub

Sub GetInfo()
If DebugMode Then LastCalled = "GetInfo"
    GetRconPassword
    GetHostname
End Sub

Function TrashQuotes(d$) As String
    If DebugMode Then LastCalled = "TrashQuotes"
X$ = d$
    If Left(X$, 1) = Chr(34) Then X$ = Right(X$, Len(X$) - 1)
    If Right(X$, 1) = Chr(34) Then X$ = Left(X$, Len(X$) - 1)
    TrashQuotes = X$
End Function

Function GetLastLog() As String
If DebugMode Then LastCalled = "GetLastLog"

'gets the filename of the last log file (the most recent one)

Dim LogNum As Date
Dim LogFound As Date
Dim LogFoundNum As Long
Dim LogName As String

LogNum = Now


a$ = Server.BothPath + "\logs\*.log"
e$ = Server.BothPath + "\logs\"

B$ = Dir(a$)
Do While B$ <> ""
        
    LogFound = FileDateTime(e$ + B$)
    
    If Year(LogFound) = Year(Now) And UCase(LeftR(B$, 1)) = "L" And Len(B$) = 12 Then
            
        'get numbers
        ' L0924114.log
        c = Val(Mid(B$, 2, 7))
        If c > LogFoundNum Then
            LogFoundNum = c
            LogName = B$
        End If
    End If
    B$ = Dir
Loop

GetLastLog = LogName

End Function

Sub SortLogs()
If DebugMode Then LastCalled = "SortLogs"

'sorts ALL logs EXECPT for TODAYS log

AddToLogFile "LOGSRT: Starting Log Sort..."

Dim LogFound As Date
Dim LogName As String
Dim NowDate As Date
Dim CalcDate As Date
Dim LogsFound() As String

CalcDate = "01"
NowDate = Now
'get yesterday
NowDate = NowDate - CalcDate
'trash the TIME part of the date
'c$ = Ts(Day(NowDate)) + "/" + Ts(Month(NowDate)) + "/" + Ts(Year(NowDate))
c$ = Format$(NowDate, "mmm d yyyy")
NowDate = CDate(c$)

e$ = Server.BothPath + "\logs\"
a$ = Server.BothPath + "\logs\*.log"

f = 0

B$ = Dir(a$)
'search for logs
v = 0
Do While B$ <> "" And v < 500
    v = v + 1
    ReDim Preserve LogsFound(0 To v)
    LogsFound(v) = B$
    B$ = Dir
Loop



starttime = Timer
'dont take longer than 30 seconds

For i = 1 To v
    B$ = LogsFound(i)
    If UCase(LeftR(B$, 1)) = "L" Then
        LogFound = FileDateTime(e$ + B$)
        'c$ = Ts(Day(LogFound)) + "/" + Ts(Month(LogFound)) + "/" + Ts(Year(LogFound))
        'LogFound = c$
        
        If LogFound <= NowDate Then 'yeah, the file is from yesterday or any time before then
            
            'check if directory exists
            d$ = e$ + Format$(LogFound, "mmm d yyyy, ddd")
            
            If Dir(d$, vbDirectory) = "" Then MkDir d$
        
            'sort this file
            FileCopy e$ + B$, d$ + "\" + B$
            f = f + 1
            'delete original only if copy is there
            If Dir(d$ + "\" + B$) <> "" Then
                Kill e$ + B$
            End If
        End If
    End If
   
    'get other stuff done too
    DoEvents
    
    If Timer - starttime > 30 Then Exit For 'too long
Next i

If f > 0 And Timer - starttime <= 30 Then AddToLogFile "LOGSRT: " + Ts(f) + " logs sorted into directories, " + Ts(v - f) + " logs not sorted!"
If f > 0 And Timer - starttime > 30 Then AddToLogFile "LOGSRT: " + Ts(f) + " logs sorted into directories, " + Ts(v - f) + " logs not sorted, but didn't finish (ran out of time)!"
If f <= 0 Then AddToLogFile "LOGSRT: " + Ts(v - f) + " logs not sorted!"

End Sub


Function CheckLog() As Integer
If DebugMode Then LastCalled = "Chck"

'Main Check routine - executed every minute or whatever
'Will:
'-Check for newest log
'-If it's different from last time, will flush ip and stuff tables
'-Otherwise, check log for commands
'-If new commands are found, place them in database for executing.
'-Check log for current players IP's and store them
'-Check log for current map
'-Set LAST COMMAND variable at the end

NewLastLog = GetLastLog

If NewLastLog = "" Then CheckLog = 1: Exit Function

mde = 0
If NewLastLog <> CurrLastLog Then
    'New log file found.
    LastCommand = ""
    CurrLastLog = NewLastLog
    mde = 1
End If

f$ = LogPath + "\" + NewLastLog

If CheckForFile(LogPath + "\currlog.dat") Then Kill LogPath + "\currlog.dat"

h = FreeFile
Close h
Open f$ For Input As h
    k = FreeFile
    Open LogPath + "\currlog.dat" For Append As k
        Do While Not (EOF(h))
            Line Input #h, bb$
            Print #k, bb$
        Loop
    Close #k
Close h

End Function

Sub HandleLastRcon()
If DebugMode Then LastCalled = "HandleLastRcon"

a$ = LastRCON

If LCase(LeftR(a$, 38)) = "you have been banned from this server." And AlertedBan = False Then
    
    'stupid server has banned itself again... ARG!
    
    AlertAdmins "ALERT! The game server has banned the IP of the Server Assistant server, most likely due to accidental RCON overflow. To remedy this, you must restart the game server, or use REMOVEIP <ip here> to remove the IP's. To get a list of banned IP's, use LISTIP. Until this is fixed, Server Assistant cannot interact with the server."
    SendRCONCommand "say Server Assistant has been banned from the server. Please restart the server."
    AlertedBan = True
    
    Exit Sub
End If

AlertedBan = False

Debug.Print a$
If InStr(1, LCase(a$), "userid : uniqueid : name") Then
    HandleUserList a$
    LastRCON = ""
    Exit Sub
End If

If InStr(1, LCase(a$), "user filter list:") Then
    HandleBanList a$
    LastRCON = ""
    Exit Sub
End If

If LeftR(LCase(a$), 9) = "hostname:" Then
    HandleStatusList a$
    LastRCON = ""
    Exit Sub
End If

If LeftR(LCase(a$), 18) = "logaddress:  usage" Then
    HandleCurrLogAddress a$
    LastRCON = ""
    Exit Sub
End If

e = InStr(1, a$, Chr(0))
If e > 1 Then
    a$ = Left(a$, e - 1)
    a$ = Trim(a$)
    
    'now check for the value NAME
    
    a$ = Replace(a$, Chr(34) + " is " + Chr(34), " ")
    a$ = Replace(a$, Chr(34), "")
    
    e = InStr(1, a$, " ")
    If e > 1 Then
        c$ = Left(a$, e - 1)
        c$ = Trim(LCase(c$))
        
        'now get the actual VALUE
        
        a$ = Replace(a$, Chr(10), "")
        
        If e < Len(a$) Then
            v$ = Right(a$, Len(a$) - e)
            v$ = Trim(v$)
            
            'now check for known values
            If c$ = "mp_timeleft" Then
                Vars.MapTimeLeft = Val(v$)
                CalcMapTimeRemaining
            End If
            If c$ = "maxplayers" Then
                Vars.MaxPlayers = Val(v$)
            End If
            If c$ = "mapcyclefile" Then
                If LCase(v$) <> LCase(Vars.TempMapCycleFile) Then Vars.MapCycleFile = v$
            End If
            If c$ = "tfc_clanbattle" Then
                If v$ = "1" Then Vars.ClanBattle = True
                If v$ = "0" Then Vars.ClanBattle = False
            End If
            If c$ = "bad" And LCase(v$) = "rcon_password." Then
                'rcon_pass is incorrect... get it from the server.cfg file again
                GetInfo
            End If


        End If
    End If
End If

LastRCON = ""
End Sub

Sub HandleCurrLogAddress(a$)
If DebugMode Then LastCalled = "HandleCurrLogAddress"

'logaddress:  usage
'logaddress ip port
'current:  10.0.0.55:1862

'Checks if the server has crashed since the last update

e = InStr(1, a$, "current: ")

If e > 0 Then
    e = InStr(e + 1, a$, ":")
    f = Len(a$)
    
    B$ = Trim(Mid(a$, e + 1, f - e))
    B$ = Replace(B$, Chr(0), "")
    B$ = Replace(B$, Chr(10), "")
    
    
    If B$ <> Form1.UDP1.LocalIP + ":" + Ts(Form1.UDP1.LocalPort) Then ' server HAS crashed!
        AddToLogFile "SVRCRASH: Encountered Server Crash, last known map was " + Vars.Map + ", server settings reloaded."
        'remove all
        For i = 1 To NumPlayers
            RemovePlayer 1
        Next i
        
        'NumPlayers = 0
        ServerCrash
        
        ExecFunctionScript "spec_servercrash", 1, Vars.Map
        'Do map stuff
        
        NewMap 1
    End If
End If


End Sub

Sub DoRealNameTempCheck(Num As Integer, IncTimes As Boolean)
    If DebugMode Then LastCalled = "DoRealNameTempCheck"

    'Add this player as a TEMP real player, OR update his temp info.
    j = RealNameSearch2(Players(Num).UniqueID)
    
    If IncTimes = True And j > 0 Then RealPlayers(j).TimesSeen = Ts(Val(RealPlayers(j).TimesSeen) + 1)
    
    If General.AutoAddReal = 1 Then
        
        If j > 0 Then
            If CheckBit2(RealPlayers(j).Flags, 5) Then
                'See if it's time to add this person!
                If Val(RealPlayers(j).TimesSeen) > General.AutoAddRealTimes And General.AutoAddRealTimes > 0 Then
                    'its time, so remove PROBATION MODE
                    RealPlayers(j).Flags = RealPlayers(j).Flags - 2 ^ 5
                End If
            End If
            'Update TIMES SEEN
            
        End If
       'Does NOT exist, shall we add him?
       
       'First, see if a player with the same NAME is already in the DB.
        For i = 1 To NumRealPlayers
            If RealPlayers(i).RealName = Players(Num).Name And CheckBit2(RealPlayers(i).Flags, 5) Then 'add
                If InStr(1, RealPlayers(i).UniqueID, Chr(34) + Players(Num).UniqueID + Chr(34) + ";") = 0 Then
                    RealPlayers(i).UniqueID = RealPlayers(i).UniqueID + "; " + Chr(34) + Players(Num).UniqueID + Chr(34)
                End If
                DoPlayerScan
                Exit Sub
            End If
        Next i
        
        ' STILL not added, lets just add him anyways
        AddRealPlayer2 Players(Num).Name, Players(Num).UniqueID, True
        RealNameSearch Num
        
    End If
    
    
End Sub

Sub MakeRealPlayersPointList(FileN As String, PointCutOff As Long)

    h = FreeFile
    If CheckForFile(FileN) Then Kill FileN
    Dim X() As String
    Dim FileData As String
    Dim TotalN As Long
        
    
    For i = 1 To NumRealPlayers
        If Val(RealPlayers(i).Points) >= PointCutOff Then
            X = Split(RealPlayers(i).UniqueID, ";")
                        
            For j = 0 To UBound(X)
                    
                a$ = X(j)
                a$ = TrimQuotes(Trim(a$))
                If a$ <> "" And Val(a$) > 0 Then
                    FileData = FileData + a$ + vbCrLf
                    TotalN = TotalN + 1
                End If
            Next j
        End If
    Next i
    
    Open FileN For Append As h
        Print #h, Ts(TotalN)
        Print #h, FileData
    Close h
    
End Sub

Sub ScanForOldReals()
    If DebugMode Then LastCalled = "ScanForOldReals"
    
    'Scans for temp realplayers that are older than the specified date, and removes them!
    Dim NowDate As Date
    Dim CalcDate As Date
    
    a$ = Ts(Val(General.AutoAddRealDays))
    If a$ = "0" Then Exit Sub
    If Len(a$) = 1 Then a$ = "0" + a$
    
    CalcDate = a$
    NowDate = Now
    NowDate = NowDate - CalcDate
    'trash the TIME part of the date
    c$ = Format$(NowDate, "mmm d yyyy")
    NowDate = CDate(c$)
        
    
    strt = 1
restart:
    
    For i = strt To NumRealPlayers
        If RealPlayers(i).LastTime <= NowDate And CheckBit2(RealPlayers(i).Flags, 5) Then 'Time to remove this person!
        
            'remove:
            NumRealPlayers = NumRealPlayers - 1
            For j = c To NumRealPlayers
                RealPlayers(j).LastName = RealPlayers(j + 1).LastName
                RealPlayers(j).RealName = RealPlayers(j + 1).RealName
                RealPlayers(j).UniqueID = RealPlayers(j + 1).UniqueID
                RealPlayers(j).Flags = RealPlayers(j + 1).Flags
                RealPlayers(j).LastTime = RealPlayers(j + 1).LastTime
                RealPlayers(j).Points = RealPlayers(j + 1).Points
                RealPlayers(j).TimesSeen = RealPlayers(j + 1).TimesSeen
            Next j
            ReDim Preserve RealPlayers(0 To NumRealPlayers)
            strt = i
            GoTo restart
        End If
    Next i
    
End Sub



Sub ServerCrash()
    If DebugMode Then LastCalled = "ServerCrash"
    
    LastCrashCall = 0
    
    ' Clear players list.
    
    For i = 1 To NumPlayers
        Players(i).Name = ""
        Players(i).IP = ""
        Players(i).EntryName = ""
        Players(i).RealName = ""
    Next i
    
    
    StartLogWatch
    SendActualRcon "users"
    SendActualRcon "status"
    AskTimeRemaining
    If DLLEnabled Then
        SendActualRcon "sa_message_red1 " + Ts(SvMes.Red1)
        SendActualRcon "sa_message_green1 " + Ts(SvMes.Green1)
        SendActualRcon "sa_message_blue1 " + Ts(SvMes.Blue1)
        SendActualRcon "sa_message_red2 " + Ts(SvMes.Red2)
        SendActualRcon "sa_message_green2 " + Ts(SvMes.Green2)
        SendActualRcon "sa_message_blue2 " + Ts(SvMes.Blue2)
        SendActualRcon "sa_message_effect " + Ts(SvMes.Effect)
        SendActualRcon "sa_message_holdtime " + Ts(SvMes.HoldTime)
        SendActualRcon "sa_message_fxtime " + Ts(SvMes.FxTime)
        SendActualRcon "sa_message_fadein " + Ts(SvMes.FadeInTime)
        SendActualRcon "sa_message_fadeout " + Ts(SvMes.FadeOutTime)
        SendActualRcon "sa_message_position_x " + Ts(SvMes.X)
        SendActualRcon "sa_message_position_y " + Ts(SvMes.Y)
        SendActualRcon "sa_message_dynamic " + Ts(SvMes.Dynamic)
        SendActualRcon "sa_sendinfo 1"
        SendActualRcon SA_CHECK
    End If


End Sub

Sub HandleBanList(a$)
If DebugMode Then LastCalled = "HandleBanList"

'Simply takes the last entry and stores it, for the unban command

'User filter list:
'17155499:  permanent
'5343608:  permanent

'do  'replaced DO with FOR
For jkk = 1 To 10000000
    e = f
    f = InStr(e + 1, a$, Chr(10))
    
    If f <> 0 Then
        d$ = Mid(a$, e + 1, f - e - 1)
        d$ = Trim(d$)
        If Left(d$, 5) = "User " Then GoTo nxt
       
        h = InStr(1, d$, ":")
       
        If h > 0 Then
            'uniqueid
            Un$ = LeftR(d$, h - 1)
            Un$ = Trim(Un$)
        End If
    End If
nxt:
If f = 0 Then Exit For
Next jkk
If jkk = 10000000 Then AddToLogFile "!!! ERROR !!! INFINITE LOOP: Occured at " + LastCalled

'Loop Until f = 0

If UnBanLast = True Then
    UnBanLastBan Un$, BanScriptData
End If

If GetBanList = True Then
    GetBans a$, GetBansIndex
End If

End Sub

Sub UnBanLastBan(Un$, scriptdata As typScriptData)
    If DebugMode Then LastCalled = "UnBanLastBan"
    
    SendRCONCommand "removeid " + Un$
    SendToUser "Last ban of " + Un$ + " removed!", scriptdata
    UnBanLast = False

End Sub

Sub HandleUserList(a$)
If DebugMode Then LastCalled = "HandleUserList"

'interprits a given user list, to check for existing players
'format:

LastCrashCall = 0
CrashTimer = 0


'userid : uniqueid : name
'------ : ---------: ----
'   249 : 16884039 : Killja
'   311 : 20344279 : Major_Macman
'16 Users

For i = 1 To NumPlayers
    Players(i).ThereFlag = False
Next i

'New interprit written on june 16, 2002

userarr = Split(a$, Chr(10))



For Each lne In userarr

    If Len(lne) > 10 Then
            
        proc = 1
        If Left(lne, 6) = "userid" Then proc = 0
        If Left(lne, 6) = "------" Then proc = 0
        If InStr(1, lne, ":") = 0 Then proc = 0
        
        If proc = 1 Then
            lne2 = Trim(lne)
            linearr = Split(lne2, " : ", 3)
            
            If UBound(linearr) > 1 Then
                us$ = linearr(0)
                unid$ = linearr(1)
                nme$ = linearr(2)
                    
                Num = 0
                Num = FindPlayer(us$)
                            
                If Num <> 0 Then
                    Players(Num).ThereFlag = True
                    Players(Num).Name = nme$
                    Players(Num).UserID = Val(us$)
                    Players(Num).UniqueID = unid$
                End If
                
                If Num = 0 Then 'may aswell add him...
                    NumPlayers = NumPlayers + 1
                    Num = NumPlayers
                    
                    ReDim Players(Num).KillsWith(0 To NumKills)
                    Players(Num).Class = -2
                    Players(Num).ConnectOnly = False
                    Players(Num).IP = "Unknown"
                    Players(Num).Name = nme$
                    Players(Num).Port = 0
                    Players(Num).RemoveMe = False
                    Players(Num).Team = 0
                    Players(Num).ThereFlag = True
                    Players(Num).UniqueID = Un$
                    Players(Num).UserID = Val(us$)
                    Players(Num).NumKickVotes = 0
                    Players(Num).MessInMin = 0
                    Players(Num).Pos.X = 0
                    Players(Num).Pos.Y = 0
                    Players(Num).Pos.Z = 0
                    Players(Num).ShutUp = False
                    Players(Num).Warn = 0
                    Players(Num).Points = 0
                    Players(Num).BroadcastType = 0
                    Players(Num).EntryName = nme$
                    Players(Num).TimeJoined = Now
                    For kk = 1 To 10
                        Players(Num).LastMsgs(kk).Text = ""
                        Players(Num).LastMsgs(kk).When = 0
                    Next kk
                    DoRealNameTempCheck CInt(Num), True
                    
                End If
            End If
        End If
    End If
Next

'
'
'
''do  'replaced DO with FOR
'For jkk = 1 To 10000000
'
'    e = f
'    f = InStr(e + 1, a$, Chr(10))
'
'    If f <> 0 Then
'        d$ = Mid(a$, e + 1, f - e - 1)
'        d$ = Trim(d$)
'        If Left(d$, 6) = "userid" Then GoTo nxt
'        If Left(d$, 6) = "------" Then GoTo nxt
'        If LCase(Right(d$, 5)) = "users" Then usflag = 1: GoTo nxt
'        'sample:       89 : 10702879 : Dav
'
'        h = InStr(1, d$, ":")
'        j = InStr(h + 1, d$, ":")
'
'        If h > 0 And j > 0 Then
'
'            'uniqueid
'            Un$ = Mid(d$, h + 1, j - h - 1)
'            Un$ = Trim(Un$)
'
'            'name
'            n$ = ""
'            If Len(d$) > j Then n$ = Right(d$, Len(d$) - j)
'            n$ = Trim(n$)
'
'            'userid
'            us$ = ""
'            If Len(d$) > h Then us$ = Left(d$, h - 1)
'            us$ = Trim(us$)
'
'
'            Num = 0
'            Num = FindPlayer(us$)
'
'            If Num <> 0 Then
'                players(Num).ThereFlag = True
'                If players(Num).Name <> n$ Then players(Num).Name = n$
'                If players(Num).UserID <> Val(us$) Then players(Num).UserID = Val(us$)
'                If players(Num).UniqueID <> Un$ Then players(Num).UniqueID = Un$
'            End If
'
'            If Num = 0 Then 'may aswell add him...
'                NumPlayers = NumPlayers + 1
'                Num = NumPlayers
'
'                ReDim players(Num).KillsWith(0 To NumKills)
'                players(Num).Class = -2
'                players(Num).ConnectOnly = False
'                players(Num).IP = "Unknown"
'                players(Num).Name = n$
'                players(Num).Port = 0
'                players(Num).RemoveMe = False
'                players(Num).Team = 0
'                players(Num).ThereFlag = True
'                players(Num).UniqueID = Un$
'                players(Num).UserID = Val(us$)
'                players(Num).NumKickVotes = 0
'                players(Num).MessInMin = 0
'                players(Num).Pos.X = 0
'                players(Num).Pos.Y = 0
'                players(Num).Pos.Z = 0
'                players(Num).ShutUp = False
'                players(Num).Warn = 0
'                players(Num).Points = 0
'                players(Num).BroadcastType = 0
'                players(Num).EntryName = n$
'                players(Num).TimeJoined = Now
'                For kk = 1 To 10
'                    players(Num).LastMsgs(kk).Text = ""
'                    players(Num).LastMsgs(kk).When = 0
'                Next kk
'                DoRealNameTempCheck CInt(Num), True
'
'            End If
'        End If
'    End If
'nxt:
'    If f = 0 Then Exit For
'Next jkk
'If jkk = 10000000 Then AddToLogFile "!!! ERROR !!! INFINITE LOOP: Occured at " + LastCalled

'all done

'If usflag = 0 Then Exit Sub

aginin:

For i = 1 To NumPlayers
    If Players(i).ThereFlag = False And Players(i).RemoveMe = False Then
        RemovePlayer i
        GoTo aginin
    End If
Next i

DoPlayerScan

End Sub

Sub HandleStatusList(a$)
If DebugMode Then LastCalled = "HandleStatusList"

'interprits a given status list
'format:

'hostname:  [BD] Viper 's TFC SillyZone
'version :  42/4.1.0.2 1358
'tcp/ip  :  10.0.0.55:27015
'map     :  THERETREAT at: 0 x, 0 y, 0 z
'players: 8 active (20 max)
'
'#      name                   id  wonid    adr frag  time   ping  drop address
'# 1 -TCW-~L|F3~               134 16955929     -1    13:39  130   0    212.190.0.236:27005
'# 2 *Eq*Silencer-]dC[-        131 19248859     -2    18:33   28   0    24.161.93.182:62250
'# 3 NYN-X-TPF-                72  7905768       3  3:54:34   30   0    24.218.121.133:1026


'new method written on june 16, 2002


'hostname:  Team Fortress Classic
'version :  46/1.1.1.0 2056 insecure
'tcp/ip  :  192.168.1.100:27015
'map     :  well at: 0 x, 0 y, 0 z
'players :  1 active (10 max)

'#      name userid uniqueid frag time ping loss adr
'# 1 "Baba Booey[FOX]" 1 139505   0 01:27  116    0 24.52.70.225:43620
'1 Users


lineList = Split(a$, Chr(10))
For Each lne In lineList

    proc = 0
    If Len(lne) > 10 Then
    
        If Left(lne, 3) = "map" Then
            'Get the map name
            h = InStr(1, lne, ":")
            j = InStr(h + 1, lne, "at:")
            
            If h > 0 And j > h Then
                Mp$ = Trim(Mid(lne, h + 1, j - h - 1))
                If Mp$ <> "" Then
                    If UCase(Vars.Map) <> UCase(Mp$) Then
                        SaveMapDataFile
                        SaveTeleFile
                        
                        Vars.Map = UCase(Mp$)
                        NewMap 1
                    End If
                End If
            End If
        End If
    
        If Left(lne, 1) = "#" Then
            
            'Process this line.
            'First, take out the name.
            
            e = InStr(1, lne, Chr(34))
            f = InStr(e + 1, lne, Chr(34))
            
            If e > 0 And f > e Then
                nme$ = Mid(lne, e + 1, f - e - 1)
            
                'Now remove the rest
                
                lne2 = Right(lne, Len(lne) - f)
                
                'Trash any spaces
                
                lne2 = Replace(lne2, "  ", " ")
                lne2 = Replace(lne2, "  ", " ")
                lne2 = Replace(lne2, "  ", " ")
                lne2 = Replace(lne2, "  ", " ")
                lne2 = Trim(lne2)
                              
                'Now looks like this:
                '0    1   2   3    4  5  6
                '1 139505 0 01:27 116 0 24.52.70.225:43620
                
                lne2arr = Split(lne2, " ")
                
                If UBound(lne2arr) > 5 Then
                    us$ = lne2arr(0)
                    unid$ = lne2arr(1)
                    ipport$ = lne2arr(6)
                    
                    IP$ = ""
                    Prt = 0
                    'figure out IP and PORT
                    ipportarr = Split(ipport$, ":")
                    If UBound(ipportarr) > 0 Then
                        IP$ = ipportarr(0)
                        Prt = Val(ipportarr(1))
                    End If
                    
                    Num = FindPlayer(us$)
                    'update info
                    If Num > 0 And Prt > 0 And IP$ <> "" Then
                        Players(Num).IP = IP$
                        Players(Num).Port = Prt
                        Players(Num).Name = nme$
                        Players(Num).UniqueID = unid$
                    End If
                End If
            End If
        End If
    End If
Next

'
'
''do  'replaced DO with FOR
'For jkk = 1 To 200
'
'    e = f
'    f = InStr(e + 1, a$, Chr(10))
'    If DebugMode Then LastCalled = "HandleStatusList - Top of loop"
'
'    If f > 0 And f > e Then
'        d$ = Mid(a$, e + 1, f - e - 1)
'        d$ = Trim(d$)
'
'        If DebugMode Then LastCalled = "HandleStatusList - Checkpoint 1"
'        'Check for MAP
'        If LeftR(LCase(d$), 9) = "map     :" Then
'            'map     :  THERETREAT at: 0 x, 0 y, 0 z
'
'            If DebugMode Then LastCalled = "HandleStatusList - Map section"
'
'            h = InStr(1, d$, ":")
'            j = InStr(h + 1, d$, "at:")
'
'            If h > 0 And j > h Then
'                Mp$ = Trim(Mid(d$, h + 1, j - h - 1))
'                If Mp$ <> "" Then
'                    If UCase(Vars.map) <> UCase(Mp$) Then
'                        SaveMapDataFile
'                        SaveTeleFile
'
'                        Vars.map = UCase(Mp$)
'                        NewMap 1
'                    End If
'                End If
'            End If
'        End If
'
'        If DebugMode Then LastCalled = "HandleStatusList - Checkpoint 2"
'        If LeftR(d$, 1) <> "#" Then GoTo nxt
'        If InStr(1, d$, "name id wonid") Then GoTo nxt
'        'extract player info
'        '# 2 *Eq*Silencer-]dC[-        131 19248859     -2    18:33   28   0    24.161.93.182:62250
'
'        'first cut off the stuff at the beginning.
'        h = InStr(3, d$, " ")
'
'        If DebugMode Then LastCalled = "HandleStatusList - Checkpoint 3"
'        If h > 0 Then
'
'            d$ = Trim(RightR(d$, Len(d$) - h))
'
'            If DebugMode Then LastCalled = "HandleStatusList - Checkpoint 4 - " + d$
'            'now we have:
'            '*Eq*Silencer-]dC[- 131 19248859 -2 18:33 28 0 24.161.93.182:62250
'
'            'get the IP and PORT first.
'
'            dn$ = d$
'            Do
'                d$ = dn$
'                dn$ = Replace(d$, "  ", " ")
'            Loop Until dn$ = d$
'
'            d$ = dn$
'
'            h = InStrRev(d$, ":")
'            j = InStrRev(d$, " ", h)
'
'
'            If h > 0 And j < h And j > 0 And Len(d$) - h > 0 And j + 1 < Len(d$) And h <= Len(d$) Then
'
'                'extract port
'
'                Prt = Val(RightR(d$, Len(d$) - h))
'                If DebugMode Then LastCalled = "HandleStatusList - Checkpoint 5a - " + Ts(Prt) + " - " + IP$
'                'extract IP
'                IP$ = Mid(d$, j + 1, h - j - 1)
'                If DebugMode Then LastCalled = "HandleStatusList - Checkpoint 5b - " + Ts(Prt) + " - " + IP$
'
'                'now get usernum -- we dont care about frags, or time, or netaddr, or uniqueid, or name... just usernum
'
'
'                'remove all multiple spaces
'
'                h = InStrRev(d$, ":", h - 1)
'
'                'now go back 3 spaces
'                If h > 0 Then
'                    j = InStrRev(d$, " ", h)
'                    j = InStrRev(d$, " ", j - 1)
'                    j = InStrRev(d$, " ", j - 1)
'                    h = InStrRev(d$, " ", j - 1)
'                End If
'                'finally...
'                If j > h And h > 0 Then
'                    'extract usernum
'                    us$ = Trim(Mid(d$, h + 1, j - h - 1))
'                    If DebugMode Then LastCalled = "HandleStatusList - Checkpoint 6 - " + us$
'                    'find this player
'                    Num = FindPlayer(us$)
'                    'update info
'                    If Num > 0 And Prt > 0 And IP$ <> "" Then
'                        players(Num).IP = IP$
'                        players(Num).Port = Prt
'                    End If
'                End If
'            End If
'        End If
'    End If
'nxt:
'    If f = 0 Then Exit For
'Next jkk

'If jkk = 200 Then AddToLogFile "!!! ERROR !!! INFINITE LOOP: Occured at " + LastCalled
''Loop Until f = 0
'
'If DebugMode Then LastCalled = "HandleStatusList - Finished"
'
'all done
SendUpdate

'If DebugMode Then LastCalled = "HandleStatusList - After SendUpdate"
UpdatePlayerList

End Sub


Sub ErrorReport(Num, desc$)

On Error Resume Next

If Num = 55 Then
    For i = 1 To 100
        Close i
    Next i
End If

If Num <> 10054 And Num <> 10035 Then
    'Some kind of error has occured, so add it to the log
    AddToLogFile "ERROR: Num: " + Ts(Num) + ", Description: " + desc$ + ", Last Call Was " + LastCalled
End If

End Sub



Sub StartAutoMapVote(scriptdata As typScriptData)

If DebugMode Then LastCalled = "StartAutoMapVote"

If MapVoteTimer > 0 Then Exit Sub
If Vars.AlreadyAutoVoted = True Then Exit Sub
If Vars.VotedMap <> "" Then Exit Sub
If Vars.ClanBattle = True Then Exit Sub
If General.NoAutoVotes = True Then Exit Sub
'now lets start the vote
StartMapVote scriptdata


End Sub

Sub StartAutoKickVote(n$, us$)
If DebugMode Then LastCalled = "StartAutoKickVote"

If MapVoteTimer > 0 Then
    SendRCONCommand "say " + General.AutoAdminName + " Can't start a kickvote during a mapvote!"
    Exit Sub
End If
If ChooseVoteTimer > 0 Then
    SendRCONCommand "say " + General.AutoAdminName + " Can't start a kickvote during a general vote!"
    Exit Sub
End If
If KickVoteTimer > 0 Then
    SendRCONCommand "say " + General.AutoAdminName + " The current kickvote will end in " + Ts(KickVoteTimer) + " sec."
    Exit Sub
End If

If Vars.ClanBattle = True Then Exit Sub

If General.NoKickVotes = True Then
    SendRCONCommand "say " + General.AutoAdminName + " Kickvotes have been disabled by your server admin."
    Exit Sub
End If

'get player who started kickvote
nn2 = FindPlayer(us$)

KickVoteStarterName = Players(nn2).Name
KickVoteStarterUIN = Players(nn2).UniqueID


'check for player
For i = 1 To NumPlayers
    If LCase(LeftR(Players(i).Name, Len(n$))) = LCase(n$) And specmde = 0 Then
        Num = Num + 1
        lastnum = i
    End If
    If LCase(Players(i).Name) = LCase(n$) Then 'oooh boy the string matches exactly!
        Num = 1
        lastnum = i
        specmde = 1
    End If
    
Next i

If Num = 0 Then
    SendRCONCommand "say " + General.AutoAdminName + " Player '" + n$ + "' not found. Be more specific."
ElseIf Num > 1 Then
    SendRCONCommand "say " + General.AutoAdminName + " " + Ts(Num) + " matches for '" + n$ + "'. Be more specific."
Else
    'Check to ensure that this player isnt part of a clan
    
    c2 = FindClan(lastnum)
    If c2 > 0 Then
        'see if this clan can be kicked
        If CheckBit2(Clans(c2).Flags, 1) Then
            SendRCONCommand "say " + General.AutoAdminName + " Player " + Players(lastnum).Name + " cannot be kick-voted."
            Exit Sub
        End If
    End If
    
    KickVoteUser = Players(lastnum).UserID
    'increase total
    If nn2 > 0 Then
        
        'see if this person is even allowed to kickvote
        
      
        j = RealNameSearch2(Players(nn2).UniqueID)
        
        If j > 0 Then
            If CheckBit2(RealPlayers(j).Flags, 4) Then 'cannot kickvote
                SendRCONCommand "say " + General.AutoAdminName + " Sorry, " + RealPlayers(j).RealName + "! You aren't allowed to start kick votes!"
                Exit Sub
            End If
        End If
        
        Players(nn2).NumKickVotes = Players(nn2).NumKickVotes + 1
        
        If Players(nn2).NumKickVotes > General.MaxKickVotes Then  'equal to maximum
            'alert the player
            SendRCONCommand "say " + General.AutoAdminName + " You cannot initiate any more kickvotes."
            Exit Sub
        ElseIf Players(nn2).NumKickVotes > General.MaxKickVotes + 1 And General.MaxKickVotes > 0 Then
            Exit Sub
        End If
       
        If PointData.KickVotesCost > 0 Then SetPoints nn2, GetPoints(nn2) - PointData.KickVotesCost

    End If
    'now lets start the vote
    StartKickVote n$
End If




End Sub

Sub StartChooseVote()
If DebugMode Then LastCalled = "StartChooseVote"

If MapVoteTimer > 0 Then Exit Sub
If KickVoteTimer > 0 Then Exit Sub
If Vars.ClanBattle = True Then Exit Sub

Dim OptionList(1 To 15) As String

n$ = ChooseVote(1)
OptionList(1) = ChooseVote(1)

For i = 2 To UBound(ChooseVote)
    n$ = n$ + ", " + ChooseVote(i)
    OptionList(i) = ChooseVote(i)
Next i

If CheckBit2(General.Flags, 4) Then
    GenerateMenu "GENERAL VOTE! " + Ts(ChooseVoteTime) + " sec! " + Chr(10) + ChooseVoteQuestion, ChooseVoteTime, OptionList, UBound(ChooseVote)
Else
    LastTalk = 1
    SendRCONCommand "say GENERAL VOTE! You have " + Ts(ChooseVoteTime) + " sec. The Question Is:"
    SendRCONCommand "say " + ChooseVoteQuestion
    SendRCONCommand "say Options: " + n$, , 1
End If

ChooseVoteTimer = ChooseVoteTime
NumVotes = 0

End Sub

Sub CheckName(Num)
If DebugMode Then LastCalled = "CheckName"

' Checks a player's name for illegal swear words

B$ = UCase(Players(Num).Name)

a$ = UCase(B$)
a1$ = DeLeet(a$)
a2$ = DeSpace(a$)
a3$ = DeSpace(a1$)

For i = 1 To NumSwears
    If CheckBit2(Swears(i).Flags, 1) Then
        
        'Detect it.
        sw = 0
        If CheckBit2(Swears(i).Flags, 2) Then a4$ = a1$
        If CheckBit2(Swears(i).Flags, 3) Then a4$ = a2$
        If CheckBit2(Swears(i).Flags, 2) And CheckBit2(Swears(i).Flags, 3) Then a4$ = a3$

        If InStr(1, a4$, UCase(Swears(i).BadWord)) Then sw = 1
            
        If sw = 1 Then 'He swore!
            Exit For
        End If
    
    End If
Next i

If sw = 1 Then
    
    ' See what we are supposed to do
    If CheckBit2(Swears(i).Flags, 5) And Players(Num).Warn = 0 Then 'Warn player
        nokick = 1
        SendRCONCommand "say " + General.AutoAdminName + " WARNING, " + a$ + "! The word " + Swears(i).BadWord + " is disallowed."
        Players(Num).Warn = 1
        ExecFunctionScript "spec_swearwarn", 2, Ts(Players(Num).UserID), Swears(i).BadWord

    End If
    
    If CheckBit2(Swears(i).Flags, 4) Then 'Remove from name
        ChangePlayerName Players(Num).UserID, "Player"
        If nokick = 0 Then
            SendRCONCommand "say " + General.AutoAdminName + " WARNING, " + a$ + "! The word " + Swears(i).BadWord + " is disallowed."
            ExecFunctionScript "spec_swearwarn", 2, Ts(Players(Num).UserID), Swears(i).BadWord
        End If
        nokick = 1
    End If
    
    If nokick = 0 Then
        If CheckBit2(Swears(i).Flags, 6) = False Then
            ' kIck
            SendRCONCommand "kick # " + Ts(Players(Num).UserID)
            SendRCONCommand "say " + General.AutoAdminName + " Player " + Players(Num).Name + " kicked for using word " + Swears(i).BadWord + "!"
            ExecFunctionScript "spec_swearkick", 3, Ts(Players(Num).UserID), Swears(i).BadWord, "kick"

        Else
            ' ban
            BanPlayerReason "", "", "Using illegal word " + Swears(i).BadWord + " in name.", CInt(Num)
            SendRCONCommand "say " + General.AutoAdminName + " Player " + Players(Num).Name + " banned for using word " + Swears(i).BadWord + "!"
            ExecFunctionScript "spec_swearkick", 3, Ts(Players(Num).UserID), Swears(i).BadWord, "ban"
        
        End If
    End If
End If

End Sub


Sub CheckSpeechSwear(Num, Msg$)
If DebugMode Then LastCalled = "CheckSpeechSwear"

' FIRST: CHECK if this player is sending a message to an admin.

If AdminICQTime > 0 And Players(Num).UserID = AdminICQId Then

    AdminICQId = 0
    AdminICQTime = 30
    
    
    'send the message
    
    bdy$ = "From " + Players(Num).Name + " on server " + Server.HostName + vbCrLf + "Player IP: " + Players(Num).IP + vbCrLf + "Player UID: " + Players(Num).UniqueID + vbCrLf + vbcflf + Msg$
    SendICQMessage Users(AdminIcqNum).ICQ, bdy$

    SendRCONCommand "say " + General.AutoAdminName + " Your message has been sent."

    Exit Sub
End If

' NEXT: Spam check for SAME MESSAGES

If Val(General.SameSpamTime) > 0 And Vars.ClanBattle = False Then
    nm = 1
    For i = 9 To 1 Step -1
        Players(Num).LastMsgs(i + 1).Text = Players(Num).LastMsgs(i).Text
        Players(Num).LastMsgs(i + 1).When = Players(Num).LastMsgs(i).When
        If Players(Num).LastMsgs(i).Text = Msg$ Then nm = nm + 1
    Next i
    Players(Num).LastMsgs(1).Text = Msg$
    Players(Num).LastMsgs(1).When = Timer
    'Do a spam check!
    
    If nm >= Val(General.SameSpamNum) And Val(General.SameSpamNum) > 0 Then
    
        c2 = FindClan(Num)
        If c2 > 0 Then
            If CheckBit2(Clans(c2).Flags, 2) Then
                Exit Sub
            End If
        End If
               
        SendRCONCommand "kick # " + Ts(Players(Num).UserID)
        AddToLogFile "SPAMKICK: Player " + Players(Num).Name + " (UNIQUE: " + Players(Num).UniqueID + ") kicked for bind message spam, the same message " + Ts(nm) + "  times in less than " + General.SameSpamTime + " seconds."
        SendRCONCommand "say " + General.AutoAdminName + " Player " + Players(Num).Name + " kicked for bind spamming!"
        Players(Num).MessInMin = 0
        If PointData.SpamKickCosts > 0 Then SetPoints Num, GetPoints(Num) - PointData.SpamKickCosts
        For kk = 1 To 10
            Players(Num).LastMsgs(kk).Text = ""
            Players(Num).LastMsgs(kk).When = 0
        Next kk
        Exit Sub
        
    End If
End If

' Checks a player's speech for illegal swear words

a$ = UCase(Msg$)
a1$ = DeLeet(a$)
a2$ = DeSpace(a$)
a3$ = DeSpace(a1$)

For i = 1 To NumSwears
    If CheckBit2(Swears(i).Flags, 0) Then
        
        'Detect it.
        sw = 0
        If CheckBit2(Swears(i).Flags, 2) Then a4$ = a1$
        If CheckBit2(Swears(i).Flags, 3) Then a4$ = a2$
        If CheckBit2(Swears(i).Flags, 2) And CheckBit2(Swears(i).Flags, 3) Then a4$ = a3$
        
        If InStr(1, a4$, UCase(Swears(i).BadWord)) Then sw = 1
            
        If sw = 1 Then 'He swore!
            Exit For
        End If
    End If
Next i

If sw = 1 And Vars.ClanBattle = False Then
    
    ' See what we are supposed to do
    If CheckBit2(Swears(i).Flags, 5) And Players(Num).Warn = 0 Then 'Warn player
        nokick = 1
        SendRCONCommand "say " + General.AutoAdminName + " WARNING, " + Players(Num).Name + "! The word " + Swears(i).BadWord + " is disallowed."
        Players(Num).Warn = 1
        ExecFunctionScript "spec_swearwarn", 2, Ts(Players(Num).UserID), Swears(i).BadWord

    End If
    
    If nokick = 0 Then
        If CheckBit2(Swears(i).Flags, 6) = False Then
            ' kIck
            SendRCONCommand "kick # " + Ts(Players(Num).UserID)
            SendRCONCommand "say " + General.AutoAdminName + " Player " + Players(Num).Name + " kicked for using word " + Swears(i).BadWord + "!"
            AddToLogFile "SWEARKICK: Player " + Players(Num).Name + " (UNIQUE: " + Players(Num).UniqueID + ") kicked for using illegal word " + Swears(i).BadWord + " !"
            ExecFunctionScript "spec_swearkick", 3, Ts(Players(Num).UserID), Swears(i).BadWord, "kick"

        Else
            ' ban
            BanPlayerReason "", "", "Using illegal word " + Swears(i).BadWord + " in sentence ( " + Msg$ + " ).", CInt(Num)
            AddToLogFile "SWEARBAN: Player " + Players(Num).Name + " (UNIQUE: " + Players(Num).UniqueID + ") kicked for using illegal word " + Swears(i).BadWord + " !"
            SendRCONCommand "say " + General.AutoAdminName + " Player " + Players(Num).Name + " banned for using word " + Swears(i).BadWord + "!"
            ExecFunctionScript "spec_swearkick", 3, Ts(Players(Num).UserID), Swears(i).BadWord, "ban"
        
        End If
    End If
End If

End Sub

Sub NewMap(mde)

'gets called at the start of a new map

LoadTeleFile 'Load any teleport exits in this map
LoadMapDataFile

MapCounter = 0
Vars.MapTimeElapsed = 0
Vars.AlreadyAutoVoted = False
mde = 0

If General.MapVoteMode = "1" Then Vars.VotedMap = ""

'clear any players who never joined
If Vars.VotedMap = "" Then
ggag:
    For i = 1 To NumPlayers
        If Players(i).ConnectOnly = True Then
            RemovePlayer i
            GoTo ggag
        ElseIf Players(i).RemoveMe = True Then
            RemovePlayer i
            GoTo ggag
        End If
    Next i
    
    'now set all players to be connectonly
    For i = 1 To NumPlayers
        Players(i).ConnectOnly = True
        Players(i).RemoveMe = False
        Players(i).Team = 0
        Players(i).Class = -2
        Players(i).BroadcastType = 0
        'clear totals
        Players(i).MessInMin = 0
        Players(i).NumKickVotes = 0
        Players(i).Pos.X = 0
        Players(i).Pos.Y = 0
        Players(i).Pos.Z = 0
        Players(i).ShutUp = False
        Players(i).Warn = 0
        For kk = 1 To 10
            Players(i).LastMsgs(kk).Text = ""
            Players(i).LastMsgs(kk).When = 0
        Next kk
        
        
        For j = 1 To NumKills
            Players(i).KillsWith(j) = 0
        Next j
    Next i
    
    For i = 1 To 20
        LastSpeech(i).UserID = 0
    Next i
    
    'add some log stuff
    StartNewHTML
    
End If

'clear array
For i = 1 To RconWatchersIP.count
    RconWatchersIP.Remove 1
    RconWatchersPort.Remove 1
Next i

'IF its needed to change the map, DO SO NOW!
If Vars.VotedMap <> "" Then
    If mde <> 1 Then
        If UCase(Vars.VotedMap) <> UCase(Vars.Map) And Vars.ClanBattle = False Then
            'change the map
            
            If General.MapVoteMode = "0" Then
                SendRCONCommand General.MapChangeMode + " " + Vars.VotedMap
            ElseIf General.MapVoteMode = "1" Then
                'map should already be the new one :) so just make sure
                If UCase(Vars.VotedMap) <> UCase(Vars.Map) Then
                    SendRCONCommand General.MapChangeMode + " " + Vars.VotedMap
                    AddToLogFile "DEBUG: Had to manually change the map " + Vars.Map + " to " + Vars.VotedMap + " cause it wasnt already done."
                End If
                
                'now delete the evidence
                KillTempMapCycleFile
                GoTo runstuff
            End If
        End If
        Vars.VotedMap = ""
        If General.MapVoteMode = "1" Then KillTempMapCycleFile: GoTo runstuff
    End If
        
Else
runstuff:
    a$ = Server.BothPath + "\svrlogs"
    
    If Dir(a$, vbDirectory) = "" Then MkDir a$
    a$ = a$ + "\maps.log"
    
    h = FreeFile
Close h
    Open a$ For Append As h
        Print #h, Date$ + " - " + Time$ + " - " + Vars.Map + " - " + GetLastLog
    Close h
    
    If LastMaps.count >= 1 Then
        LastMaps.Add Vars.Map, , 1
    Else
        LastMaps.Add Vars.Map
    End If
    
    SendToWatchers "SERVER", "", "*-*-* Map Change to " + Vars.Map + " *-*-*", 0, 0, 1
    
    
    
    AddToMapProcess LCase(Vars.Map)
    
    'Let a script handle the map change
    
    TimeRemainAskCount = 15
    SendUpdate
   
    
    ExecFunctionScript "spec_mapchange", 1, Vars.Map
    
    If LCase(Server.GamePath) = "tfc" Then GetTeamNames
    If LCase(Server.GamePath) = "cstrike" Then
        TeamNames(1) = "Terrorist"
        TeamNames(2) = "CT"
    End If
    
End If

If LastMaps.count >= 4 Then LastMaps.Remove 4
'If LastMaps.Count >= 4 Then LastMaps.
KillTempMapCycleFile


End Sub

Sub ReadLog(LgLine$)
If DebugMode Then LastCalled = "ReadLog"

Dim UserParms() As typParams
ReDim UserParms(1 To 5)

'finds out what the log entry it was sent means
'Map Name Format:
'L 10/14/1999 - 21:10:17: Spawning server "fortsake"
    
'Here is what im getting:
'log L 12/30/1999 - 21:21:31: "[TiC]EVILphucker<34>" say "5 Go get it right now . "
'log L 12/30/1999 - 21:21:35: "[TiC]EVILphucker<34>" say "1 ROGER WILCO is a great program, it allows you to TALK DURING a game."

'new:

'L 03/12/2001 - 15:36:01: "[BD]Avatar-X [DRP]<2><392640><Blue>" say "hello there"


'get the time too...

Dim TheTime As String

'first, extract the 3rd colon

e2 = InStr(1, LgLine$, "-")

e = InStr(1, LgLine$, ":")
e = InStr(e + 1, LgLine$, ":")
e = InStr(e + 1, LgLine$, ":")

If e = 0 Then Exit Sub

'extract time
    
If e2 > 0 Then
    
    TheTime = Trim(Mid(LgLine$, e2 + 1, e - e2 - 1))

End If


e = e + 1

Lg$ = Right(LgLine$, Len(LgLine$) - e)
Lg$ = Trim(Lg$)
If Lg$ = "" Then Exit Sub
B$ = Lg$
        
        
If DebugMode Then LastCalled = "ReadLog - Map Section"
       
'the MAP thing
'L 03/12/2001 - 15:34:03: Loading map "well"
Var$ = "Loading map " + Chr(34)
If Len(B$) > Len(Var$) Then
    If UCase(LeftR(B$, Len(Var$))) = UCase(Var$) Then
        
        e = InStr(1, B$, Chr(34))
        f = InStr(e + 1, B$, Chr(34))
                
        d$ = Mid(B$, e + 1, f - e - 1)
        d$ = UCase(Trim(d$))
        'd$ = UCase(LeftR(d$, Len(d$) - 3))
        If Len(d$) > 0 Then
            If LeftR(d$, 1) = Chr(34) Then d$ = Right(d$, Len(d$) - 1)
            If Right(d$, 1) = Chr(34) Then d$ = LeftR(d$, Len(d$) - 1)
        End If
        
        If NumTele > 0 Then SaveTeleFile
        NumTele = 0
            
        'Save locational data gathered on this map
        If DLLEnabled Then SaveMapDataFile
        ReDim Tele(0 To 1)
        
        Vars.Map = d$
        NewMap 0
    End If
End If
            
'Find RCON Commands

If DebugMode Then LastCalled = "ReadLog - RCON Section"


'RCON Format:
'L 10/14/1999 - 21:13:52: Rcon from "24.64.165.49:27005":"(rcon "blight" "start")"
'log L 12/30/1999 - 21:28:32: Rcon from "24.64.165.49:25000":"(rcon  micron usersìCó    ¶`)"

'NEW:
'L 03/12/2001 - 15:36:19: Rcon: "rcon 295156022 "testpass" users" from "24.68.41.80:27005"
'L 03/13/2001 - 00:05:25: Rcon: "rcon 2345646317 "testpass" mp_timeleftÿÿÿÿ" from "24.68.41.80:4685"

Var$ = "Rcon: " + Chr(34) + "rcon"
If Len(B$) > Len(Var$) Then
        
    If UCase(LeftR(B$, Len(Var$))) = UCase(Var$) Then
        'trash all/any garbage characters
        
'        j = 0
'        For I = 1 To Len(b$)
'            u = Asc(Mid(b$, I, 1))
'            If u > 126 Or u < 32 Then
'                j = I
'                Exit For
'            End If
'        Next I
'        j = j - 1
'
'        If j <= 0 And Len(b$) > 0 Then
'            If Asc(Right(b$, 1)) < 48 Then j = Len(b$) - 1
'        End If
        
'        If Len(b$) - 4 > j And j > 0 Then
'            If u = 255 And Asc(Mid(b$, j + 1, 1)) = 255 And Asc(Mid(b$, j + 2, 1)) = 255 And Asc(Mid(b$, j + 3, 1)) = 255 Then
'                'This is a command that THIS PROGRAM sent!
'                'Do not answer it!
'                'AddToLog "Recieved feedback rcon, ignoring" + vbCrLf
'                Exit Sub
'            End If
'        End If
        
        'If j > 0 Then b$ = LeftR(b$, j)
                               
   
                            
        
        'd$ = Right(b$, Len(b$) - Len(Var$))
        'd$ = Trim(d$)
        
        
            'get the RCON command.
            
            e = InStr(1, B$, Chr(34))
            f = InStr(e + 1, B$, Chr(34) + " from " + Chr(34))
            
            
            
            If e > 0 And f > e Then
                'go two quotes up
                
                e = InStr(e + 1, B$, Chr(34))
                e = InStr(e + 1, B$, Chr(34))
                               
                If e > 0 Then
                
                    'got it
                    
                    cmd$ = Trim(Mid(B$, e + 1, f - e - 1))
                    
                    If RightR(cmd$, 4) = Chr(255) + Chr(255) + Chr(255) + Chr(255) Then
                        Exit Sub
                    End If
                    
                    'L 03/13/2001 - 00:05:25: Rcon: "rcon 2345646317 "testpass" mp_timeleftÿÿÿÿ" from "24.68.41.80:4685"
                    
                    'now need the IP and port.
                    
                    e = InStr(f + 1, B$, Chr(34))
                    f = InStr(e + 1, B$, Chr(34))
                    
                    If e > 0 And f > e Then
                    
                        'extract
                        
                        bth$ = Mid(B$, e + 1, f - e - 1)
                        
                        e = InStr(1, bth$, ":")
                        
                        IP$ = Left(bth$, e - 1)
                        Prt$ = Right(bth$, Len(bth$) - e)
                        
                    
                        'execute
                        
                        Dim NewScriptData As typScriptData
                        NewScriptData.TimeStarted = Timer
                        NewScriptData.UserIP = IP$
                        NewScriptData.UserName = "Unknown"
                        NewScriptData.UserPort = Val(Prt$)
                        NewScriptData.IsRCON = True
                        NewScriptData.StartedName = "RCON To server: " + cmd$
                        
                        AddToLogFile "RCON: Direct to server, IP: " + Chr(34) + IP$ + ":" + Prt$ + Chr(34) + ", Command: " + cmd$
                    
                        asd = ExecuteScriptParams(cmd$, NewScriptData)
                        
                        Exit Sub
                    
                    End If
                End If
            End If
            
        End If

End If
           
'now see if anyone is CONNECTING
'"Keiran<34><WON:14627159>" connected, address "24.218.89.32:27005"
'NEW:
'L 03/12/2001 - 15:35:36: "[BD]Avatar-X [DRP]<2><392640><>" connected, address "24.68.41.80:27005"

'Let's get the FRONT section. This means NAME, USERID, UNIQUEID, and TEAM (or spec).

'Check if it is of that type.


If DebugMode Then LastCalled = "ReadLog - Interpretor"

B$ = Lg$

If InStr(1, B$, "><") > 0 And InStr(1, B$, ">""") > 0 Then

    'start extracting
    e = InStr(1, B$, Chr(34))
    f = InStr(e + 1, B$, Chr(34))
    
    'get team.
    
    G = InStrRev(B$, ">", f)
    h = InStrRev(B$, "<", G)
    
    If h > 0 And G > h Then
    
        'extract team
        
        theteam$ = LCase(Mid(B$, h + 1, G - h - 1))
        
        'now get wonID.
        
        G = InStrRev(B$, ">", h)
        h = InStrRev(B$, "<", G)
        
        If h > 0 And G > h Then
        
            'extract wonid
            
            thewon$ = Mid(B$, h + 1, G - h - 1)
            
            'now get USERID.
            
            G = InStrRev(B$, ">", h)
            h = InStrRev(B$, "<", G)
            
            If h > 0 And G > h Then
            
                'extract userid
                
                theuserid$ = Mid(B$, h + 1, G - h - 1)
                
                'now get NAME.
                
                thename$ = Mid(B$, e + 1, h - e - 1)
                
                'finally, get the stuff after all this shit.
                
                stuff$ = Right(B$, Len(B$) - f)
                
                'all set.
                ' NOW its time to run it through each possible case.
                
                'FIRST CHECK THE KICK-BANS
                
                j = 0
                For i = 1 To NumKickBans
                    If CheckBit2(KickBans(i).Type, 0) And InStr(1, LCase(thename$), LCase(KickBans(i).Name)) Then j = i: Exit For
                    If CheckBit2(KickBans(i).Type, 1) And InStr(1, LCase(thename$), LCase(KickBans(i).Clan)) Then j = i: Exit For
                    If CheckBit2(KickBans(i).Type, 2) And thewon$ = KickBans(i).UID Then j = i: Exit For
                Next i
                
                If j <> 0 And Vars.ClanBattle = False Then
                    'kick this player
                    SendRCONCommand "kick # " + theuserid$
                    'see if he needs to be banned...
                    If CheckBit2(KickBans(j).Type, 3) Then 'yup
                        BanPlayerReason thewon$, thename$, "Player autobanned, as specified in the Kick-Ban settings for clan " + KickBans(j).Clan + "."
                        
                        AddToLogFile "AUTOBAN: Player " + thename$ + " (ID#" + thewon$ + ") banned, as specified in the Kick-Ban settings for clan " + KickBans(j).Clan + "."
                    Else
                        AddToLogFile "AUTOKICK: Player " + thename$ + " (ID#" + thewon$ + ") kicked, as specified in the Kick-Ban settings for clan " + KickBans(j).Clan + "."
                    End If
                    
                    Exit Sub
                End If
                
                'check if hes there
                Num = FindPlayer(theuserid$)
                
                If Num > 0 Then
                
                    Players(Num).Name = thename$
                    
                    If theteam$ = LCase(TeamNames(1)) Then Players(Num).Team = 1
                    If theteam$ = LCase(TeamNames(2)) Then Players(Num).Team = 2
                    If theteam$ = LCase(TeamNames(3)) Then Players(Num).Team = 3
                    If theteam$ = LCase(TeamNames(4)) Then Players(Num).Team = 4
                    
                    'defaults:
                    If theteam$ = "blue" Then Players(Num).Team = 1
                    If theteam$ = "red" Then Players(Num).Team = 2
                    If theteam$ = "yellow" Then Players(Num).Team = 3
                    If theteam$ = "green" Then Players(Num).Team = 4

                    If theteam$ = "spectator" Then Players(Num).Team = 6
                    
                    
                    Players(Num).UniqueID = thewon$
                    Players(Num).LastEvent = Now
                    
                End If
                
                
                'connected
                
                ' connected, address "24.68.41.80:27005"
                If DebugMode Then LastCalled = "ReadLog - Interpretor - Connected"

                Var$ = " connected, address " + Chr(34)
                If LeftR(stuff$, Len(Var$)) = Var$ Then
                    'extract IP and PORT
                    
                    e = InStr(1, stuff$, Chr(34))
                    f = InStr(e + 1, stuff$, Chr(34))
                    
                    If e > 0 And f > e Then
                    
                        'extract
                        
                        bth$ = Mid(stuff$, e + 1, f - e - 1)
                        
                        e = InStr(1, bth$, ":")
                        
                        IP$ = Left(bth$, e - 1)
                        Prt$ = Right(bth$, Len(bth$) - e)
                
                        If CheckBit2(General.Flags, 8) Then
                            For i = 1 To NumPlayers
                                If Players(i).UniqueID = thewon$ And thewon$ <> "0" Then
                                    ' kick this person
                                    SendActualRcon "kick # " & Players(i).UserID
                                End If
                            Next i
                        End If
                        
                
                        'check other
                        If Num = 0 Then
                            For i = 1 To NumPlayers
                                If Players(i).IP = IP$ And Players(i).UniqueID = thewon$ And Players(i).Name = thename$ Then Num = i: Exit For
                            Next i
                        End If
                        If Num = 0 Then
                            'not found, so make a new one for him
                            NumPlayers = NumPlayers + 1
                            Num = NumPlayers
                            Players(Num).RealName = ""
                            ReDim Players(Num).KillsWith(0 To NumKills)
                            Players(Num).RemoveMe = False
                            Players(Num).EntryName = ""
                        End If
                        
                        Players(Num).Class = -2
                        Players(Num).UniqueID = thewon$
                        Players(Num).UserID = theuserid$
                        Players(Num).Team = 0
                        Players(Num).Name = thename$
                        Players(Num).IP = IP$
                        Players(Num).Port = Val(Prt$)
                        Players(Num).ConnectOnly = True
                        Players(Num).MessInMin = 0
                        Players(Num).BroadcastType = 0
                        Players(Num).NumKickVotes = 0
                        Players(Num).Pos.X = 0
                        Players(Num).Pos.Y = 0
                        Players(Num).Pos.Z = 0
                        Players(Num).ShutUp = False
                        Players(Num).Warn = 0
                        Players(Num).Points = 0
                        Players(Num).TimeJoined = Now
                        Players(Num).LastEvent = Now
                        For kk = 1 To 10
                            Players(Num).LastMsgs(kk).Text = ""
                            Players(Num).LastMsgs(kk).When = 0
                        Next kk
                        
                        DoPlayerScan
                        DoRealNameTempCheck CInt(Num), True
                        gcs = CheckPlayerName(Num)
                        
                        'Let a script handle it
                        ExecFunctionScript "spec_connect", 1, theuserid$
                        
                        For i = 1 To NumConnectUsers
                            c = ConnectUsers(i).UserNum
                            d = ConnectUsers(i).LogLevel
                            
                            If CheckBit(c, 10) Then
                                If CheckBit2(d, 5) Then SendPacket "TY", " <> " + Players(Num).Name + " connected.", ConnectUsers(i).Index
                            End If
                        Next i
                        'UpdatePlayerList
                        
                        'make sure that there isnt already a person with this WONID playing if not supposed to.
                        
                        
                        
                    End If
                End If
                
                'joining
                'L 03/12/2001 - 15:35:43: "[BD]Avatar-X [DRP]<2><392640><>" entered the game
                If DebugMode Then LastCalled = "ReadLog - Interpretor - Entered"
                Var$ = " entered the game"
                If LeftR(stuff$, Len(Var$)) = Var$ Then
                
                    'check if hes there
                    If Num = 0 Then
                        'not found, so make a new one for him
                        NumPlayers = NumPlayers + 1
                        Num = NumPlayers
                        Players(Num).IP = ""
                        Players(Num).Port = 0
                        Players(Num).RemoveMe = False
                        Players(Num).EntryName = ""
                        ReDim Players(Num).KillsWith(0 To NumKills)
                    End If
                    
                    Players(Num).Class = -2
                    Players(Num).UserID = theuserid$
                    Players(Num).MessInMin = 0
                    Players(Num).Pos.X = 0
                    Players(Num).Pos.Y = 0
                    Players(Num).Pos.Z = 0
                    Players(Num).ShutUp = False
                    Players(Num).Warn = 0
                    Players(Num).Points = 0
                    Players(Num).BroadcastType = 0
                    Players(Num).NumKickVotes = 0
                    Players(Num).LastEvent = Now
                    
                    For kk = 1 To 10
                        Players(Num).LastMsgs(kk).Text = ""
                        Players(Num).LastMsgs(kk).When = 0
                    Next kk
                    
                    DoPlayerScan True
                    DoRealNameTempCheck CInt(Num), False
                    Players(Num).ConnectOnly = False
                    If Players(Num).EntryName = "" Then Players(Num).EntryName = thename$
                    gcs = CheckPlayerName(Num, True)
                    'Let a script handle it
                    ExecFunctionScript "spec_enter", 1, theuserid$
                    
                    UpdatePlayerList
                    AddToHTMLFile "", 6, 0, 0, thename$, ""
                    For i = 1 To NumConnectUsers
                        c = ConnectUsers(i).UserNum
                        d = ConnectUsers(i).LogLevel
                        
                        If CheckBit(c, 10) Then
                            If CheckBit2(d, 5) Then SendPacket "TY", " <> " + Players(Num).Name + " entered the game.", ConnectUsers(i).Index
                        End If
                    Next i
                End If
                
                
                'L 03/12/2001 - 15:35:46: "[BD]Avatar-X [DRP]<2><392640><Blue>" changed role to "Scout"
                If DebugMode Then LastCalled = "ReadLog - Interpretor - Changed class"
                Var$ = " changed role to " + Chr(34)
                If LeftR(stuff$, Len(Var$)) = Var$ Then
                
                    e = InStr(1, stuff$, Chr(34))
                    f = InStr(e + 1, stuff$, Chr(34))
                    
                    If e > 0 And f > e Then
                    
                        'extract
                        
                        Cl$ = Mid(stuff$, e + 1, f - e - 1)
                        
                        Cl$ = UCase(Cl$)
                        
                        clss = -2
                        'identify the class
                        If Cl$ = "SCOUT" Then clss = 1
                        If Cl$ = "SNIPER" Then clss = 2
                        If Cl$ = "SOLDIER" Then clss = 3
                        If Cl$ = "DEMOMAN" Then clss = 4
                        If Cl$ = "MEDIC" Then clss = 5
                        If Cl$ = "HWGUY" Then clss = 6
                        If Cl$ = "PYRO" Then clss = 7
                        If Cl$ = "SPY" Then clss = 8
                        If Cl$ = "ENGINEER" Then clss = 9
                        If Cl$ = "RANDOMPC" Then clss = 0
                        If Cl$ = "CIVILIAN" Then clss = -1
                        
                        'Let a script handle it
                        ExecFunctionScript "spec_changeclass", 2, theuserid$, Ts(clss)
                    
                        If Num > 0 Then
                            'player has been found, updating class
                            Players(Num).Class = clss
                            Players(Num).ConnectOnly = False
                            
                            
                            AddToHTMLFile Cl$, 11, Players(Num).Team, 0, Players(Num).Name, ""
                            UpdatePlayerList
                            For i = 1 To NumConnectUsers
                                c = ConnectUsers(i).UserNum
                                d = ConnectUsers(i).LogLevel
                                
                                If CheckBit(c, 10) Then
                                    If CheckBit2(d, 5) Then SendPacket "TY", " <> " + Players(Num).Name + " changed class to " + Cl$ + ".", ConnectUsers(i).Index
                                End If
                            Next i
                        End If
                    End If
                End If
                
                
                'L 03/12/2001 - 15:35:45: "[BD]Avatar-X [DRP]<2><392640><SPECTATOR>" joined team "Blue"
                If DebugMode Then LastCalled = "ReadLog - Interpretor - Joined team "
                Var$ = " joined team " + Chr(34)
                If LeftR(stuff$, Len(Var$)) = Var$ Then
                
                    e = InStr(1, stuff$, Chr(34))
                    f = InStr(e + 1, stuff$, Chr(34))
                    
                    If e > 0 And f > e Then
                    
                        'extract
                        
                        tm$ = LCase(Mid(stuff$, e + 1, f - e - 1))
                        
                        If tm$ = LCase(TeamNames(1)) Then Players(Num).Team = 1
                        If tm$ = LCase(TeamNames(2)) Then Players(Num).Team = 2
                        If tm$ = LCase(TeamNames(3)) Then Players(Num).Team = 3
                        If tm$ = LCase(TeamNames(4)) Then Players(Num).Team = 4
                        If tm$ = "spectator" Then Players(Num).Team = 6
                        
                        'Let a script handle it
                        ExecFunctionScript "spec_jointeam", 2, theuserid$, CInt(Players(Num).Team)
                        
                        AddToHTMLFile tm$, 12, Players(Num).Team, 0, Players(Num).Name, ""
                        
                        UpdatePlayerList
                        For i = 1 To NumConnectUsers
                            c = ConnectUsers(i).UserNum
                            d = ConnectUsers(i).LogLevel
                            
                            If CheckBit(c, 10) Then
                                If CheckBit2(d, 5) Then SendPacket "TY", " <> " + Players(Num).Name + " changed to team " + tm$ + ".", ConnectUsers(i).Index
                            End If
                        Next i
                    End If
                End If
                
                'L 03/13/2001 - 01:08:41: "[BD]Avatar-X [DRP]<1><4294967295><Blue>" changed name to "hey"
                If DebugMode Then LastCalled = "ReadLog - Interpretor - changed name"
                Var$ = " changed name to " + Chr(34)
                If LeftR(stuff$, Len(Var$)) = Var$ Then
                
                    e = InStr(1, stuff$, Chr(34))
                    f = InStr(e + 1, stuff$, Chr(34))
                    
                    If e > 0 And f > e Then
                    
                        'extract
                        
                        nwname$ = Mid(stuff$, e + 1, f - e - 1)
                        
                        If Num > 0 Then
                            'player has been found, updating team
                            Players(Num).Name = nwname$
                            
                            DoPlayerScan
                            
                            ExecFunctionScript "spec_changename", 3, theuserid$, thename$, nwname$

                            For i = 1 To NumConnectUsers
                                c = ConnectUsers(i).UserNum
                                d = ConnectUsers(i).LogLevel
                                
                                If CheckBit(c, 10) Then
                                    If CheckBit2(d, 5) Then SendPacket "TY", " <> " + thename$ + " changed name to " + nwname$, ConnectUsers(i).Index
                                End If
                            Next i
                            
                            'finally, update the list
                            gcs = CheckPlayerName(Num)
                            AddToHTMLFile "", 10, Players(Num).Team, Players(Num).Team, thename$, nwname$
                            UpdatePlayerList
                        End If
                    End If
                End If
                
                
                
                Var$ = " say " + Chr(34)
                Var2$ = " say_team " + Chr(34)
                If DebugMode Then LastCalled = "ReadLog - Interpretor - Say"
                If LeftR(stuff$, Len(Var$)) = Var$ Or LeftR(stuff$, Len(Var2$)) = Var2$ Then
                
                    e = InStr(1, stuff$, Chr(34))
                    f = InStrRev(stuff$, Chr(34))
                    
                    If e > 0 And f > e Then
                        
                        Sy$ = Mid(stuff$, e + 1, f - e - 1)
                    
                        mde = 0 'team mode
                        If InStr(1, UCase(stuff$), UCase(Var2$)) Then mde = 1
                        
                        If Num > 0 Then
                            If mde = 0 Then AddToHTMLFile Sy$, 1, Players(Num).Team, 0, Players(Num).Name, ""
                            If mde = 1 Then AddToHTMLFile Sy$, 2, Players(Num).Team, 0, Players(Num).Name, ""
                        
                            'now see if any basic log people need this
                            
                            ttm = Players(Num).Team
                            If ttm = 6 Then ttm = 0
                            SendToWatchers theuserid$, (thename$), Sy$, mde, ttm
                        
                            sy2$ = Sy$
                            If MapVoteTimer > 0 Then TrackVote theuserid$, sy2$
                            If KickVoteTimer > 0 Then TrackKickVote theuserid$, sy2$
                            If ChooseVoteTimer > 0 Then TrackChooseVote theuserid$, sy2$
                            
                            'Let the script called "serversay" handle this too...
                            ExecFunctionScript "spec_serversay", 4, Players(Num).Name, theuserid$, Ts(mde), Sy$
                            
                            e = InStr(1, Sy$, " ")
                            If e > 1 Then
                                cm$ = "saycmd_" + LeftR(Sy$, e - 1)
                                If e < Len(Sy$) Then rst$ = Right(Sy$, Len(Sy$) - e)
                            ElseIf Sy$ <> "" Then
                                rst$ = ""
                                cm$ = "saycmd_" + Sy$
                            End If
                            ExecFunctionScript LCase(cm$), 5, Players(Num).Name, theuserid$, Ts(mde), Sy$, rst$
                            
                            'see if theres a command here
                            
                            'see if the player needs to be kicked for spamming
                            CheckSpeechSwear Num, Sy$
                            CheckSpam theuserid$
                            CheckBroadcast Num, Sy$, 0
                        End If
    
                        If Len(Sy$) > 5 Then
                            If LCase(Left(Sy$, 5)) = "admin" Then
                                AdminSpeech Sy$, thename$, theuserid$
                            End If
                        End If
                    End If
                End If
                
                If DebugMode Then LastCalled = "ReadLog - Interpretor - disconnect"
                Var$ = " disconnected"
                If LeftR(stuff$, Len(Var$)) = Var$ Then
                
                    If Num > 0 Then
                        AddToHTMLFile "", 7, Players(Num).Team, 0, Players(Num).Name, ""
                       
                        'player has been found, updating team
                        If Not Players(Num).RemoveMe Then
                            ExecFunctionScript "spec_disconnect", 1, theuserid$
                            
                            For i = 1 To NumConnectUsers
                                c = ConnectUsers(i).UserNum
                                d = ConnectUsers(i).LogLevel
                                
                                If CheckBit(c, 10) Then
                                    If CheckBit2(d, 5) Then SendPacket "TY", " <> " + Players(Num).Name + " disconnected.", ConnectUsers(i).Index
                                End If
                            Next i
                            
                            ' if there is a kickvote against him that is pending, and he is trying to avoid the kickvote, ban him.
                            If KickVoteUser = Val(theuserid$) Then
                            
                                If KickVoteTimer > 0 And General.BanTime > 0 Then
                                    
                                    KickVoteTimer = 0
                                    SendRCONCommand "banid " + Ts(General.BanTime) + " " + Players(Num).UniqueID + " kick"
                            
                                    AddToLogFile "KICKVOTE: Player " + Players(Num).Name + " (UNIQUE: " + Players(Num).UniqueID + ") banned for " + Ts(General.BanTime) + " minutes, for attempting to avoid a kick vote, started by " + KickVoteStarterName + " (UIN " + KickVoteStarterUIN + ")"
                                    LastTalk = 1
                                    SendRCONCommand "say " + General.AutoAdminName + " Player " + Players(Num).Name + " banned for " + Ts(General.BanTime) + " minutes,", , 1
                                    SendRCONCommand "say " + General.AutoAdminName + " for trying to avoid the kick vote against him! "
                                End If
                            End If
                            RemovePlayer Num
                            UpdatePlayerList
                        End If
                    End If
                End If
                
                'L 03/13/2001 - 11:55:50: "Guy Lamar<7><4294967295><Blue>" killed "Stimpy<1><0><Red>" with "supershotgun"
                If DebugMode Then LastCalled = "ReadLog - Interpretor - killed"
                Var$ = " killed " + Chr(34)
                If LeftR(stuff$, Len(Var$)) = Var$ Then
                
                    e = InStr(1, stuff$, Chr(34))
                    f = InStr(e + 1, stuff$, Chr(34))
                    
                    If e > 0 And f > e Then
                        
                        who$ = Mid(stuff$, e + 1, f - e - 1)
                        
                        
                        'get team.
                        G = InStrRev(who$, ">", -1)
                        h = InStrRev(who$, "<", G)
                        
                        If h > 0 And G > h Then
                        
                            'extract team
                            
                            kilteam$ = LCase(Mid(who$, h + 1, G - h - 1))
                            
                            'now get wonID.
                            
                            G = InStrRev(who$, ">", h)
                            h = InStrRev(who$, "<", G)
                            
                            If h > 0 And G > h Then
                            
                                'extract wonid
                                
                                kilwon$ = Mid(who$, h + 1, G - h - 1)
                                
                                'now get USERID.
                                
                                G = InStrRev(who$, ">", h)
                                h = InStrRev(who$, "<", G)
                                
                                If h > 0 And G > h Then
                                
                                    'extract userid
                                    
                                    kiluserid$ = Mid(who$, h + 1, G - h - 1)
                                    
                                    'now get NAME.
                                    
                                    kilname$ = Mid(who$, 1, h - 1)
                        
                                    
                                    'LAST, Get the WEAPON!
                                    
                                    
                                    e = InStr(f + 1, stuff$, Chr(34))
                                    f = InStr(e + 1, stuff$, Chr(34))
                                    wep$ = Mid(stuff$, e + 1, f - e - 1)
                                    
                                    'all set!
                        
                                    wp = FindKillsWith(wep$)
                                    
                                    'get first and second player numbers
                                    
                                    
                                    num2 = FindPlayer(kiluserid$)
                                    
                                    ExecFunctionScript "spec_kill", 3, theuserid$, kiluserid$, wep$
                            
                                    'add it
                                    If Num > 0 And num2 > 0 Then
                                        AddToHTMLFile wep$, 5, Players(Num).Team, Players(num2).Team, Players(Num).Name, Players(num2).Name
                                    End If
                                    
                                    For i = 1 To NumConnectUsers
                                        c = ConnectUsers(i).UserNum
                                        d = ConnectUsers(i).LogLevel
                                        
                                        If CheckBit(c, 10) Then
                                            If CheckBit2(d, 2) Then SendPacket "TY", " -- " + Players(Num).Name + " killed " + Players(num2).Name + " with " + wep$, ConnectUsers(i).Index
                                        End If
                                    Next i
                                    
                                    If wp > 0 And Num > 0 And num2 > 0 Then
                                        Players(num1).KillsWith(wp) = Players(num1).KillsWith(wp) + 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                
'
'                Var$ = " connected, address " + Chr(34)
'                If LeftR(stuff$, Len(Var$)) = Var$ Then
'
'                End If
'
'
'
'                Var$ = " connected, address " + Chr(34)
'                If LeftR(stuff$, Len(Var$)) = Var$ Then
'
'                End If
'
'
'
'                Var$ = " connected, address " + Chr(34)
'                If LeftR(stuff$, Len(Var$)) = Var$ Then
'
'                End If
                
                                                
            End If
        End If
    End If
End If


'
'If DebugMode Then LastCalled = "ReadLog - Connected Section"
'
''now see if anyone is joining
''"[WP]Mimic<33><WON:14070149>" has entered the game
'
'If DebugMode Then LastCalled = "ReadLog - Entered Section"
'
''now it's time to find out what they are playing as!
''"[DoH]-Crowbar<12>" changed class to "Engineer"
'
'If DebugMode Then LastCalled = "ReadLog - Changed Class Section"
'
''now it's time to find out what team they are on
''"[DoH]-Crowbar<12>" joined team "2".
''"[ts][cfc]SilkySaddam<12>" changed to team "1".
'
'If DebugMode Then LastCalled = "ReadLog - Joined Team Section"
'
''Now, watch for a name change:
''"[AR]Avatar-X [BD]<71>" changed name to "[AR]Avatar-X [BD]-O<71>"
'
'If DebugMode Then LastCalled = "ReadLog - Name Change Section"
'
''Now, watch for a disconnect:
''"whoami<66>" disconnected
'
'If DebugMode Then LastCalled = "ReadLog - Disconnect"
'
'If DebugMode Then LastCalled = "ReadLog - Say Section"
'
'
''TELL commands:
''"[BD]Avatar-X [DRP]<119>" tell "[BD]Avatar-X" yo
'
If DebugMode Then LastCalled = "ReadLog - Tell Section"

B$ = Lg$
Var$ = ">" + Chr(34) + " tell "
If Len(B$) > Len(Var$) Then
    If InStr(1, UCase(B$), UCase(Var$)) Then
        
        'Need 3 pieces of information:
        '-the userid
        '-the message
        '-the name
        
        e = InStr(1, B$, Chr(34))
        G = InStr(e + 1, B$, Chr(34))
        
        If G > 0 And e > 0 Then
                                
            f = InStrRev(B$, "<", G)
            h = InStr(f + 1, B$, ">")
            
            If f > 0 And h > 0 Then
                    
                'got the userid
                n$ = Mid(B$, f + 1, h - f - 1)
                'get the name
                nm$ = Mid(B$, e + 1, f - e - 1)
                
                'now get the message
                
                e = InStr(G + 1, B$, "tell") + 4
                f = Len(B$) + 1
                    
                If f > 0 And e > 0 Then
                                
                    Num = FindPlayer(n$)
                                
                    'got the message
                    Sy$ = Mid(B$, e + 1, f - e - 1)
                   
                    If Num > 0 Then
                    
                        'now see if any basic log people need this
                        SendToWatchers "TELL", (nm$), Sy$, mde, Players(Num).Team
                    
                        sy2$ = Sy$
                        
                        ExecFunctionScript "spec_tell", 3, Players(Num).Name, n$, Sy$
                    
                    End If
                End If
            End If
        End If
    End If
End If

'teleport
'format:    PVTELE: Added teleporter at "12, 12, 12" angle "360" name "joe"

If DebugMode Then LastCalled = "ReadLog - Teleport"

B$ = Lg$
Var$ = "PVTELE: Added teleporter"
If Len(B$) > Len(Var$) Then
    If InStr(1, UCase(B$), UCase(Var$)) Then
        
        'Need 3 pieces of information:
        '-the co-ords
        '-the angle
        '-the name
        
        e = InStr(1, B$, Chr(34))
        G = InStr(e + 1, B$, Chr(34))
        
        If G > 0 And e > 0 Then
                   
            'got the coords
            coords$ = Mid(B$, e + 1, G - e - 1)
            
            'now get the angle
            
            e = InStr(G + 1, B$, Chr(34))
            G = InStr(e + 1, B$, Chr(34))
            
            If G > 0 And e > 0 Then
                       
                'got the angle
                Angle$ = Mid(B$, e + 1, G - e - 1)
            
                'now get the name
                
                e = InStr(G + 1, B$, Chr(34))
                G = InStr(e + 1, B$, Chr(34))
                
                If G > 0 And e > 0 Then
        
                    nm$ = Trim(Mid(B$, e + 1, G - e - 1))
                    
                    Num = FindTele(nm$)
                    
                    If Num = 0 And nm$ <> "" Then 'safe to add
                        NumTele = NumTele + 1
                        ReDim Preserve Tele(0 To NumTele)
                    
                        Tele(NumTele).Name = nm$
                        Tele(NumTele).Angle = Val(Angle$)
                        
                        'get co-ords out of co-ords var
                        e = InStr(1, coords$, ",")
                        G = InStr(e + 1, coords$, ",")
                        Tele(NumTele).X = Val(Trim(LeftR(coords$, e - 1)))
                        Tele(NumTele).Y = Val(Trim(Mid(coords$, e + 1, G - e - 1)))
                        Tele(NumTele).Z = Val(Trim(RightR(coords$, Len(coords$) - G)))
                        AddToLogFile "TELEADD: Teleporter " + nm$ + " added to map " + Vars.Map
                    End If
                End If
            End If
        End If
    End If
End If

'SACMD: Player "<id>" used "<command>"

If DebugMode Then LastCalled = "ReadLog - PV Command"

B$ = Lg$
Var$ = "SACMD: Player " + Chr(34)
If Len(B$) > Len(Var$) Then
    If InStr(1, UCase(B$), UCase(Var$)) Then
        
        'Need 2 pieces of information:
        '-the id
        '-the command
                
        e = InStr(1, B$, Chr(34))
        G = InStr(e + 1, B$, Chr(34))
        
        If G > 0 And e > 0 Then
                   
            'got the coords
            us$ = Mid(B$, e + 1, G - e - 1)
            
            'now get the angle
            
            e = InStr(G + 1, B$, Chr(34))
            G = InStr(e + 1, B$, Chr(34))
            
            If G > 0 And e > 0 Then
                    
                'got the command
                cmdd$ = Mid(B$, e + 1, G - e - 1)
            
                'send it to the script
                ExecFunctionScript "spec_consolecmd", 2, us$, cmdd$

                
            End If
        End If
    End If
End If


If DebugMode Then LastCalled = "ReadLog - User Pos"

B$ = Lg$
Var$ = "SAPOS: "
If Len(B$) > Len(Var$) Then
    If InStr(1, UCase(B$), UCase(Var$)) Then
        
        'Decodes a list of user position data
        'Sample line:
        ' SAPOS: 1=22,22,22&2=33,33,33
       
        G = InStr(1, B$, " ")
       
        'Do
        'do  'replaced DO with FOR
        For jkk = 1 To 10000000
            f = G
            e = InStr(G + 1, B$, "=")
            G = InStr(e + 1, B$, "&")
            
            If G = 0 And e > 0 Then G = Len(B$) + 1
            
            If G > 0 And e > 0 And f > 0 Then
                       
                'get the userid
                us$ = Mid(B$, f + 1, e - f - 1)
                                
                'Get the Coords
                coords$ = Mid(B$, e + 1, G - e - 1)
                
                'Find this player
                Num = FindPlayer(us$)
                
                If Num > 0 Then
                    'get co-ords out of co-ords var
                    If DebugMode Then LastCalled = "ReadLog - User Pos - Section 1"
                    
                    e1 = InStr(1, coords$, ",")
                    G1 = InStr(e1 + 1, coords$, ",")
                    
                    If DebugMode Then LastCalled = "ReadLog - User Pos - Section 2"

                    X2 = Players(Num).Pos.X
                    Y2 = Players(Num).Pos.Y
                    z2 = Players(Num).Pos.Z
                    
                    If DebugMode Then LastCalled = "ReadLog - User Pos - Section 3"
                    
                    X = CInt(Trim(LeftR(coords$, e1 - 1)))
                    Y = CInt(Trim(Mid(coords$, e1 + 1, G1 - e1 - 1)))
                    Z = CInt(Trim(RightR(coords$, Len(coords$) - G1)))
                                       
                    Players(Num).Pos.X = X
                    Players(Num).Pos.Y = Y
                    Players(Num).Pos.Z = Z
                    
                    If DebugMode Then LastCalled = "ReadLog - User Pos - Section 4"
                    
                    If X2 <> Players(Num).Pos.X Or Y2 <> Players(Num).Pos.Y Or z2 <> Players(Num).Pos.Z Then
                        Players(Num).LastEvent = Now
                    End If
                    'Set this value in MapArray
                    
                    If DebugMode Then LastCalled = "ReadLog - User Pos - Section 5"
                    
                    'Map Data Format:
                    ' -4096   to   4096 -> more than one team / old format
                    '  4097   to  12288 -> blue team   (norm: -8192)
                    ' 12289   to  20480 -> red team    (norm: -16384)
                    ' -4097   to -12288 -> yellow team (norm: +8192)
                    '-12289   to -20480 -> green team  (norm: +16384)
                    ' 20481   to  28972 -> more than one team    (norm: -24576)

                    X1 = Int((Players(Num).Pos.X / 128) + 32)
                    Y1 = Int((Players(Num).Pos.Y / 128) + 32)
                    z1 = Players(Num).Pos.Z
                    
                    t1 = 0
                    If Players(Num).Team = 1 Then t1 = 8192
                    If Players(Num).Team = 2 Then t1 = 16384
                    If Players(Num).Team = 3 Then t1 = -8192
                    If Players(Num).Team = 4 Then t1 = -16384
                                      
                    'See if this is a dual-team point.
                    t2 = GetMapArrayTeam(MapArray(X1, Y1))
                    
                    If t2 = 5 Then t1 = 24675
                    If t2 <> 5 And t2 <> 0 And t2 <> Players(Num).Team Then t1 = 24675 'This point is NOT a one-team point.
                    
                    z1 = z1 + t1
                    MapArray(X1, Y1) = z1
                End If
            End If
    

            If e = 0 Then Exit For
        Next jkk
        If jkk = 10000000 Then AddToLogFile "!!! ERROR !!! INFINITE LOOP: Occured at " + LastCalled
'        Loop Until e = 0
        
        If DebugMode Then LastCalled = "ReadLog - User Pos - Section 6"
        UpdatePlayerList
    End If
End If



' AV PLAYER INFO!!

'PLAYERINFO: 20  |   253   |  255  |  0   |    0   |   0   |    NO    |   NO  |  NO
'Format: ENT NUM | USER ID | CLASS | TEAM | HEALTH | ARMOR | DEVOICED | ADMIN | FROZEN
'                    1^       2^      3^      4        5       6^


If DebugMode Then LastCalled = "ReadLog - PlayerInfo"

B$ = Lg$
Var$ = "PLAYERINFO: "
If Len(B$) > Len(Var$) Then
    If Left(UCase(B$), Len(Var$)) = UCase(Var$) Then
        
        'Need 4 pieces of information:
        '-the id
        '-the class
        '-the team
        '-shutup = true/false
                
        f = 0
        n = 0
        f = InStr(1, B$, "|")
        'do  'replaced DO with FOR
        For jkk = 1 To 10000000

            e = f
            f = InStr(e + 1, B$, "|")
            
            
            If e > 0 And f > e Then
                       
                n = n + 1
                'got the data
                If n = 1 Then
                    us$ = Trim(Mid(B$, e + 1, f - e - 1))
                    Num = FindPlayer(us$)
                Else
                    dat$ = Trim(Mid(B$, e + 1, f - e - 1))
                    If Num > 0 Then
                        If n = 2 And Val(dat$) <> 255 Then Players(Num).Class = Val(dat$)
                        If n = 3 Then Players(Num).Team = Val(dat$)
                        If n = 6 And dat$ = "NO" Then Players(Num).ShutUp = False
                        If n = 6 And dat$ = "YES" Then Players(Num).ShutUp = True
                        
                        If Players(Num).Class > 0 And Players(Num).Team > 0 Then Players(Num).ConnectOnly = False
                        
                    End If
                End If
            End If
            If f = 0 Then Exit For
        Next jkk
        If jkk = 10000000 Then AddToLogFile "!!! ERROR !!! INFINITE LOOP: Occured at " + LastCalled
        
    End If
End If


'PV VOTE STUFF

'VOTE: 1 registered for 10.
'VOTE: 5 registered for 243.  (OPTION#, SERVERID)
'VOTE: Completed.

If DebugMode Then LastCalled = "ReadLog - Vote Tracking"

B$ = Lg$
Var$ = "VOTE: "
Var2$ = "registered for"
If Len(B$) > Len(Var$) Then  ' use menu kickvoting is enabled
    If Left(UCase(B$), Len(Var$)) = UCase(Var$) And InStr(1, UCase(B$), UCase(Var2$)) Then
        'If
        
        'Need 2 pieces of information:
        '-the id
        '-the vote
                
        e = InStr(1, B$, ":")
        f = InStr(e + 1, B$, "registered")
        
        If e > 0 And f > e Then
                   
            vt$ = Trim(Mid(B$, e + 1, f - e - 1))
            
            e = InStr(f + 1, B$, "for")
            f = InStr(e + 1, B$, ".")
            
            If e > 0 And f > e Then
                
                e = e + 2
                us$ = Trim(Mid(B$, e + 1, f - e - 1))
                Num = FindPlayer(us$)
                       
                If Num > 0 Then
                    If KickVoteTimer > 0 And CheckBit2(General.Flags, 3) Then
                        If Val(vt$) = 1 Then TrackKickVote us$, "yes"
                        If Val(vt$) = 2 Then TrackKickVote us$, "no"
                    End If
                    If ChooseVoteTimer > 0 And CheckBit2(General.Flags, 4) Then
                        If Val(vt$) <= UBound(ChooseVote) And Val(vt$) >= 0 Then
                            TrackChooseVote us$, ChooseVote(Val(vt$))
                        End If
                    End If
                End If
            End If
        End If
    End If
End If


'world kills
'format:    "_+*Chicka*+_<21>" killed by world with "train"
If DebugMode Then LastCalled = "ReadLog - World Kills"

B$ = Lg$
Var$ = ">" + Chr(34) + " killed by world with " + Chr(34)
If Len(B$) > Len(Var$) Then
    If InStr(1, UCase(B$), UCase(Var$)) Then
        
        'Need 3 pieces of information:
        '-the first userid
        '-the first name
        '-the weapon
        
        e = InStr(1, B$, Chr(34))
        G = InStr(e + 1, B$, Chr(34))
        
        If G > 0 And e > 0 Then
                                
            f = InStrRev(B$, "<", G)
            h = InStr(f + 1, B$, ">")
            
            If f > 0 And h > 0 And f > (e + 1) Then
                    
                'got the first userid
                us1$ = Mid(B$, f + 1, h - f - 1)
                'got the first name
                nm1$ = Mid(B$, e + 1, f - e - 1)
                
               
                'now get the weapon name
                
                e = InStr(G + 1, B$, Chr(34))
                G = InStr(e + 1, B$, Chr(34))
                
                If G > 0 And e > 0 Then
        
                    wep$ = Mid(B$, e + 1, G - e - 1)
                    'search for weapon name...
                    
                    'get first and second player numbers
                    
                    num1 = FindPlayer(us1$)
                    
                    'add it
                    If num1 > 0 Then
                        AddToHTMLFile wep$, 5, Players(num1).Team, 0, Players(num1).Name, "<WORLD>"
                    End If
                    
                    For i = 1 To NumConnectUsers
                        c = ConnectUsers(i).UserNum
                        d = ConnectUsers(i).LogLevel
                        
                        If CheckBit(c, 10) Then
                            If CheckBit2(d, 2) Then SendPacket "TY", " -- " + Players(num1).Name + " killed by world with " + wep$, ConnectUsers(i).Index
                        End If
                    Next i
                    
                End If
            End If
        End If
    End If
End If

'L 11/09/2000 - 20:56:56:
                                
'self kills
'format:   "Crown<12>" killed self with "world"
If DebugMode Then LastCalled = "ReadLog - World Kills"

B$ = Lg$
Var$ = ">" + Chr(34) + " killed self with " + Chr(34)
If Len(B$) > Len(Var$) Then
    If InStr(1, UCase(B$), UCase(Var$)) Then
        
        'Need 3 pieces of information:
        '-the first userid
        '-the first name
        '-the weapon
        
        e = InStr(1, B$, Chr(34))
        G = InStr(e + 1, B$, Chr(34))
        
        If G > 0 And e > 0 Then
                                
            f = InStrRev(B$, "<", G)
            h = InStr(f + 1, B$, ">")
            
            If f > 0 And h > 0 And f > (e + 1) Then
                    
                'got the first userid
                us1$ = Mid(B$, f + 1, h - f - 1)
                'got the first name
                nm1$ = Mid(B$, e + 1, f - e - 1)
                
               
                'now get the weapon name
                
                e = InStr(G + 1, B$, Chr(34))
                G = InStr(e + 1, B$, Chr(34))
                
                If G > 0 And e > 0 Then
        
                    wep$ = Mid(B$, e + 1, G - e - 1)
                    'search for weapon name...
                    
                    'get first and second player numbers
                    
                    num1 = FindPlayer(us1$)
                    
                    'add it
                    If num1 > 0 Then
                        AddToHTMLFile wep$, 5, Players(num1).Team, 0, Players(num1).Name, "<SELF>"
                    End If
                    
                    For i = 1 To NumConnectUsers
                        c = ConnectUsers(i).UserNum
                        d = ConnectUsers(i).LogLevel
                        
                        If CheckBit(c, 10) Then
                            If CheckBit2(d, 2) Then SendPacket "TY", " -- " + Players(num1).Name + " killed self with " + wep$, ConnectUsers(i).Index
                        End If
                    Next i
                    
                End If
            End If
        End If
    End If
End If
                                
                                
If Server.GameMode = 2 Then
                                
    If DebugMode Then LastCalled = "ReadLog - Joined Team Section - CS"
                                
    B$ = Lg$
    Var$ = " is joining the CT force"
    If Len(B$) > Len(Var$) Then
        If InStr(1, UCase(B$), UCase(Var$)) Then
            
            'Need 1 piece of information:
            '-the name
            
            e = InStr(1, B$, Var$)
            
            If e > 0 Then
                                    
                'got the name
                n$ = Mid(B$, 1, e - 1)
                                
                Num = 0
                For i = 1 To NumPlayers
                    If LCase(Players(i).Name) = LCase(n$) Then Num = i: Exit For
                Next i
                                
        
                If Num > 0 Then
                    'player has been found, updating team
                    bb$ = "CT"
                    
                    'Let a script handle it
                    ExecFunctionScript "spec_jointeam", 2, us$, "2"
                    
                    AddToHTMLFile bb$, 12, Players(Num).Team, 0, Players(Num).Name, ""
                    Players(Num).Team = 2
                    Players(Num).ConnectOnly = False
                    UpdatePlayerList
                    For i = 1 To NumConnectUsers
                        c = ConnectUsers(i).UserNum
                        d = ConnectUsers(i).LogLevel
                        
                        If CheckBit(c, 10) Then
                            If CheckBit2(d, 5) Then SendPacket "TY", " <> " + Players(Num).Name + " joined the " + bb$ + " force.", ConnectUsers(i).Index
                        End If
                    Next i
                
                End If
            End If
        End If
    End If

    B$ = Lg$
    Var$ = " is joining the TERRORIST force"
    If Len(B$) > Len(Var$) Then
        If InStr(1, UCase(B$), UCase(Var$)) Then
            
            'Need 1 piece of information:
            '-the name
            
            e = InStr(1, B$, Var$)
            
            If e > 0 Then
                                    
                'got the name
                n$ = Mid(B$, 1, e - 1)
                                
                Num = 0
                For i = 1 To NumPlayers
                    If LCase(Players(i).Name) = LCase(n$) Then Num = i: Exit For
                Next i
                                
        
                If Num > 0 Then
                    'player has been found, updating team
                    bb$ = "Terrorist"
                    
                    'Let a script handle it
                    ExecFunctionScript "spec_jointeam", 2, us$, "1"
                    
                    AddToHTMLFile bb$, 12, Players(Num).Team, 0, Players(Num).Name, ""
                    Players(Num).Team = 1
                    UpdatePlayerList
                    For i = 1 To NumConnectUsers
                        c = ConnectUsers(i).UserNum
                        d = ConnectUsers(i).LogLevel
                        
                        If CheckBit(c, 10) Then
                            If CheckBit2(d, 5) Then SendPacket "TY", " <> " + Players(Num).Name + " joined the " + bb$ + " force.", ConnectUsers(i).Index
                        End If
                    Next i
                
                End If
            End If
        End If
    End If

    If DebugMode Then LastCalled = "ReadLog - Say Section - CS"

    B$ = Lg$
    Var$ = Chr(2)
    If Len(B$) > Len(Var$) Then
        If Left(B$, 1) = Var$ Then
            
            'Need 2 pieces of information:
            '-the message
            '-the name
            
            e = InStr(1, B$, Chr(2))
            G = InStr(e + 1, B$, " :    ")
            
            If G > 0 And e > 0 Then
                                    
                'get the name
                nm$ = Mid(B$, e + 1, G - e - 1)
                    
                'now get the message
                
                G = G + 5
                f = InStr(G + 1, B$, Chr(10))
                
                If f > G Then
                    'got the message
                    Sy$ = Mid(B$, G + 1, f - G - 1)
                    
                    mde = 0
                    If LeftR(nm$, 6) = "*DEAD*" Then nm$ = Right(nm$, Len(nm$) - 6): mde = 1
                    
                    Num = 0
                    For i = 1 To NumPlayers
                        If LCase(Players(i).Name) = LCase(nm$) Then Num = i: Exit For
                    Next i
                    
                    If Num > 0 Then
                        If mde = 0 Then AddToHTMLFile Sy$, 1, Players(Num).Team, 0, Players(Num).Name, ""
                        If mde = 1 Then AddToHTMLFile Sy$, 2, Players(Num).Team, 0, Players(Num).Name, ""
                    
                        'now see if any basic log people need this
                        SendToWatchers n$, (nm$), Sy$, mde, Players(Num).Team
                    
                        sy2$ = Sy$
                        If MapVoteTimer > 0 Then TrackVote Ts(Players(Num).UserID), sy2$
                        If KickVoteTimer > 0 Then TrackKickVote Ts(Players(Num).UserID), sy2$
                        If ChooseVoteTimer > 0 Then TrackChooseVote Ts(Players(Num).UserID), sy2$
                        
                        'Let the script called "serversay" handle this too...
                        ExecFunctionScript "spec_serversay", 4, Players(Num).Name, n$, Ts(mde), Sy$
                        e = InStr(1, Sy$, " ")
                        If e > 1 Then
                            cm$ = "saycmd_" + Left(Sy$, e - 1)
                            If e < Len(Sy$) Then rst$ = Right(Sy$, Len(Sy$) - e)
                        ElseIf Sy$ <> "" Then
                            rst$ = ""
                            cm$ = "saycmd_" + Sy$
                        End If
                        ExecFunctionScript LCase(cm$), 5, Players(Num).Name, n$, Ts(mde), Sy$, rst$
                        
                        'see if the player needs to be kicked for spamming
                        CheckSpeechSwear Num, Sy$
                        CheckSpam Ts(Players(Num).UserID)
                        CheckBroadcast Num, Sy$, 0
                    End If

                    If Len(Sy$) > 5 Then
                        If LCase(Left(Sy$, 5)) = "admin" Then
                            AdminSpeech Sy$, nm$, Ts(Players(Num).UserID)
                        End If
                    End If
                End If
            End If
        End If
    End If


End If

'L 11/09/2000 - 22:12:25: Precaching world (9.9 r. 0)...

If DebugMode Then LastCalled = "ReadLog - PV DLL VERSION"

B$ = Lg$
Var$ = "Precaching world ("
If Len(B$) > Len(Var$) Then
    If LeftR(UCase(B$), Len(Var$)) = UCase(Var$) Then
        

        e = InStr(1, B$, "(")
        f = InStr(e + 1, B$, ")")

        If f > e And e > 0 Then
            'get the userid
            PVVersion = Mid(B$, e + 1, f - e - 1)
        End If
    End If
End If



End Sub

Sub CheckBroadcast(Num, Sy$, mde)
sy2$ = Sy$

nn$ = Players(Num).Name

'See if this player's text needs to be broadcast
If Players(Num).BroadcastType = 1 Then

    'not fully supported yet...
    
'    SendRCONCommand "say <SPEC " + nn$ + "> " + sy2$
    DoTalk "<BROAD " + nn$ + "> " + sy2$

ElseIf Players(Num).BroadcastType = 2 Then
    SendRCONCommand "say " + nn$ + ": " + sy2$, 5
End If

End Sub

Sub CheckSpam(us$)
If DebugMode Then LastCalled = "CheckSpam"



'this player just talked... see if he talked too much
Dim NumMs As Integer

If General.MaxMsg = 0 Then General.MaxMsg = 10
If General.MaxTime = 0 Then General.MaxTime = 15

If Vars.ClanBattle = False And General.MaxMsg > 0 And General.MaxTime > 0 Then

    Num = FindPlayer(us$)

    If Num > 0 Then
        
        'Check if time has come to reset it.
        'Count how many messages in the last General.MaxTime seconds.
        
        NumMs = 0
        For i = 1 To 10
            sec = Timer - Players(Num).LastMsgs(i).When
            If sec <= General.MaxTime And Players(Num).LastMsgs(i).When > 0 Then NumMs = NumMs + 1
        Next i
        
        If NumMs > General.MaxMsg Then 'too many -- spam!
                   
            c2 = FindClan(Num)
            If c2 > 0 Then
                If CheckBit2(Clans(c2).Flags, 2) Then
                    Exit Sub
                End If
            End If
                   
            SendRCONCommand "kick # " + Ts(Players(Num).UserID)
            AddToLogFile "SPAMKICK: Player " + Players(Num).Name + " (UNIQUE: " + Players(Num).UniqueID + ") kicked for excessive message spam, " + Ts(NumMs) + " messages in less than " + Ts(General.MaxTime) + " seconds."
            SendRCONCommand "say " + General.AutoAdminName + " Player " + Players(Num).Name + " kicked for message spamming."
            If PointData.SpamKickCosts > 0 Then SetPoints Num, GetPoints(Num) - PointData.SpamKickCosts
            
        End If

'       OLD SPAMCHECK SYSTEM
'
'
'        'Check if time has come to reset it.
'        'sc = Second(Now - Players(Num).CountStart) + (Minute(Now - Players(Num).CountStart) * 60)
'        Players(Num).MessInMin = Players(Num).MessInMin + 1
'
'        If Players(Num).MessInMin > General.MaxMsg Then 'too many -- spam!
'
'            c2 = FindClan(Num)
'            If c2 > 0 Then
'                If CheckBit2(Clans(c2).Flags, 2) Then
'                    Exit Sub
'                End If
'            End If
'
'            SendRCONCommand "kick # " + Ts(Players(Num).UserID)
'            AddToLogFile "SPAMKICK: Player " + Players(Num).Name + " (UNIQUE: " + Players(Num).UniqueID + ") kicked for excessive message spam, more than " + Ts(Players(Num).MessInMin) + " messages in less than " + Ts(General.MaxTime) + " seconds."
'            SendRCONCommand "say " + general.AutoAdminName + " Player " + Players(Num).Name + " kicked for message spamming."
'            If PointData.SpamKickCosts > 0 Then SetPoints Num, GetPoints(Num) - PointData.SpamKickCosts
'            Players(Num).MessInMin = 0
'
'        End If
    End If
End If


End Sub

Sub GenerateMenu(Text As String, Duration As Integer, OptionList() As String, NumOptions As Integer)
If DebugMode = True Then LastCalled = "GenerateMenu"

'Creates an on-screen menu.
If DLLEnabled = False Then Exit Sub

SendActualRcon "sa_question " + ReadyForDLL(Text)
For i = 1 To NumOptions
    SendActualRcon "sa_option" + Ts(i) + " " + ReadyForDLL(OptionList(i))
    a$ = a$ + OptionList(i) + ", "
Next i
a$ = LeftR(a$, Len(a$) - 2)

SendToWatchers "SERVER", "", vbCrLf + "*-*-* MENU POP-UP: (" + Ts(Duration) + " sec), Question: " + Text + vbCrLf + "*-*-* MENU POP-UP: Options: " + a$, 0, 0

SendActualRcon "sa_duration " + Ts(Duration)
SendActualRcon SA_CHECK

End Sub


Sub AddRealPlayer(p$)
If DebugMode Then LastCalled = "AddRealPlayer"

'adds one real player to the list

f = 0
i = 0
'do  'replaced DO with FOR
For jkk = 1 To 10000000
    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            G = h
            h = InStr(G + 1, a$, Chr(250))
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                m$ = Mid(a$, G, h - G)
                
                If j = 1 Then nn$ = m$
                If j = 2 Then Un$ = m$
            End If
        Loop Until h = 0
    
    End If
    If f = 0 Or e = 0 Then Exit For
Next jkk
If jkk = 10000000 Then AddToLogFile "!!! ERROR !!! INFINITE LOOP: Occured at " + LastCalled
'Loop Until f = 0 Or e = 0

AddRealPlayer2 nn$, Un$

DoPlayerScan
SaveCommands

End Sub

Function AddRealPlayer2(nn$, Un$, Optional Prob As Boolean) As Boolean
If DebugMode Then LastCalled = "AddRealPlayer2"

For i = 1 To NumRealPlayers
    If RealPlayers(i).RealName = nn$ Then j = i: Exit For
Next i

'make sure id doesnt already exist
For i = 1 To NumRealPlayers
    If InStr(1, RealPlayers(i).UniqueID, Chr(34) + Un$ + Chr(34)) And CheckBit2(RealPlayers(j).Flags, 5) = False Then Exit Function
Next i

If j = 0 Then
    NumRealPlayers = NumRealPlayers + 1
    j = NumRealPlayers
    ReDim Preserve RealPlayers(0 To j)
    RealPlayers(j).UniqueID = Chr(34) + Un$ + Chr(34) + ";"
Else
    If InStr(1, RealPlayers(j).UniqueID, Chr(34) + Un$ + Chr(34) + ";") = 0 Then RealPlayers(j).UniqueID = RealPlayers(j).UniqueID + " " + Chr(34) + Un$ + Chr(34) + ";"
End If

RealPlayers(j).LastName = ""
RealPlayers(j).RealName = nn$
RealPlayers(j).LastTime = CDbl(Now)
If Prob = True Then RealPlayers(j).Flags = 2 ^ 5
If Prob = False And CheckBit2(RealPlayers(j).Flags, 5) Then RealPlayers(j).Flags = RealPlayers(j).Flags - 2 ^ 5

AddRealPlayer2 = True

End Function

Sub DoPlayerScan(Optional AddRealNow As Boolean)
If DebugMode Then LastCalled = "DoPlayerScan"

'scans player list for realnames

For i = 1 To NumPlayers
    RealNameSearch i, AddRealNow
Next i

UpdatePlayerList

End Sub

Function GetMapArrayTeam(z1) As Integer

' 20481   to  28972 -> more than one team    (norm: -24576)

t = 0
If z1 >= 4097 And z1 <= 12288 Then t1 = 1
If z1 >= 12289 And z1 <= 20480 Then t1 = 2
If z1 >= 20481 Then t1 = 5
If z1 <= -4097 And z1 >= -12288 Then t1 = 3
If z1 <= -12289 Then t1 = 4

GetMapArrayTeam = t1

End Function

Function GetMapArrayTeam2(z1) As Integer

' 20481   to  28972 -> more than one team    (norm: -24576)

t = 0
z2 = z1

If z1 >= 4097 And z1 <= 12288 Then t1 = 1: z2 = z2 - 8192
If z1 >= 12289 And z1 <= 20480 Then t1 = 2: z2 = z2 - 16384
If z1 >= 20481 Then t1 = 5: z2 = z2 - 24675
If z1 <= -4097 And z1 >= -12288 Then t1 = 3: z2 = z2 + 8192
If z1 <= -12289 Then t1 = 4: z2 = z2 + 16384

GetMapArrayTeam2 = z2

End Function

Sub ScanForAMAFiles()

' scans the half life directory for AMA files and then reads them in

Dim List(1 To 100) As String
Dim X() As String

a$ = Server.HLPath + "\"
B$ = Dir(a$ + "*.ama")

i = 1
Do While B$ <> "" And i < 99
    List(i) = B$
    B$ = Dir
    i = i + 1
Loop

'scan thru the list.
For j = 1 To i - 1

    X = Split(List(j), ".")
    
    'name of map only...
    c$ = X(0)
    
    ReadAMAFile a$ + List(j), c$
Next j




End Sub


Sub ReadAMAFile(Fle As String, MapBelong As String)

' opens a AMA file and a MAPDATA file and adds the AMA to the MAPDATA.
' If the AMA file is for the current map, it updates MAPARRAY

Dim MyMapArray(0 To 64, 0 To 64) As Integer ' Stores Z coord at this location, used for keeping map data

' first load tha map array in.

'If

f$ = Server.BothPath + "\Assistant\MapData"
If Dir(f$, vbDirectory) = "" Then MkDir f$
f$ = f$ + "\" + MapBelong + ".dat"

If CheckForFile(f$) Then
    'Load the data
    h = FreeFile
Close h
    Open f$ For Binary As h
        Get #h, , MyMapArray
    Close h
End If

' now read the AMA file.
' format:
'each line is one Y row and each character is the X..
' so in coords:
' (0,0) (0,1) (0,2)
' (1,0) (1,1) (1,2)
' (2,0) (2,1) (2,2)
' etc to 64 in each direction
' sample line..
' 0,0,0,0,0,100,0,0 ... 0,120,0

Dim InBuffer() As String

h = FreeFile
Y = 0
Open Fle For Input As h
    
    'AGHJhj!!!
    
    Do While Not EOF(h)
    
        '39,33,  39,32
        
    
        Line Input #h, a$
        
        ' convert to usable format
        
        InBuffer = Split(a$, ",")
        
        ' add to the array
           
        For X = 0 To 64
            
            '' this is part of the little client-side map thing... u know the mini-map in SAC.
            ' what i did is make the DLL keep track of where the players go, and then write that to a file
            ' this function read that from the file and integrates it with the already known map data
            ' its a bitch because in order to store team values I used different ranges of 1 integer.
            ' there are 6 ranges... 1 for no team, 4 for the 4 teams, and 1 meaning "more than one team"
            'yeah the DLL functions that are SA related like talking will come with SA.
            
            z1 = Val(InBuffer(X))
                    
            ' add the AMA data to the current map array
            
            t1 = GetMapArrayTeam(z1)
            t2 = GetMapArrayTeam(MyMapArray(X, Y))
            origz = z1
            
            t3 = 0
            If t1 = 1 Then t3 = 8192
            If t1 = 2 Then t3 = 16384
            If t1 = 3 Then t3 = -8192
            If t1 = 4 Then t3 = -16384
            If t1 = 5 Then t3 = 24675
            origz = origz - t3
                            
            If t2 = 5 And origz <> 0 Then t3 = 24675
            If t2 <> 5 And t2 <> 0 And t2 <> t1 And t1 <> 0 Then t3 = 24675 'This point is NOT a one-team point.
            
            origz = origz + t3
            If origz = 0 Then origz = MyMapArray(X, Y)
            
            If origz < 32000 And origz > -32000 Then MyMapArray(X, Y) = origz
        
        Next X
        Y = Y + 1
    Loop
Close h

'Save new map data.
If MapBelong <> "" Then
    f$ = Server.BothPath + "\Assistant" ' \MapData"
    If Dir(f$, vbDirectory) = "" Then MkDir f$
    f$ = Server.BothPath + "\Assistant\MapData"
    If Dir(f$, vbDirectory) = "" Then MkDir f$
    
    f$ = f$ + "\" + MapBelong + ".dat"
    If CheckForFile(f$) Then Kill f$
    
    'Save the data
    h = FreeFile
    Close h
    Open f$ For Binary As h
        Put #h, , MyMapArray
    Close h
End If


If LCase(MapBelong) = LCase(Vars.Map) Then
    ' copy over
    For Y = 0 To 64
        For X = 0 To 64
            
            MapArray(X, Y) = MyMapArray(X, Y)
            
        Next
    Next
End If

'now, delete the AMA file.
Kill Fle


End Sub

Sub AdminSpeech(Sy$, nm$, us$)
If DebugMode Then LastCalled = "AdminSpeech"

'Handles when people try to talk to the admin

'First check if the saying is already in the database
d$ = UCase(Sy$)

'special case:

'ADMIN KICK JOE

If Vars.ClanBattle = True Then Exit Sub

If LeftR(d$, 5) = "ADMIN" And Len(d$) > 5 Then
    dd$ = Trim(Right(d$, Len(d$) - 5))
    If LeftR(dd$, 4) = "KICK" And Len(dd$) > 4 Then
        dd$ = Trim(Right(dd$, Len(dd$) - 4))
        
        If Len(dd$) > 0 Then
            
            'kick vote
            
            StartAutoKickVote dd$, us$
            Exit Sub
        End If
    End If
End If

'ADMIN SEND MESSAGE TO VIPER

If LeftR(d$, 5) = "ADMIN" And Len(d$) > 5 Then
    dd$ = Trim(Right(d$, Len(d$) - 5))
    If LeftR(dd$, 15) = "SEND MESSAGE TO" And Len(dd$) > 15 Then
        dd$ = Trim(Right(dd$, Len(dd$) - 15))
        
        If Len(dd$) > 0 Then
            'Send message to an admin
            'Find the admin!
            
            If AdminICQTime > 0 Then
                If AdminICQId > 0 Then
                    SendRCONCommand "say " + General.AutoAdminName + " I am currently waiting for a response. Please wait."
                Else
                    SendRCONCommand "say " + General.AutoAdminName + " You must wait " + Ts(AdminICQTime) + " seconds before you can send another message."
                End If
            Else
                For i = 1 To NumUsers
                    If UCase(Users(i).Name) = dd$ Then j = i: Exit For
                Next i
                
                If j = 0 Then
                    SendRCONCommand "say " + General.AutoAdminName + " The Administrator " + dd$ + " could not be found."
                Else
                    If Users(j).ICQ = "" Or CheckBit(j, 32) = False Then
                        SendRCONCommand "say " + General.AutoAdminName + " " + Users(j).Name + " does not wish to recieve messages."
                    Else
                        SendRCONCommand "say " + General.AutoAdminName + " Please enter the message for " + Users(j).Name + ":"
                        AdminICQId = Val(us$)
                        AdminICQTime = 45
                        AdminIcqNum = j
                    End If
                End If
                    
            End If
            Exit Sub
        End If
    End If
End If
Dim NewScriptData As typScriptData

If CheckBit2(General.Flags, 2) = True Then Exit Sub 'no admin speech

'add to lastspeech
For i = 20 To 2 Step -1
    LastSpeech(i).UserID = LastSpeech(i - 1).UserID
Next i

LastSpeech(1).UserID = Val(us$)

Dim NumSpch As Integer

For i = 1 To 20
    If LastSpeech(i).UserID = Val(us$) Then
        NumSpch = NumSpch + 1
    End If
Next i

'Finally, see if this is too many.
If NumSpch >= General.MaxSpeech And General.MaxSpeech > 0 Then
    'kick
    Num = FindPlayer(us$)
    SendRCONCommand "kick # " + us$
    If Num > 0 Then
    
        c2 = FindClan(Num)
        If c2 > 0 Then
            If CheckBit2(Clans(c2).Flags, 2) Then
                Exit Sub
            End If
        End If
    
        AddToLogFile "SPEECHKICK: Player " + nm$ + " (UNIQUE: " + Players(Num).UniqueID + ") kicked, due to abusing the Admin Speech (spamming)"
        SendRCONCommand "say " + General.AutoAdminName + " " + nm$ + " was pissing me off, so I kicked him."
        If Players(Num).IP <> "" And Players(Num).Port > 0 Then
            Vars.UserIsTCP = False
            SendToUserIP Players(Num).IP, Players(Num).Port, "You have been kicked due because you were abusing the Automatic Administrator."
        End If
        Exit Sub
    End If
End If

'Find the lengthiest speech

j = 0
ln = 0
For i = 1 To NumSpeech
    If InStr(1, d$, Speech(i).ClientText) And Len(Speech(i).ClientText) > ln And Speech(i).ClientText <> " " Then
        If Speech(i).NumAnswers = 0 And ln > 0 Then
        Else
            j = i
            ln = Len(Speech(i).ClientText)
        End If
        
    End If
    If Speech(i).ClientText = " " Then G = i
Next i

If j <> 0 Then
    'found
    k = Speech(j).NumAnswers
    If k = 0 Then
        If G > 0 Then
            k = Speech(G).NumAnswers
            If k = 0 Then Exit Sub
            
            'pick a random one
            Randomize Timer
            m = Int(Rnd * k) + 1
                
            c$ = Speech(G).Answers(m)
            c$ = Replace(c$, "%a", General.AutoAdminName)
            c$ = Replace(c$, "%n", nm$)
            c$ = Replace(c$, "%u", us$)
            c$ = Replace(c$, "%s", Sy$)
            
            'exec it
            
            
            NewScriptData.TimeStarted = Timer
            NewScriptData.UserName = "<SERVER>"
            NewScriptData.StartedName = "Admin Speech " + c$
            asd = ExecuteScriptParams(c$, NewScriptData, True)
            If asd = False Then SendRCONCommand c$
        End If
        Exit Sub
    End If
    
    'pick a random one
    Randomize Timer
    m = Int(Rnd * k) + 1
        
    c$ = Speech(j).Answers(m)
    c$ = Replace(c$, "%a", General.AutoAdminName)
    c$ = Replace(c$, "%n", nm$)
    c$ = Replace(c$, "%u", us$)
    c$ = Replace(c$, "%s", Sy$)
    
    'exec it
    
    NewScriptData.TimeStarted = Timer
    NewScriptData.UserName = "<SERVER>"
    NewScriptData.StartedName = "Admin Speech " + c$
    asd = ExecuteScriptParams(c$, NewScriptData, True)
    If asd = False Then SendRCONCommand c$
    
Else
    'not found, so add it.
    If CheckBit2(General.Flags, 1) Then
        NumSpeech = NumSpeech + 1
        
        j = NumSpeech
        
        ReDim Preserve Speech(0 To j)
        
        Speech(j).NumAnswers = 0
        ReDim Speech(j).Answers(0 To 1)
        Speech(j).ClientText = d$
        SaveCommands
    End If
    
    'Now give a generic answer
    
    If G > 0 Then
        k = Speech(G).NumAnswers
        If k = 0 Then Exit Sub
        
        'pick a random one
        Randomize Timer
        m = Int(Rnd * k) + 1
            
        c$ = Speech(G).Answers(m)
        c$ = Replace(c$, "%a", General.AutoAdminName)
        c$ = Replace(c$, "%n", nm$)
        c$ = Replace(c$, "%u", us$)
        c$ = Replace(c$, "%s", Sy$)
        
        'exec it
        
        NewScriptData.TimeStarted = Timer
        NewScriptData.UserName = "<SERVER>"
        NewScriptData.StartedName = "Admin Speech " + c$
        asd = ExecuteScriptParams(c$, NewScriptData, True)
        If asd = False Then SendRCONCommand c$
    End If
        
End If

End Sub


Function SpecialAdminSpeech(Sy$, nm$, us$) As String
If DebugMode Then LastCalled = "SpecialAdminSpeech"

'First check if the saying is already in the database
d$ = UCase(Sy$)

Dim NewScriptData As typScriptData

j = 0
ln = 0
For i = 1 To NumSpeech
    If InStr(1, d$, Speech(i).ClientText) And Len(Speech(i).ClientText) > ln And Speech(i).ClientText <> " " Then
        If Speech(i).NumAnswers = 0 And ln > 0 Then
        Else
            j = i
            ln = Len(Speech(i).ClientText)
        End If
        
    End If
    If Speech(i).ClientText = " " Then G = i
Next i

If j <> 0 Then
    'found
    k = Speech(j).NumAnswers
    If k = 0 Then j = G
    
    k = Speech(j).NumAnswers
    If k = 0 Then Exit Function
            
    'pick a random one
    Randomize Timer
    m = Int(Rnd * k) + 1
        
    c$ = Speech(j).Answers(m)
    c$ = Replace(c$, "%a", "")
    c$ = Replace(c$, "%n", nm$)
    c$ = Replace(c$, "%u", us$)
    c$ = Replace(c$, "%s", Sy$)
    
    'exec it
    
        
    
Else
    'Now give a generic answer
    
    If G > 0 Then
        k = Speech(G).NumAnswers
        If k = 0 Then Exit Function
        
        'pick a random one
        Randomize Timer
        m = Int(Rnd * k) + 1
            
        c$ = Speech(G).Answers(m)
        c$ = Replace(c$, "%a", "")
        c$ = Replace(c$, "%n", nm$)
        c$ = Replace(c$, "%u", us$)
        c$ = Replace(c$, "%s", Sy$)
        
        'exec it
    End If
End If

SpecialAdminSpeech = c$

End Function


Function CheckPlayerName(Num, Optional mde As Boolean) As Boolean
If DebugMode Then LastCalled = "CheckPlayerName"

'Checks if the player is in one of the listed clans.
'If he is, checks to see if his uniqueid is in the list.
'If it isnt, kicks him and tells him how to add himself

'Check for swearing
CheckName Num

n$ = Players(Num).Name
Un$ = Players(Num).UniqueID

'ensure player isnt breaking realname laws
j = 0
For i = 1 To NumRealPlayers
    If InStr(1, RealPlayers(i).UniqueID, Chr(34) + Un$ + Chr(34) + ";") Then j = i: Exit For
Next i

If j > 0 Then
    'See if Force Player to Use This Name is enabled
    If CheckBit2(RealPlayers(j).Flags, 1) Then
        If RealPlayers(j).RealName <> n$ Then ChangePlayerName Players(Num).UserID, RealPlayers(j).RealName
    ElseIf CheckBit2(RealPlayers(j).Flags, 2) And mde = True Then
        If RealPlayers(j).RealName <> n$ Then SendRCONCommand "say Player " + n$ + " is really " + RealPlayers(j).RealName + " !", 5
    End If
End If
        
'lets see if this player is using someone elses name
For i = 1 To NumRealPlayers
    If LCase(RealPlayers(i).RealName) = LCase(n$) Then k = i: Exit For
Next i

If k > 0 And k <> j Then 'just as long as its not actually THEM!
    'is this name protected?
    If CheckBit2(RealPlayers(k).Flags, 0) Then 'yes
        ChangePlayerName Players(Num).UserID, "Player"
    End If
End If


'now do clan checking

For i = 1 To NumClans
    
    'check if his clan is restricted
    If InStr(1, n$, Clans(i).Clan) Then
        'yes it is...
        'See if his unique is in the list...
        
        nokick = 0
        For j = 1 To Clans(i).NumMembers
            If InStr(1, Clans(i).Members(j).UIN, Un$) Then 'alls well, hes there, so update some info
                Clans(i).Members(j).LastIP = Players(Num).IP
                Clans(i).Members(j).Name = n$
                Players(Num).RemoveMe = False
                nokick = 1
            Else
                'this UIN is NOT in the list
                'GoTo kickhim
            End If
        Next j
        
        If nokick = 0 Then GoTo kickhim
        If Clans(i).NumMembers = 0 Then GoTo kickhim
    End If
Next i

Exit Function
kickhim:

'see if we are supposed to kick the person
If CheckBit2(Clans(i).Flags, 0) = True Then

    'kicks the player then tells him what to do
    
    Players(Num).RemoveMe = True 'so he wont be removed
    SendRCONCommand "kick # " + Ts(Players(Num).UserID)
    
    m$ = Chr(10) + Chr(10) + "You are not currently registered as a member of clan " + Clans(i).Clan + "." + Chr(10) + "If you are legitematly a member, follow these instructions:"
    m$ = m$ + Chr(10) + Chr(10) + "Type at the console (in this order):" + Chr(10) + "rcon_port " + Ts(Form1.RconUDP.LocalPort)
    m$ = m$ + Chr(10) + "rcon_address " + Server.LocalIP
    m$ = m$ + Chr(10) + "rcon_password (your clans JOIN password)"
    m$ = m$ + Chr(10) + "rcon addme"
    m$ = m$ + Chr(10) + Chr(10) + "Once this is done, you will recieve a message to your console giving you more insctructions." + Chr(10) + Chr(10) + Chr(10)
    
    'send it twice to be sure
    SendToUserIP Players(Num).IP, Players(Num).Port, m$
    SendToUserIP Players(Num).IP, Players(Num).Port, m$

ElseIf CheckBit2(Clans(i).Flags, 5) = True And DLLEnabled = True Then 'Change player's name to not include the clan tag
    
    c$ = Replace(Players(Num).Name, Clans(i).Clan, "")
    c$ = Trim(c$)
    If c$ = "" Then c$ = "Player"
    ChangePlayerName Players(Num).UserID, c$

End If


End Function

Sub SendToWatchers(n$, nm$, Sy$, mde, Team, Optional MapChange As Integer)
If DebugMode Then LastCalled = "SendToWatchers"

'sends this message to all rcon watchers

If mde = 1 Then G$ = "(TEAM) "

cc = Team

nn$ = "<" + n$ + "> " + G$ + nm$ + ": "
nn2$ = G$ + nm$ + ": "
If nm$ = "" Then nn$ = "<ADMIN> ": nn2$ = "<ADMIN> ": cc = 6
If n$ = "SERVER" Then nn$ = "<SERVER> ": nn2$ = "<SERVER> ": cc = 7
If n$ = "MESSAGE" Then nn$ = "<MESSAGE> ": nn2$ = "<MESSAGE> ": cc = 8
If n$ = "ADMIN" Then nn2$ = "<ADMIN> " + nm$ + ": ": cc = 6
If n$ = "TELL" Then nn2$ = "<TELL " + nm$ + "> ": cc = 10
If n$ = "OTHER" Then nn2$ = nm$: nn$ = "": cc = 6

For i = 1 To RconWatchersIP.count
    
    B$ = RconWatchersIP(i)
    c = RconWatchersPort(i)
    Vars.UserIsTCP = False
    SendToUserIP B$, c, ">>> " + nn$ + Sy$
    
Next i

For i = 2 To 20
    LastLines(i - 1).Line = LastLines(i).Line
    LastLines(i - 1).Team = LastLines(i).Team
    LastLines(i - 1).Name = LastLines(i).Name
    LastLines(i - 1).TimeSent = LastLines(i).TimeSent
Next i

LastLines(20).Line = Sy$
LastLines(20).Name = nn2$
LastLines(20).Team = cc
LastLines(20).TimeSent = Time$

If n$ = "SERVER" Then
    AddToHTMLFile Sy$, 15, 0, 0, "", ""
End If
If n$ = "MESSAGE" Then
    AddToHTMLFile Sy$, 15, 0, 0, "", ""
End If

'also send this to any connected users with the right permissions
nmorig$ = nm$

If nm$ <> "" Then nm$ = nm$ + ":"

For i = 1 To NumConnectUsers
    
    c = ConnectUsers(i).UserNum
    d = ConnectUsers(i).LogLevel
    
    If CheckBit(c, 10) Then
    
        cc = Team
        If CheckBit2(d, 0) And mde = 0 And n$ <> "TELL" And n$ <> "OTHER" And n$ <> "ADMIN" And n$ <> "MESSAGE" And n$ <> "SERVER" Then SendChatPacket G$ + nm$, Sy$, ConnectUsers(i).Index, i, cc, Time$
        If CheckBit2(d, 3) And mde = 0 And n$ = "ADMIN" Then SendChatPacket "<ADMIN> " + nm$, Sy$, ConnectUsers(i).Index, i, 6, Time$
        If CheckBit2(d, 1) And n$ = "TELL" Then SendChatPacket "<TELL " + nmorig$ + ">", Sy$, ConnectUsers(i).Index, i, 10, Time$
        If CheckBit2(d, 3) And mde = 0 And n$ = "SERVER" Then SendChatPacket "<SERVER> " + nm$, Sy$, ConnectUsers(i).Index, i, 7, Time$
        If CheckBit2(d, 3) And mde = 0 And n$ = "OTHER" Then SendChatPacket nn2$, Sy$, ConnectUsers(i).Index, i, 6, Time$
        If CheckBit2(d, 3) And mde = 0 And n$ = "MESSAGE" Then SendChatPacket "<MESSAGE> " + nm$, Sy$, ConnectUsers(i).Index, i, 8, Time$
        If CheckBit2(d, 1) And mde = 1 Then SendChatPacket G$ + nm$, Sy$, ConnectUsers(i).Index, i, cc, Time$
        
        If MapChange = 1 Then
            SendPacket "NM", "", ConnectUsers(i).Index
        End If
    
    End If
Next i



End Sub

Sub SendChatPacket(nme$, Sy$, Index, Ver, Col, TimeSent As String)


If Val(Replace(ConnectUsers(Ver).Version, ".", "")) < 1118 Then
    'Older version compatibility!
    SendPacket "TY", nme$ + " " + Sy$, CInt(Index)
Else
    a$ = a$ + Chr(251)
    a$ = a$ + nme$ + " " + Sy$ + Chr(250)
    a$ = a$ + nme$ + Chr(250)
    a$ = a$ + Ts(Col) + Chr(250)
    a$ = a$ + TimeSent + Chr(250)
    a$ = a$ + Chr(251)
    
    'all set, send it
    SendPacket "TY", a$, CInt(Index)

End If


End Sub

Sub StartMapVote(scriptdata As typScriptData)
If DebugMode Then LastCalled = "StartMapVote"

Vars.AlreadyAutoVoted = True

GetMapList

If MapVoteTimer > 0 Then
    SendToUser "The mapvote is already running, you can't start it now!", scriptdata
    Exit Sub
End If


For i = 1 To 10

    'Pick a random map
    Randomize
    rn = Int(Rnd * AvailMaps.count) + 1
    bbc$ = AvailMaps(rn)

    ex = 1
    If LCase(bbc$) = LCase(Vars.Map) Then ex = 0
    For j = 1 To LastMaps.count
        If LCase(bbc$) = LCase(LastMaps(j)) Then ex = 0
    Next j
    If ex = 1 Then Exit For
Next i

LastTalk = 1
SendRCONCommand "say MAP VOTE TIME!", , 1
SendRCONCommand "say Type THE NAME of the map to vote for that map!", , 1
SendRCONCommand "say Do NOT type VOTE first! Suggested map: " + bbc$

If DLLEnabled Then
    SendMessage "MAP VOTE TIME!" + Chr(10) + "Type the NAME of the MAP to vote for that map!" + Chr(10) + "Suggested Map: " + bbc$, 1, 1, 255, 1, 255, 1, 1, 2, 15, 0.1, 0.02, 5, -1, 0.2
End If
 
ExecFunctionScript "spec_mapvotestart", 1, bbc$

KickVoteTimer = 0
ChooseVoteTimer = 0
MapVoteTimer = 130
NumVotes = 0

End Sub

Sub StartKickVote(n$)
If DebugMode Then LastCalled = "StartKickVote"

Num = FindPlayer(Ts(KickVoteUser))
If Num = 0 Then Exit Sub
n$ = Players(Num).Name

LastTalk = 1

Dim OptionList(1 To 3) As String
OptionList(1) = "Yes"
OptionList(2) = "No"
OptionList(3) = "Abstain (Don't Vote)"

If CheckBit2(General.Flags, 3) Then
    GenerateMenu "KICK VOTE! Voting to kick: " + Chr(10) + n$, 30, OptionList, 3
Else
    SendRCONCommand "say KICK VOTE! Voting to kick: " + n$, , 1
    SendRCONCommand "say Type YES or NO to vote! 60 seconds of time!"
End If

ExecFunctionScript "spec_kickvotestart", 2, Ts(KickVoteUser), KickVoteStarterName

KickVoteTimer = 60
NumVotes = 0

End Sub

Sub GetMapList()
If DebugMode Then LastCalled = "GetMapList"

'gets the list of maps and puts em in an array

Do Until AvailMaps.count = 0
    AvailMaps.Remove (1)
Loop

B$ = Server.BothPath + "\maps\*.bsp"
a$ = Dir(B$)

Do Until a$ = ""
    
    e = InStrRev(a$, ".")
    c$ = ""
    If e > 0 Then c$ = Left(a$, e - 1)
    If UCase(Left(a$, 3)) <> "XXX" And c$ <> "" Then AvailMaps.Add UCase(c$)
    
    a$ = Dir
Loop

If LCase(Server.GamePath) <> "valve" Then
    B$ = Server.HLPath + "\valve\maps\*.bsp"
    a$ = Dir(B$)
    
    Do Until a$ = ""
        
        e = InStrRev(a$, ".")
        c$ = ""
        If e > 0 Then c$ = Left(a$, e - 1)
        If UCase(Left(a$, 3)) <> "XXX" And c$ <> "" Then AvailMaps.Add UCase(c$)
        
        a$ = Dir
    Loop
End If

'these maps are in pak file
AvailMaps.Add "ROCK2"
AvailMaps.Add "2FORT"
AvailMaps.Add "HUNTED"
AvailMaps.Add "PUSH"
AvailMaps.Add "WELL"
AvailMaps.Add "CZ2"



End Sub

Sub TrackVote(n$, Sy$)
If DebugMode Then LastCalled = "TrackVote"

'checks if it is a vote

If MapVoteTimer <= 0 Then Exit Sub

Sy$ = UCase(Sy$)
'r = InStr(1, sy$, "VOTE")
'If r = 0 Then Exit Sub

If LCase(LeftR(Sy$, 4)) = "vote" Then 'remove the word vote
    
    Sy$ = Trim(RightR(Sy$, Len(Sy$) - 4))
    votemode = 1
    
End If

'e = InStr(1, sy$, " ")
'If e = 0 Then Exit Sub
f = Len(Sy$)

'If e = f Then Exit Sub
'vt$ = Mid(sy$, e + 1, f - e)

vt$ = Sy$
vt$ = Trim(vt$)
If vt$ = "" Then Exit Sub
If InStr(1, vt$, " ") Then Exit Sub

u = 0


'If vt$ = "CZ" Then vt$ = "CZ2"
'If vt$ = "ROCK" Then vt$ = "ROCK2"
'If vt$ = "BOOTCAMP" Then vt$ = "BOOT_CAMP"
'If vt$ = "BOOT_CAMP_TFC" Then vt$ = "BOOT_CAMP"
'If vt$ = "2MOREFORTS" Then vt$ = "2MORFORT"
'If vt$ = "BORDER" Then vt$ = "TFCBORDER2A"
'If vt$ = "TFCBORDER" Then vt$ = "TFCBORDER2A"
'If vt$ = "INVECT" Then vt$ = "INVECT3"
'If vt$ = "STRIKE" Then vt$ = "STRIKE2"


For i = 1 To AvailMaps.count
    If vt$ = AvailMaps(i) Then u = 1: Exit For
Next i

If u = 0 Then
    If votemode = 1 Then
        SendRCONCommand "say " + General.AutoAdminName + " The map " + vt$ + " is not on this server."
    
    End If

    Exit Sub
End If


If General.LastMapsDisabled = False Then
    For i = 1 To LastMaps.count
        If vt$ = UCase(LastMaps(i)) Then
            SendRCONCommand "say " + General.AutoAdminName + " You cannot vote for the map " + vt$ + "."
            Exit Sub
        End If
    Next i
End If

'check for player
j = 0
For i = 1 To NumVotes
    If Votes(i).UserID = n$ Then j = i: Exit For
Next i

If j = 0 Then
    NumVotes = NumVotes + 1
    j = NumVotes
End If

Votes(j).UserID = n$
Votes(j).MapChoice = vt$
Votes(j).Double = False

Num = FindPlayer(n$)
If Num > 0 Then
    c2 = FindClan(Num)
    If c2 > 0 Then
        If CheckBit2(Clans(c2).Flags, 3) Then
            Votes(j).Double = True
        End If
    End If
    If PointData.DoubleMapVotesAt > 0 And GetPoints(Num) > PointData.DoubleMapVotesAt Then Votes(j).Double = True

End If
        

End Sub

Sub TrackKickVote(n$, Sy$)
If DebugMode Then LastCalled = "TrackKickVote"

'checks if it is a vote

If KickVoteTimer <= 0 Or MapVoteTimer > 0 Then Exit Sub

Sy$ = UCase(Sy$)

f = Len(Sy$)

'make sure the kickee isnt voting
If Val(n$) = KickVoteUser Then Exit Sub


vt$ = Sy$
vt$ = Trim(vt$)
If vt$ = "" Then Exit Sub
If InStr(1, vt$, " ") Then Exit Sub

u = 0

'see if its either YES or NO

If vt$ = "YES" Then u = 1
If vt$ = "NO" Then u = 1

If u = 0 Then Exit Sub

'check for player
j = 0
For i = 1 To NumVotes
    If Votes(i).UserID = n$ Then j = i: Exit For
Next i

If j = 0 Then
    NumVotes = NumVotes + 1
    j = NumVotes
End If

Votes(j).UserID = n$
Votes(j).MapChoice = vt$
Votes(j).Double = False

Num = FindPlayer(n$)
If Num > 0 Then
    c2 = FindClan(Num)
    If c2 > 0 Then
        If CheckBit2(Clans(c2).Flags, 4) Then
            Votes(j).Double = True
        End If
    End If
    If PointData.DoubleKickVotesAt > 0 And GetPoints(Num) > PointData.DoubleKickVotesAt Then Votes(j).Double = True
    
End If


End Sub

Sub TrackChooseVote(n$, Sy$)
If DebugMode Then LastCalled = "TrackChooseVote"

'checks if it is a vote

If ChooseVoteTimer <= 0 Or MapVoteTimer > 0 Or KickVoteTimer > 0 Then Exit Sub

Sy$ = UCase(Sy$)
f = Len(Sy$)


vt$ = Sy$
vt$ = Trim(vt$)
If vt$ = "" Then Exit Sub

u = 0

'see if its any of the choices
For i = 1 To UBound(ChooseVote)
    If LeftR(vt$, Len(ChooseVote(i))) = UCase(ChooseVote(i)) Then u = 1
Next i

If u = 0 Then Exit Sub

'check for player
j = 0
For i = 1 To NumVotes
    If Votes(i).UserID = n$ Then j = i: Exit For
Next i

If j = 0 Then
    NumVotes = NumVotes + 1
    j = NumVotes
End If

Votes(j).UserID = n$
Votes(j).MapChoice = vt$
Votes(j).Double = False

End Sub

Function TotalVotes(mde) As String
If DebugMode Then LastCalled = "TotalVotes"

'totals the votes
'mde = 0 - returns the status
'mde = 1 - sets the next map

Dim MapNames(1 To 40) As String
Dim MapVotes(1 To 40) As Integer
Dim NumMaps As Integer

For i = 1 To NumVotes

    vt$ = Votes(i).MapChoice
      
    If FindPlayer(Ts(Votes(i).UserID)) Then
    
        k = 0
        For j = 1 To NumMaps
            If MapNames(j) = vt$ Then k = j: Exit For
        Next j
        
        If k = 0 Then
            NumMaps = NumMaps + 1
            k = NumMaps
            MapVotes(k) = 0
        End If
        
        MapNames(k) = vt$
        MapVotes(k) = MapVotes(k) + 1
        
        If Votes(i).Double = True Then 'add another point
            MapVotes(k) = MapVotes(k) + 1
        End If
    End If
Next i

If mde = 0 Then
    'assemble rcon send package
    
    a$ = "Votes: "
    
    For i = 1 To NumMaps
        a$ = a$ + MapNames(i) + "=" + Ts(MapVotes(i)) + ", "
    Next i

    If Len(a$) > 2 Then a$ = Left(a$, Len(a$) - 2)
    If a$ = "Votes" Then a$ = "No Votes Yet"
End If

If mde = 1 Then
    'total the votes and set the next map
    
    a$ = "say No map determined"
    
    topvote = 0
    TopChoice = 0
    
    'check to see if one map got all the votes
    If NumMaps = 1 Then
        TopChoice = 1
        topvote = MapVotes(1)
    Else
        For i = 1 To NumMaps
            If MapNames(i) <> Vars.Map Then
                If MapVotes(i) > topvote Then topvote = MapVotes(i): TopChoice = i
            End If
        Next i
    End If
    
    If TopChoice > 0 Then
        B$ = MapNames(TopChoice)
        
        SendRCONCommand "say Map Vote Complete! The NEXT map is " + B$ + "!"
        
        If DLLEnabled = True Then
           SendMessage "Map Vote Complete!" + Chr(10) + "The NEXT map is " + B$ + "!" + Chr(10) + "It won with " + Ts(topvote) + " votes!", 1, 1, 255, 1, 255, 1, 1, 2, 15, 0.1, 0.02, 5, -1, 0.2
        End If
        
        ExecFunctionScript "spec_mapvotecomplete", 2, B$, Ts(topvote)
        Vars.VotedMap = B$
        
        If General.MapVoteMode = "3" Then 'use SA_NEXTMAP
            SendRCONCommand "sa_nextmap " + B$
        End If
        
    Else
        SendRCONCommand "say No map was determined."
    End If
End If

TotalVotes = a$

End Function

Sub ExecFunctionScript(ScriptName As String, Num As Integer, Optional Param1 As String, Optional Param2 As String, Optional Param3 As String, Optional Param4 As String, Optional Param5 As String, Optional Param6 As String)

Dim UserParms() As typParams
ReDim UserParms(0 To Num)


If Num >= 1 Then UserParms(1).Value = Param1
If Num >= 2 Then UserParms(2).Value = Param2
If Num >= 3 Then UserParms(3).Value = Param3
If Num >= 4 Then UserParms(4).Value = Param4
If Num >= 5 Then UserParms(5).Value = Param5
If Num >= 6 Then UserParms(6).Value = Param6

'see if it needs to be broadcast
Dim NewScriptData As typScriptData
NewScriptData.TimeStarted = Timer
NewScriptData.UserName = "<SERVER>"
NewScriptData.StartedName = "ExecFunctionScript: " + ScriptName

ExecuteScript ScriptName, UserParms, Num, "", NewScriptData

End Sub

Sub ExecFunctionScript2(ScriptName As String, Num As Integer, scriptdata As typScriptData, Optional Param1 As String, Optional Param2 As String, Optional Param3 As String, Optional Param4 As String, Optional Param5 As String, Optional Param6 As String)
If scriptdata.ExitNow = True Then Exit Sub
Dim UserParms() As typParams
ReDim UserParms(0 To Num)

If Num >= 1 Then UserParms(1).Value = Param1
If Num >= 2 Then UserParms(2).Value = Param2
If Num >= 3 Then UserParms(3).Value = Param3
If Num >= 4 Then UserParms(4).Value = Param4
If Num >= 5 Then UserParms(5).Value = Param5
If Num >= 6 Then UserParms(6).Value = Param6

'see if it needs to be broadcast
ExecuteScript ScriptName, UserParms, Num, "", scriptdata

End Sub

Sub TotalKickVotes()
If DebugMode Then LastCalled = "TotalKickVotes"

'Err.Source
'totals the votes
'and kicks the player if needed

Dim YesVotes As Integer
Dim NoVotes As Integer
Dim TotVotes As Integer

For i = 1 To NumVotes

    If FindPlayer(Ts(Votes(i).UserID)) Then

        vt$ = Votes(i).MapChoice
        
        If vt$ = "YES" And Votes(i).Double = False Then YesVotes = YesVotes + 1
        If vt$ = "NO" And Votes(i).Double = False Then NoVotes = NoVotes + 1
        If vt$ = "YES" And Votes(i).Double = True Then YesVotes = YesVotes + 2
        If vt$ = "NO" And Votes(i).Double = True Then NoVotes = NoVotes + 2

    End If
Next i
   
TotVotes = YesVotes + NoVotes
If TotVotes > 0 Then YesPerc = Int((YesVotes / TotVotes) * 100)
 
'see if player is there, and check all cases

Num = FindPlayer(Ts(KickVoteUser))
If Num > 0 Then
    n$ = Players(Num).Name

    If TotVotes = 0 Then
        LastTalk = 1
        SendRCONCommand "say " + General.AutoAdminName + " There are no votes.", , 1
        SendRCONCommand "say " + General.AutoAdminName + " KickVote Complete."
    Else
        If General.VotePercent = 0 Then General.VotePercent = 75
        If YesPerc < General.VotePercent Then
            
            LastTalk = 1
            SendRCONCommand "say " + General.AutoAdminName + " Player " + n$ + " not kicked.", , 1
            SendRCONCommand "say " + General.AutoAdminName + " Vote was " + Ts(YesPerc) + "% yes, need " + Ts(General.VotePercent) + "% or more"
            
            AddToVoteFile Num, YesPerc, 2
            
            ExecFunctionScript "spec_kickvotecomplete", 5, Ts(Players(Num).UserID), "0", Ts(YesPerc), Ts(General.VotePercent), KickVoteStarterName

            
        Else
            'set last var
            For i = 19 To 1 Step -1
                LastKickVotes(i + 1) = LastKickVotes(i)
            Next i
            LastKickVotes(1) = Players(Num).UniqueID
            
            If PointData.KickedCosts > 0 Then SetPoints Num, GetPoints(Num) - PointData.KickedCosts
            
            'COUNT
            c1 = 0
            For i = 1 To 20
                If LastKickVotes(i) = Players(Num).UniqueID Then c1 = c1 + 1
            Next i
            
            If General.MaxKicks = 0 Then General.MaxKicks = 2
            If General.BanTime = 0 Then General.BanTime = 15
            If c1 >= General.MaxKicks And General.MaxKicks > 0 And General.BanTime > 0 Then
                
                SendRCONCommand "banid " + Ts(General.BanTime) + " " + Players(Num).UniqueID + " kick"
                                
                AddToLogFile "KICKVOTE: Player " + n$ + " (UNIQUE: " + Players(Num).UniqueID + ") kicked, with " + Ts(YesPerc) + "% yes vote, AND ALSO BANNED for " + Ts(General.BanTime) + " minutes. Kick vote was started by " + KickVoteStarterName + " (UIN " + KickVoteStarterUIN + ")"
                LastTalk = 1
                SendRCONCommand "say " + General.AutoAdminName + " Player " + n$ + " kicked and banned for " + Ts(General.BanTime) + " minutes!", , 1
                SendRCONCommand "say " + General.AutoAdminName + " Vote was " + Ts(YesPerc) + "% yes"
                
                If Players(Num).IP <> "" And Players(Num).Port > 0 Then
                    Vars.UserIsTCP = False
                    SendToUserIP Players(Num).IP, Players(Num).Port, "You have been kicked and BANNED for " + Ts(General.BanTime) + " minutes due to a kick vote that was " + Ts(YesPerc) + "% yes vote"
                    SendToUserIP Players(Num).IP, Players(Num).Port, "Most likely you were abusing the server. Feel free to join again in " + Ts(General.BanTime) + " minutes."
                End If
                                
                AddToVoteFile Num, YesPerc, 1
                
                ExecFunctionScript "spec_kickvotecomplete", 5, Ts(Players(Num).UserID), "2", Ts(YesPerc), Ts(General.VotePercent), KickVoteStarterName

                                
            Else
                SendRCONCommand "kick # " + Ts(KickVoteUser)
                AddToLogFile "KICKVOTE: Player " + n$ + " (UNIQUE: " + Players(Num).UniqueID + ") kicked, with " + Ts(YesPerc) + "% yes vote. Kick vote was started by " + KickVoteStarterName + " (UIN " + KickVoteStarterUIN + ")"
                LastTalk = 1
                SendRCONCommand "say " + General.AutoAdminName + " Player " + n$ + " kicked!", , 1
                SendRCONCommand "say " + General.AutoAdminName + " Vote was " + Ts(YesPerc) + "% yes"
                
                If Players(Num).IP <> "" And Players(Num).Port > 0 Then
                    Vars.UserIsTCP = False
                    SendToUserIP Players(Num).IP, Players(Num).Port, "You have been kicked due to a kick vote that was " + Ts(YesPerc) + "% yes vote"
                    SendToUserIP Players(Num).IP, Players(Num).Port, "Most likely you were abusing the server. Feel free to join again if you think you can control yourself."
                End If
                        
                AddToVoteFile Num, YesPerc, 0
                
                ExecFunctionScript "spec_kickvotecomplete", 5, Ts(Players(Num).UserID), "1", Ts(YesPerc), Ts(General.VotePercent), KickVoteStarterName

            
            End If
        End If
    End If
Else
    LastTalk = 1
    SendRCONCommand "say " + General.AutoAdminName + " The player could not be found.", , 1
    SendRCONCommand "say " + General.AutoAdminName + " KickVote Complete."
End If
   

End Sub

Sub TotalChooseVotes()
If DebugMode Then LastCalled = "TotalChooseVotes"

'totals the votes

Dim ChVotes() As Integer
Dim TotVotes As Integer

ReDim ChVotes(1 To UBound(ChooseVote))

For i = 1 To NumVotes

    If FindPlayer(Ts(Votes(i).UserID)) Then

        vt$ = Votes(i).MapChoice
        
        For j = 1 To UBound(ChooseVote)
            If vt$ = UCase(ChooseVote(j)) Then ChVotes(j) = ChVotes(j) + 1
        Next j
    End If
    
Next i

TotVotes = 0
topvote = 0
For j = 1 To UBound(ChooseVote)
    TotVotes = TotVotes + ChVotes(j)
    If ChVotes(j) > topvote Then topvote = ChVotes(j): topnum = j
Next j




If TotVotes > 0 Then
    For j = 1 To UBound(ChooseVote)
        ChVotes(j) = Round((ChVotes(j) / TotVotes) * 100)
    Next j
End If
'see if player is there, and check all cases

If TotVotes = 0 Then
    LastTalk = 1
    SendRCONCommand "say " + General.AutoAdminName + " There are no votes.", , 1
    SendRCONCommand "say " + General.AutoAdminName + " General Vote Complete."
Else

    n$ = ""
    
    For j = 1 To UBound(ChooseVote)
        If ChooseVote(j) <> "" Then n$ = n$ + UCase(ChooseVote(j)) + "=" + Ts(ChVotes(j)) + "% "
    Next j
    
    LastTalk = 1
    SendRCONCommand "say The Results: Of " + Ts(TotVotes) + " votes,", , 1
    SendRCONCommand "say " + n$
    
    If topnum > 0 And topvote > 0 Then ExecFunctionScript "spec_choosevotecomplete", 3, ChooseVote(topnum), Ts(topvote), Ts(TotVotes)

   
End If

End Sub

Sub BanPlayerReason(UniqueID As String, Name As String, Reason As String, Optional Num As Integer)
If DebugMode Then LastCalled = "BanPlayerReason"

'Bans a player, and adds them to the BANLIST.CFG file.

If Num = 0 Then 'no num

    SendRCONCommand "banid 0 " + UniqueID + " kick"

    a$ = Server.BothPath + "\"

    If Dir(a$, vbDirectory) = "" Then MkDir a$
    a$ = a$ + "\banlist.cfg"
    
    h = FreeFile
Close h
    Open a$ For Append As h
        Print #h, "//Date/Time/Map: " + Date$ + " " + Time$ + " - " + Vars.Map
        Print #h, "//Player Name: " + Name
        Print #h, "//UniqueID: " + UniqueID
        Print #h, "//Reason: " + Reason
        Print #h, "banid 0 " + UniqueID
        Print #h, ""
    Close h

Else

    SendRCONCommand "banid 0 " + Players(Num).UniqueID + " kick"

    a$ = Server.BothPath + "\"

    If Dir(a$, vbDirectory) = "" Then MkDir a$
    a$ = a$ + "\banlist.cfg"
    
    h = FreeFile
Close h
    Open a$ For Append As h
        Print #h, "//Date/Time/Map: " + Date$ + " " + Time$ + " - " + Vars.Map
        Print #h, "//Player Name: " + Players(Num).Name
        Print #h, "//UniqueID: " + Players(Num).UniqueID
        Print #h, "//Reason: " + Reason
        Print #h, "//IP: " + Players(Num).IP
        If Players(Num).RealName <> "" Then Print #h, "//Real Name: " + Players(Num).RealName
        If Players(Num).EntryName <> "" Then Print #h, "//Entry Name: " + Players(Num).EntryName
        Print #h, "banid 0 " + Players(Num).UniqueID
        Print #h, ""
    Close h
End If


End Sub

Sub FinishVote(B$)
If DebugMode Then LastCalled = "FinishVote"

On Error Resume Next


'make a temp file

'a$ = Server.BothPath + "\temp.rc"
'If Dir(a$) <> "" Then Kill a$
'
'h = freefile
Close h
'Open a$ For Append As #H
'    Print #H, b$
'Close #H

'now attempt to execute it
SendRCONCommand "exec temp.rc"

'good, now we need to check server.cfg for specific lines and remove em as needed
h = FreeFile
Close h


If Dir(Server.BothPath + "\server.tmp") <> "" Then Kill Server.BothPath + "\server.tmp"
If Dir(Server.BothPath + "\server.cfg") = "" Then Exit Sub

Open Server.BothPath + "\server.cfg" For Input As #h
    j = FreeFile
    Open Server.BothPath + "\server.tmp" For Append As #j
        Do While Not (EOF(h))
            Line Input #h, n$
                
            If Not InStr(1, LCase(a$), "nextmap") Then
                Print #j, n$
            End If
        Loop
        
        Print #j, ""
        Print #j, "//Nextmap Settings - For Map Vote"
        Print #j, "nextmap"
        Print #j, "//Nextmap - Clear the Nextmap Alias"
        Print #j, "alias nextmap echo"
        
    Close #j
Close #h

'copy the file
Kill Server.BothPath + "\server.cfg"
Name Server.BothPath + "\server.tmp" As Server.BothPath + "\server.cfg"

End Sub

Sub RemovePlayer(Num)
If DebugMode Then LastCalled = "RemovePlayer"

'trashes a player from the list

'FIRST add points based on how long they played


ap = Minute(Now - Players(Num).TimeJoined) + (Hour(Now - Players(Num).TimeJoined) * 60)
ap = Int(ap * PointData.JoiningAdds)
If PointData.JoiningAdds > 0 Then SetPoints Num, GetPoints(Num) + ap


If Num < NumPlayers Then
    For i = Num To NumPlayers - 1
        Players(i).Class = Players(i + 1).Class
        Players(i).IP = Players(i + 1).IP
        Players(i).Name = Players(i + 1).Name
        Players(i).Team = Players(i + 1).Team
        Players(i).UniqueID = Players(i + 1).UniqueID
        Players(i).UserID = Players(i + 1).UserID
        Players(i).Port = Players(i + 1).Port
        Players(i).RemoveMe = Players(i + 1).RemoveMe
        Players(i).ConnectOnly = Players(i + 1).ConnectOnly
        Players(i).ThereFlag = Players(i + 1).ThereFlag
        Players(i).RealName = Players(i + 1).RealName
        Players(i).NumKickVotes = Players(i + 1).NumKickVotes
        Players(i).MessInMin = Players(i + 1).MessInMin
        'Players(I).CountStart = Players(I + 1).CountStart
        Players(i).EntryName = Players(i + 1).EntryName
        Players(i).BroadcastType = Players(i + 1).BroadcastType
        Players(i).Pos.X = Players(i + 1).Pos.X
        Players(i).Pos.Y = Players(i + 1).Pos.Y
        Players(i).Pos.Z = Players(i + 1).Pos.Z
        Players(i).ShutUp = Players(i + 1).ShutUp
        Players(i).TimeJoined = Players(i + 1).TimeJoined
        Players(i).Warn = Players(i + 1).Warn
        Players(i).TempRealMode = Players(i + 1).TempRealMode
        Players(i).Points = Players(i + 1).Points
        Players(i).LastEvent = Players(i + 1).LastEvent
        
        For j = 0 To NumKills
            Players(i).KillsWith(j) = Players(i + 1).KillsWith(j)
        Next j
    Next i
End If

'Form6.ListView1.ListItems.Remove NumPlayers

NumPlayers = NumPlayers - 1

End Sub

Sub RemoveEvent(Num)
If DebugMode Then LastCalled = "RemoveEvent"

'trashes an event from the event list

If Num < NumEvents And NumEvents > 1 Then
    For i = Num To NumEvents - 1
        Players(i).Class = Players(i + 1).Class
        Events(i).ComPara = Events(i + 1).ComPara
        For j = 0 To 6
            Events(i).Days(j) = Events(i + 1).Days(j)
        Next j
        Events(i).Every = Events(i + 1).Every
        Events(i).FirstCheck = Events(i + 1).FirstCheck
        Events(i).mde = Events(i + 1).mde
        Events(i).Name = Events(i + 1).Name
        Events(i).ScriptName = Events(i + 1).ScriptName
        Events(i).Times = Events(i + 1).Times
        Events(i).WhatToDo = Events(i + 1).WhatToDo
    
    Next i
End If

NumEvents = NumEvents - 1

'redim the array
ReDim Preserve Events(0 To NumEvents)

End Sub


Function FindKillsWith(Ent As String) As Integer
If DebugMode Then LastCalled = "FindKillsWith"

'returns the kill number of the kill with this ent name

For i = 1 To NumKills
    If InStr(1, LCase(KillList(i).Ent), LCase(Ent)) Then j = i: Exit For
Next i

FindKillsWith = j

End Function


Sub KickPlayer(UIN$, Index As Integer)
If DebugMode Then LastCalled = "KickPlayer"

'kicks a player on the server

For i = 1 To NumPlayers
    If Players(i).UserID = Val(UIN$) Then j = i: Exit For
Next i

If j = 0 Then
    SendPacket "MS", "User not found, therefore could not kick.", Index
    Exit Sub
End If

'kick him
SendRCONCommand "kick # " + Ts(Players(j).UserID)
SendPacket "MS", "User kicked successfully.", Index

End Sub


Sub KillTempMapCycleFile()
If DebugMode Then LastCalled = "KillTempMapCycleFile"

'If CheckForFile(Server.BothPath + "\" + Vars.TempMapCycleFile) Then Kill Server.BothPath + "\" + Vars.TempMapCycleFile

'set the real one again
If Vars.MapCycleFile = "" Or Vars.MapCycleFile = "-1" Then Vars.MapCycleFile = "mapcycle.txt"

Vars.TempFileMade = False

End Sub

Sub BanPlayer(UID$, Index As Integer, UserNum)
If DebugMode Then LastCalled = "BanPlayer"

'bans a player on the server

For i = 1 To NumPlayers
    If Players(i).UserID = Val(UID$) Then j = i: Exit For
Next i

If j = 0 Then
    SendPacket "MS", "User not found, therefore could not ban.", Index
    Exit Sub
End If

'ban him
BanPlayerReason "", "", "Player banned by " + Vars.UserName + " via CLIENT login.", CInt(j)
SendPacket "MS", "User banned and kicked successfully.", Index

End Sub

Sub UpdatePlayerList()
If DebugMode Then LastCalled = "UpdatePlayerList"

'sends to all clients the new player list

For i = 1 To NumConnectUsers
    If CheckBit(i, 13) Then PackagePlayers ConnectUsers(i).Index
Next i

End Sub

Sub UpdateUsersList()
If DebugMode Then LastCalled = "UpdateUsersList"

'sends to all clients the new player list

For i = 1 To NumConnectUsers
    If CheckBit(i, 15) Then PackageConnectUsers ConnectUsers(i).Index
Next i

End Sub


Function FindPlayer(UsID$) As Integer
If DebugMode Then LastCalled = Replace(LastCalled, " - in/after FindPlayer", "") + " - in/after FindPlayer"

'searches the player records for a certain player

For i = 1 To NumPlayers
    If Players(i).UserID = Val(UsID$) Then j = i: Exit For
Next i

FindPlayer = j

End Function

Function FindClan(Num) As Integer
If DebugMode Then LastCalled = "FindClan"

'searches the clan list to see if this player is in this clan

n$ = Players(Num).UniqueID
Dim FoundClan As Integer
Dim FoundFlags As Long

For i = 1 To NumClans
    For j = 1 To Clans(i).NumMembers
        If InStr(1, Clans(i).Members(j).UIN, n$ + ";") Then
            If Clans(i).Flags >= FoundFlags Then
                FoundClan = i
                FoundFlags = Clans(i).Flags
            End If
        End If
    Next j
Next i

FindClan = FoundClan

End Function

Function CheckMember(Num, Cln) As Boolean
If DebugMode Then LastCalled = "CheckMember"

'checks if this player is really a member of this clan

For i = 1 To Clans(Cln).NumMembers
    If InStr(1, Clans(Cln).Members(i).UIN, Players(Num).UniqueID) Then CheckMember = True: Exit Function
Next i

End Function

Sub RealNameSearch(Num, Optional AddTime As Boolean)
If DebugMode Then LastCalled = "RealNameSearch"

'searches the real player lists for someone with this UNIQUEID and then returns the players real name

j = 0
For i = 1 To NumRealPlayers
    If RealPlayers(i).RealName = Players(Num).Name Then k = i
    If InStr(1, RealPlayers(i).UniqueID, Chr(34) + Players(Num).UniqueID + Chr(34) + ";") Then j = i: Exit For
Next i

Players(Num).TempRealMode = False
Players(Num).RealName = ""

If j > 0 Then
    'set the last name
    RealPlayers(j).LastName = Players(Num).Name
    RealPlayers(j).LastTime = CDbl(Now)
    If CheckBit2(RealPlayers(j).Flags, 5) Then Players(Num).TempRealMode = True
    Players(Num).Points = Val(RealPlayers(j).Points)
    Players(Num).RealName = RealPlayers(j).RealName
Else
    'Player not found. See if we are supposed to add...
    If k > 0 And Players(Num).Name <> "Player" And CheckBit2(General.Flags, 6) And AddTime Then
        'OK, add.
        If CheckBit2(RealPlayers(k).Flags, 5) = False Then
        
            'count number of id's
                        
            Dim Tem() As String
            Tem = Split(RealPlayers(k).UniqueID, ";")
            nmid = UBound(Tem)
            
            If CheckBit2(RealPlayers(k).Flags, 7) And nmid > 3 Then
                AddToLogFile "DEBUG: REALADD: Player " + Players(Num).Name + " NOT ADDED TO EXISTING PLAYER " + RealPlayers(k).RealName + " (player UID: " + Players(Num).UniqueID + ") because the original player already had 3 ids."
                Exit Sub
            End If
        
            RealPlayers(k).UniqueID = RealPlayers(k).UniqueID + Players(Num).UniqueID + "; "
            RealPlayers(k).LastName = Players(Num).Name
            RealPlayers(k).LastTime = CDbl(Now)
            Players(Num).Points = Val(RealPlayers(k).Points)
            Players(Num).RealName = RealPlayers(k).RealName
            AddToLogFile "DEBUG: REALADD: Player " + Players(Num).Name + " auto-added to already existing realplayer " + RealPlayers(k).RealName + " (player UID: " + Players(Num).UniqueID + ")"
        End If
    End If
End If

End Sub



Function RealNameSearch2(unid$) As Integer
If DebugMode Then LastCalled = "RealNameSearch2"

'searches the real player lists for someone with this UNIQUEID and then returns the players real name

For i = 1 To NumRealPlayers
    If InStr(1, RealPlayers(i).UniqueID, Chr(34) + unid$ + Chr(34) + ";") Then j = i: Exit For
Next i

RealNameSearch2 = j

End Function


Sub ReadLogOld(Lg$, mde)
If DebugMode Then LastCalled = "ReadLogOld"

'Log Opener
'MDE:
'0 - Normal
'1 - Get Current Map


'Map Name Format:
'L 10/14/1999 - 21:10:17: Spawning server "fortsake"
Dim GetRCON As Integer

GetRCON = 0
If LastCommand = "" Then GetRCON = 1

If CheckForFile(Lg$) Then
    
    Open Lg$ For Input As #1
        Do While Not (EOF(1))
        
            Line Input #1, B$
            B$ = Trim(B$)
            full$ = B$
            If Len(B$) > 25 Then
                B$ = Right(B$, Len(B$) - 25)
            Else
                GoTo end1
            End If
            
            'the MAP thing
            If mde = 1 Then
                Var$ = "Spawning server"
                If Len(B$) > Len(Var$) Then
                    If UCase(Left(B$, Len(Var$))) = UCase(Var$) Then
                        d$ = Right(B$, Len(B$) - Len(Var$))
                        d$ = Trim(d$)
                        If Len(d$) > 0 Then
                            If Left(d$, 1) = Chr(34) Then d$ = Right(d$, Len(d$) - 1)
                            If Right(d$, 1) = Chr(34) Then d$ = Left(d$, Len(d$) - 1)
                        End If
                        Vars.Map = d$
                        mde = 0
                    End If
                End If
            End If
            
            'Find RCON Commands
            
            'RCON Format:
            'L 10/14/1999 - 21:13:52: Rcon from "24.64.165.49:27005":"(rcon "blight" "start")"

            
            Var$ = "Rcon from"
            If Len(B$) > Len(Var$) Then
                
                If GetRCON = 1 Then
                    
                    If UCase(Left(B$, Len(Var$))) = UCase(Var$) Then
                        LastCommand = full$
                        d$ = Right(B$, Len(B$) - Len(Var$))
                        d$ = Trim(d$)
                        
                        If Len(d$) > 0 Then
                            'First get the IP
                            e = InStr(1, d$, Chr(34))
                            f = InStr(1, d$, ":")
                            e = e + 1
                            If f > 0 Then
                                p$ = Mid(d$, e, f - e)
                                Vars.UserIP = p$ 'Got the IP
                            
                            
                                'Now extract the command stuff.
                                f = InStr(f + 1, d$, ":")
                                f = InStr(f + 2, d$, Chr(34))
                                f = InStr(f + 1, d$, Chr(34))
                                f = InStr(f + 1, d$, Chr(34)) 'Got starting location
                                
                                If f > 0 Then
                                   
                                    'p$ = Right(d$, Len(d$) - f)
                                    G = InStrRev(d$, Chr(34))
                                    G = G - 1
                                    If G > f Then
                                        p$ = Mid(d$, f, G - f)
                                        'We have the command now in this format:
                                        '"changelevel" "undertow"
                                        '(still with quotes)
                                        
                                        'So now remove the quotes:
                                        p$ = Replace(p$, Chr(34), "")
                                        
                                        'MsgBox ">>>" + p$ + "<<<"
                                        
                                        'Now we must extract the parameters.
                                        'We can see if there are any by finding the number of spaces.
                                        
                                        p$ = Trim(p$)
                                        
                                        e = InStr(1, p$, " ")
                                                                                
                                        If e = 0 Then 'No paramerers.
                                            Vars.Command = p$
                                        Else
                                            'There are parameters.
                                            'Start extracting them.
                                            
                                            'First the command:
                                            Vars.Command = Left(p$, e - 1)
                                            
                                            'Now the others:
                                            NumParams = 0
                                            
                                            o$ = Vars.Command + ">>>"
                                            
                                            i = 0
                                            Do
                                                G = e
                                                e = InStr(e + 1, p$, " ")
                                                If e = 0 Then e = Len(p$)
                                                
                                                i = i + 1
                                                r$ = Mid(p$, G + 1, e - G)
                                                r$ = Trim(r$)
                                                                                               
                                                'Now replace special commands in the parameter.
                                                r$ = Replace(r$, "\\", "\")
                                                r$ = Replace(r$, "\n", vbCrLf)
                                                r$ = Replace(r$, "_", " ")
                                                r$ = Replace(r$, "\q", Chr(34))
                                                r$ = Replace(r$, "\t", "~")
                                                r$ = Replace(r$, "\s", ";")
                                                r$ = Replace(r$, "%hlpath%", Server.HLPath)
                                                r$ = Replace(r$, "%gamedir%", Server.GamePath)
                                                r$ = Replace(r$, "%password%", Server.RCONPass)
                                                r$ = Replace(r$, "%map%", Vars.Map)
                                                r$ = Replace(r$, "%userip%", Vars.UserIP)
                                                r$ = Replace(r$, "%newestlog%", NewLastLog)
                                                'r$ = ReplaceUserVars(r$)
                                                
                                                Params(i) = r$
                                                
                                                'DEBUG STUFF
                                                o$ = o$ + "<>" + r$
                                                
                                            Loop Until e >= Len(p$)
                                            
                                            NumParams = i
                                            
                                            o$ = o$ + "<<<"
                                            'MsgBox o$
                                        End If
                                        
                                        
                                        'TIME TO EXECUTE!
                                        If Vars.Command <> "" Then
                                            AddToLog "Starting " + Vars.Command + vbCrLf
                                                                                        
                                            'ExecuteScript Vars.Command
                                            
                                        End If
                                    End If
                                End If
                            End If
                            
                        End If
                        '"24.64.165.49:27005":"(rcon "blight" "changelevel" "undertow")"
                    End If
                End If
                
                If full$ = LastCommand Then GetRCON = 1
                                
            End If
            
end1:
        Loop
    Close #1

End If

End Sub
Function SpecialCommandCheck(p$, UserParms() As typParams, NumUserParams, scriptdata As typScriptData, Optional NoSayMode As Boolean, Optional Scret As String) As Boolean
If DebugMode Then LastCalled = "SpecialCommandCheck"

p$ = LCase(p$)

If LCase(p$) = "mapvote" Then
    If LastUser <> 0 Then
        If CheckBit(LastUser, 3) Then
            StartMapVote scriptdata
            SpecialCommandCheck = True
        Else
            SendToUser "Sorry, " + Users(LastUser).Name + "! You aren't allowed to start mapvotes!", scriptdata
        End If
    Else
        StartMapVote scriptdata
        SpecialCommandCheck = True
    End If
ElseIf LCase(p$) = "help" Then
    
    HelpFeature UserParms, NumUserParams, scriptdata
    SpecialCommandCheck = True

ElseIf LCase(p$) = "about" Then
    
    If Vars.ClanBattle = False Then
        SendMessage "This Server uses Avatar-X's Server Assistant version " + Ts(App.Major) + "." + Ts(App.Minor) + "." + Ts(App.Revision), 1, 255, 50, 50, 1, 255, 1, 2, 7, 0.1, 0.02, 5, -1, 0.6
        SpecialCommandCheck = True
    End If
ElseIf p$ = "currmap" Then
    SpecialCommandCheck = True
    SendToUser "Current map is " + Replace(Vars.Map, Chr(34), ""), scriptdata
ElseIf p$ = "timeleft" Then
    AskTimeRemaining
    SpecialCommandCheck = True
    CopyScriptData scriptdata, BanScriptData
ElseIf p$ = "talk" Then 'Talk with the user name first
    
    Msg$ = ""
    For i = 1 To NumUserParams
        Msg$ = Msg$ + UserParms(i).Value + " "
    Next i
    
    For i = 1 To NumUsers
        If UCase(Users(i).Name) = UCase(scriptdata.UserName) Then j = i: Exit For
    Next i
    
    If j > 0 Then
        
        If CheckBit(j, 34) Then
            ExecFunctionScript "spec_chatonlytalk", 2, (scriptdata.UserName), Msg$
            DoTalk "<CHAT " + (scriptdata.UserName) + "> " + Msg$, Len("<CHAT " + (scriptdata.UserName) + ">")
        Else
            Chat Msg$, scriptdata.UserName
        End If
    End If
    
    SpecialCommandCheck = True
ElseIf p$ = "say" And NoSayMode = False Then  'Say
    
    If scriptdata.IsRCON = False Then
        Msg$ = ""
        For i = 1 To NumUserParams
            Msg$ = Msg$ + UserParms(i).Value + " "
        Next i
        SendRCONCommand "say " + Msg$, 2
        SpecialCommandCheck = True
    End If
    
ElseIf p$ = "anntimeleft" Then 'Announce the time remaining
    AnnTime = True
    AskTimeRemaining
    SpecialCommandCheck = True
ElseIf p$ = "setvote" Then 'sets the voted map
    SpecialCommandCheck = True
    If NumUserParams = 0 Then
        If Vars.VotedMap = "" Then
            SendToUser "Next map is not determined.", scriptdata
        Else
            SendToUser "Next map is " + Vars.VotedMap, scriptdata
        End If
    Else
        Vars.VotedMap = Trim(UserParms(1).Value)
        If General.MapVoteMode = "3" Then 'use SA_NEXTMAP
            SendRCONCommand "sa_nextmap " + Vars.VotedMap
        End If
        
        SendToUser "Next map set to " + UserParms(1).Value, scriptdata
    End If
ElseIf p$ = "changelevel" Then 'changes the level
    
    
    If NumUserParams <> 0 And scriptdata.IsRCON = False Then
        SpecialCommandCheck = True
        Vars.VotedMap = ""
        SendToUser "Attempting to start map " + UserParms(1).Value + " ...", scriptdata
        SendRCONCommand "changelevel " + UserParms(1).Value
    End If

ElseIf p$ = "map" Then 'changes the level
    If NumUserParams <> 0 And scriptdata.IsRCON = False Then
        SpecialCommandCheck = True
        Vars.VotedMap = ""
        SendToUser "Attempting to start map " + UserParms(1).Value + " ...", scriptdata
        SendRCONCommand "map " + UserParms(1).Value
    End If

ElseIf p$ = "nextmap" Then 'change to voted map
    SpecialCommandCheck = True
    If Vars.VotedMap <> UCase(Vars.Map) And Vars.ClanBattle = False Then
        'change the map
        SendRCONCommand General.MapChangeMode + " " + Vars.VotedMap
        Vars.VotedMap = ""
    End If
ElseIf p$ = "rcon_password" Then 'the user is trying to change the rcon_password
    
'    If NumUserParams <> 0 And ScriptData.IsRCON = True Then
'        B$ = UserParms(1).Value
'        c$ = Server.RCONPass
'        'so we change to the pass, reset it, then change back
'        If B$ <> c$ Then
'            Server.RCONPass = B$
'            SendRCONCommand "rcon_password " + c$
'            Server.RCONPass = c$
'        End If
'    End If
ElseIf p$ = "autovote" Then
    'turn off autovoting?
    If NumUserParams = 0 Then
    
        If General.NoAutoVotes = True Then B$ = "Automatic map votes are off."
        If General.NoAutoVotes = False Then B$ = "Automatic map votes are on."
    
    Else
        If UCase(UserParms(1).Value) = "1" Or UCase(UserParms(1).Value) = "ON" Then
            General.NoAutoVotes = False
            B$ = "Automatic map votes enabled."
        End If
        If UCase(UserParms(1).Value) = "0" Or UCase(UserParms(1).Value) = "OFF" Then
            General.NoAutoVotes = True
            B$ = "Automatic map votes disabled."
        End If
    End If
    SendToUser B$, scriptdata
    Scret = B$
    SpecialCommandCheck = True
ElseIf p$ = "kickvote" Then
    'turn off kickvoting?
    If NumUserParams = 0 Then
    
        If General.NoKickVotes = True Then B$ = "Kick votes are off."
        If General.NoKickVotes = False Then B$ = "Kick votes are on."
    
    Else
        If UCase(UserParms(1).Value) = "1" Or UCase(UserParms(1).Value) = "ON" Then
            General.NoKickVotes = False
            B$ = "Kick votes enabled."
        End If
        If UCase(UserParms(1).Value) = "0" Or UCase(UserParms(1).Value) = "OFF" Then
            General.NoKickVotes = True
            B$ = "Kick votes disabled."
        End If
    End If
    SendToUser B$, scriptdata
    Scret = B$
    SpecialCommandCheck = True
ElseIf p$ = "stopkickvote" Then
    'stop kickvote
    If KickVoteTimer > 0 Then
        KickVoteUser = 0
        KickVoteTimer = 0
        SendRCONCommand "say " + General.AutoAdminName + " The KickVote was stopped."
        If CheckBit2(General.Flags, 3) And DLLEnabled = True Then SendActualRcon "sa_abort 1": SendActualRcon SA_CHECK
        SendToUser "Kick Vote Stopped.", scriptdata
    End If
    SpecialCommandCheck = True

ElseIf p$ = "stopmapvote" Then
    'stop mapvote
    If MapVoteTimer > 0 Then
        
        MapVoteTimer = 0
        If DLLEnabled = True Then
            SendMessage "The Map Vote was STOPPED by an Administrator.", 1, 1, 255, 1, 1, 1, 1, 0, 8, 1, 3, 3, -1, 0.25
        Else
            SendRCONCommand "say The mapvote was stopped."
        End If
        
        SendToUser "Map Vote Stopped.", scriptdata
    End If
    SpecialCommandCheck = True

ElseIf p$ = "stopchoosevote" Then
    'stop choosevote
    If ChooseVoteTimer > 0 Then
        ChooseVoteTimer = 0
        SendRCONCommand "say " + General.AutoAdminName + " The General Vote was stopped."
        SendToUser "General Vote Stopped.", scriptdata
        If CheckBit2(General.Flags, 4) And DLLEnabled = True Then SendActualRcon "sa_abort 1": SendActualRcon SA_CHECK
    End If
    SpecialCommandCheck = True
ElseIf p$ = "users2" Then
    'REAL user list
    'userid : uniqueid : name
    '------ : ---------: ----
    B$ = "userid : uniqueid : name : real name" + vbCrLf
    B$ = B$ + "------ : -------- : ---- : ---------" + vbCrLf
        
    For i = 1 To NumPlayers
        bb$ = " : " + Players(i).RealName
        If Players(i).RealName = "" Then bb$ = ""
        B$ = B$ + "   " + Ts(Players(i).UserID) + " : " + Players(i).UniqueID + " : " + Players(i).Name + bb$ + vbCrLf
    Next i
    B$ = B$ + Ts(NumPlayers) + " users." + vbCrLf
    
    SendToUser B$, scriptdata, True
    SpecialCommandCheck = True
    
ElseIf p$ = "settimeleft" Then
    'sets amount of time left in a map in MINUTES
    'the time left is the total time minus the time elapsed.
    tl1 = Val(UserParms(1).Value) * 60
    el = (Vars.MapTimeElapsed * 60) + MapCounter
    
    tm = ((el + tl1) / 60)
    tm = Round(tm, 3)
    SendRCONCommand "mp_timelimit " + Ts(tm)
    SendToUser "Time Limit set to " + Ts(tm), scriptdata
    
    SpecialCommandCheck = True
ElseIf p$ = "whois" Then
    'gets real name on a player
    If NumUserParams = 0 Then
        SendToUser "Usage: whois <partial player name>", scriptdata
    Else
        B$ = ""
        For i = 1 To NumUserParams
            B$ = B$ + UserParms(i).Value + " "
        Next i
        B$ = Trim(B$)
        For i = 1 To NumPlayers
            If LCase(LeftR(Players(i).Name, Len(B$))) = LCase(B$) Then Num = Num + 1: lastnum = i
            If LCase(B$) = LCase(Players(i).Name) Then specmd = 1: specnm = i
        Next i
        
        If specmd = 1 Then
            Num = 1
            lastnum = specnm
        End If
        
        If Num <> 1 And NumUserParams = 2 Then
            If UserParms(1).Value = "#" Then
                lsst = FindPlayer(UserParms(2).Value)
                If lsst > 0 Then
                    Num = 1: lastnum = lsst
                Else
                    SendToUser "UserID " + UserParms(2).Value + " not found!", scriptdata
                    SpecialCommandCheck = True
                    Exit Function
                End If
            End If
        End If
        
        If Num = 0 Then
            SendToUser "Player not found.", scriptdata
        ElseIf Num > 1 Then
            SendToUser "More than one match, be more specific.", scriptdata
        Else
            If Players(lastnum).RealName <> "" Then SendToUser "Player " + Players(lastnum).Name + " is really " + Players(lastnum).RealName, scriptdata
            If Players(lastnum).RealName = "" Then
                If Players(lastnum).EntryName = "" Then
                    SendToUser "Player " + Players(lastnum).Name + " is not known.", scriptdata
                Else
                    SendToUser "Player " + Players(lastnum).Name + " is not known, but they entered the game under the name " + Players(lastnum).EntryName, scriptdata
                End If
            End If
        End If
    End If
    SpecialCommandCheck = True

ElseIf p$ = "annid" Or p$ = "annid2" Then
    'gets real name on a player
    If NumUserParams = 0 Then
        If p$ = "annid" Then SendToUser "Usage: annid <partial player name>", scriptdata
    Else
        B$ = ""
        For i = 1 To NumUserParams
            B$ = B$ + UserParms(i).Value + " "
        Next i
        B$ = Trim(B$)
                
        For i = 1 To NumPlayers
            If LCase(LeftR(Players(i).Name, Len(B$))) = LCase(B$) Then Num = Num + 1: lastnum = i
            If LCase(B$) = LCase(Players(i).Name) Then specmd = 1: specnm = i
        Next i
        
        If specmd = 1 Then
            Num = 1
            lastnum = specnm
        End If
        
        If Num <> 1 And NumUserParams = 2 Then
            If UserParms(1).Value = "#" Then
                lsst = FindPlayer(UserParms(2).Value)
                If lsst > 0 Then
                    Num = 1: lastnum = lsst
                Else
                    SendToUser "UserID " + UserParms(2).Value + " not found!", scriptdata
                    SpecialCommandCheck = True
                    Exit Function
                End If
            End If
        End If
        
        If Num = 0 Then
            SendToUser "Player not found.", scriptdata
        ElseIf Num > 1 Then
            SendToUser "More than one match, be more specific.", scriptdata
        Else
            jj = RealNameSearch2(Players(lastnum).UniqueID)
            
            If jj > 0 Then If CheckBit2(RealPlayers(jj).Flags, 3) Then kn = 1
            
            If p$ = "annid2" And kn = 1 Then
                If Players(lastnum).RealName <> "" Then
                    LastTalk = 1
                    SendRCONCommand "say ID Result: Player " + Players(lastnum).Name, , 1
                    SendRCONCommand "say ID Result: Is really " + Players(lastnum).EntryName
                    Scret = "1"
                End If
            Else
                If Players(lastnum).RealName <> "" Then
                    LastTalk = 1
                    SendRCONCommand "say ID Result: Player " + Players(lastnum).Name, , 1
                    SendRCONCommand "say ID Result: Is really " + Players(lastnum).RealName
                    Scret = "1"
                End If
            End If
            If Players(lastnum).RealName = "" Then
                If Players(lastnum).EntryName <> "" Then
                    LastTalk = 1
                    SendRCONCommand "say ID Result: Player " + Players(lastnum).Name, , 1
                    SendRCONCommand "say ID Result: Is NOT known, but joined as " + Players(lastnum).EntryName
                    Scret = "1"
                Else
                    SendToUser "Player " + Players(lastnum).Name + " is not known.", scriptdata
                End If
            End If
        End If
    End If
    
    SpecialCommandCheck = True

ElseIf p$ = "addreal" Then
    'adds a "real" player
    If NumUserParams = 0 Then
        SendToUser "Usage: addreal <partial player name>", scriptdata
    Else
        B$ = ""
        For i = 1 To NumUserParams
            B$ = B$ + UserParms(i).Value + " "
        Next i
        B$ = Trim(B$)
        For i = 1 To NumPlayers
            If LCase(LeftR(Players(i).Name, Len(B$))) = LCase(B$) Then Num = Num + 1: lastnum = i
            If LCase(B$) = LCase(Players(i).Name) Then specmd = 1: specnm = i
        Next i
        
        If specmd = 1 Then
            Num = 1
            lastnum = specnm
        End If
        
        If Num <> 1 And NumUserParams = 2 Then
            If UserParms(1).Value = "#" Then
                lsst = FindPlayer(UserParms(2).Value)
                If lsst > 0 Then
                    Num = 1: lastnum = lsst
                Else
                    SendToUser "UserID " + UserParms(2).Value + " not found!", scriptdata
                    SpecialCommandCheck = True
                    Exit Function
                End If
            End If
        End If
        
        If Num = 0 Then
            SendToUser "Player not found.", scriptdata
        ElseIf Num > 1 Then
            SendToUser "More than one match, be more specific.", scriptdata
        Else
            'add this dude
            bfh = AddRealPlayer2(Players(lastnum).Name, Players(lastnum).UniqueID)
            DoPlayerScan
            SaveCommands
            If bfh Then SendToUser "Real Player " + Players(lastnum).Name + " Added.", scriptdata
            If Not bfh Then SendToUser "Error Adding Real Player " + Players(lastnum).Name + ".", scriptdata
        End If
    End If
    SpecialCommandCheck = True
ElseIf p$ = "addentry" Then
    'adds a "real" player, using their entry name
    If NumUserParams = 0 Then
        SendToUser "Usage: addentry <partial player name>", scriptdata
    Else
        B$ = ""
        For i = 1 To NumUserParams
            B$ = B$ + UserParms(i).Value + " "
        Next i
        B$ = Trim(B$)
        For i = 1 To NumPlayers
            If LCase(LeftR(Players(i).Name, Len(B$))) = LCase(B$) Then Num = Num + 1: lastnum = i
            If LCase(B$) = LCase(Players(i).Name) Then specmd = 1: specnm = i
        Next i
        
        If specmd = 1 Then
            Num = 1
            lastnum = specnm
        End If
        
        If Num <> 1 And NumUserParams = 2 Then
            If UserParms(1).Value = "#" Then
                lsst = FindPlayer(UserParms(2).Value)
                If lsst > 0 Then
                    Num = 1: lastnum = lsst
                Else
                    SendToUser "UserID " + UserParms(2).Value + " not found!", scriptdata
                    SpecialCommandCheck = True
                    Exit Function
                End If
            End If
        End If
        
        If Num = 0 Then
            SendToUser "Player not found.", scriptdata
        ElseIf Num > 1 Then
            SendToUser "More than one match, be more specific.", scriptdata
        Else
            'add this dude
            AddRealPlayer2 Players(lastnum).EntryName, Players(lastnum).UniqueID
            DoPlayerScan
            SaveCommands
            SendToUser "Real Player Added, using ENTRY name.", scriptdata
        End If
    End If
    SpecialCommandCheck = True

ElseIf p$ = "addclan" Then
    'adds a "clan" player to his clan
    If NumUserParams = 0 Then
        SendToUser "Usage: addclan <partial player name>", scriptdata
    Else
        B$ = ""
        For i = 1 To NumUserParams
            B$ = B$ + UserParms(i).Value + " "
        Next i
        B$ = Trim(B$)
        For i = 1 To NumPlayers
            If LCase(LeftR(Players(i).Name, Len(B$))) = LCase(B$) Then Num = Num + 1: lastnum = i
            If LCase(B$) = LCase(Players(i).Name) Then specmd = 1: specnm = i
        Next i
        
        If specmd = 1 Then
            Num = 1
            lastnum = specnm
        End If
        
        If Num <> 1 And NumUserParams = 2 Then
            If UserParms(1).Value = "#" Then
                lsst = FindPlayer(UserParms(2).Value)
                If lsst > 0 Then
                    Num = 1: lastnum = lsst
                Else
                    SendToUser "UserID " + UserParms(2).Value + " not found!", scriptdata
                    SpecialCommandCheck = True
                    Exit Function
                End If
            End If
        End If
        
        If Num = 0 Then
            SendToUser "Player not found.", scriptdata
        ElseIf Num > 1 Then
            SendToUser "More than one match, be more specific.", scriptdata
        Else
            'add this dude
            AddClanPlayer2 lastnum, scriptdata
            SaveCommands
        End If
    End If
    SpecialCommandCheck = True

ElseIf p$ = "banid" Then
    'adds a "clan" player to his clan
    If NumUserParams < 2 Then
        SendToUser "Usage: banid <time> <uniqueid> [kick]", scriptdata
    Else
        
        If Val(UserParms(1).Value) = 0 Then
        
            For i = 1 To NumPlayers
                If Players(i).UniqueID = UserParms(2).Value Then j = i
            Next i
            
            If j > 0 Then
                BanPlayerReason "", "", "Banned by " + Vars.UserName + ".", CInt(j)
            Else
                BanPlayerReason UserParms(2).Value, "Unknown", "Banned by " + Vars.UserName + "."
            End If
        
        Else
            SendRCONCommand "banid " + UserParms(1).Value + " " + UserParms(2).Value + " " + UserParms(3).Value
        End If
    End If
    SpecialCommandCheck = True

ElseIf p$ = "addcommand" Then
        
    'adds a command to a user, like allowing someone to kick
    
    For i = 1 To NumUsers
        If UCase(Users(i).Name) = UCase(scriptdata.UserName) Then j = i: Exit For
    Next i
    
    If j > 0 Then
        
        If CheckBit(j, 11) Then
            'is allowed
            If NumUserParams < 2 Then
                SendToUser "Usage: addcommand <username> <command to add>", scriptdata
            Else
                B$ = UserParms(1).Value
                c$ = UserParms(2).Value
                
                For i = 1 To NumUsers
                    If UCase(Users(i).Name) = UCase(B$) Then k = i: Exit For
                Next i
                
                If k > 0 Then
                    If CheckIfInStr(Users(k).Allowed, c$) = False Then
                        'add it
                        Users(k).Allowed = Users(k).Allowed + vbCrLf + c$
                        SendToUser "Command " + c$ + " added for admin " + B$ + "!", scriptdata
                    Else
                        SendToUser "The administrator " + B$ + " already has command " + c$ + " registered.", scriptdata
                    End If
                Else
                    SendToUser "The administrator " + B$ + " could not be found.", scriptdata
                End If
            End If
        Else
            SendToUser "Sorry, " + scriptdata.UserName + "! You aren't allowed to edit users!", scriptdata
        End If
    
    
    End If
    
    SpecialCommandCheck = True
ElseIf p$ = "getinfo" Then
    'stop kickvote
    GetInfo
    SendToUser "DEBUG - Server Info Retrieved.", scriptdata
    SpecialCommandCheck = True
ElseIf p$ = "startlogwatch" Then
    'stop kickvote
    TempStop = False
    StartLogWatch
    SendToUser "DEBUG - Log Watch STARTED.", scriptdata
    SpecialCommandCheck = True
ElseIf p$ = "stoplogwatch" Then
    'stop kickvote
    TempStop = True
    SendToUser "DEBUG - Log Watch STOPPED.", scriptdata
    SpecialCommandCheck = True
ElseIf LCase(p$) = "startlog" Then
    RconWatchersIP.Add Vars.UserIP
    RconWatchersPort.Add Ts(Vars.UserPort)
    SendToUser "**** Now logging to your console. ****", scriptdata
    SpecialCommandCheck = True
ElseIf LCase(p$) = "stoplog" Then
    
    For i = 1 To RconWatchersIP.count
        If RconWatchersIP(i) = Vars.UserIP Then
            RconWatchersIP.Remove (i)
            RconWatchersPort.Remove (i)
            SendToUser "**** Stopped logging to your console. ****", scriptdata
        End If
    Next i
    
    SpecialCommandCheck = True
ElseIf LCase(p$) = "lastmaps" Then
    
    B$ = ""
    For i = 1 To LastMaps.count
        B$ = B$ + LastMaps(i) + vbCrLf
    Next i
    
    SendToUser B$, scriptdata
    SpecialCommandCheck = True
        
ElseIf LCase(p$) = "clearlastmaps" Then
    
    For i = 1 To LastMaps.count
        LastMaps.Remove 1
    Next i
    
    SendToUser "LastMaps Cleared!", scriptdata
    SpecialCommandCheck = True
    
ElseIf LCase(p$) = "choose" Then
    
    If NumUserParams < 4 Then
        SendToUser "Usage: choose <time in seconds> <number of choices> <choice 1> <choice 2> ..etc.. <choice #> <question>", scriptdata
    ElseIf ChooseVoteTimer <= 0 Then
        ChooseVoteTime = Val(UserParms(1).Value)
        numchoice = Val(UserParms(2).Value)
        If ChooseVoteTime <= 180 And ChooseVoteTime >= 5 And numchoice > 0 Then
                      
            ReDim ChooseVote(1 To numchoice)
                      
            For j = 3 To numchoice + 2
                ChooseVote(j - 2) = UserParms(j).Value
            Next j
            
            B$ = ""
            For i = numchoice + 3 To NumUserParams
                B$ = B$ + UserParms(i).Value + " "
            Next i
    
            ChooseVoteQuestion = B$
        
            StartChooseVote
        
        End If
    End If
    SpecialCommandCheck = True
ElseIf LCase(p$) = "unbanlast" Then
    SendToUser "Unbanning the last ban...", scriptdata
    UnBanLast = True
    
    CopyScriptData scriptdata, BanScriptData
    SendRCONCommand "listid"
    
    SpecialCommandCheck = True

ElseIf LCase(p$) = "sortlogs" Then
    SortLogs
    SendToUser "Logs sorted.", scriptdata
    SpecialCommandCheck = True

ElseIf LCase(p$) = "debugmode" Then
    
    If UserParms(1).Value = "on" Then
        SendToUser "Debug Mode Enabled", scriptdata
        DebugMode = True
    Else
        SendToUser "Debug Mode Disabled", scriptdata
        DebugMode = False
    End If
    
    SpecialCommandCheck = True

ElseIf LCase(p$) = "testproc" Then
    
    SendToUser "Test Proc...", scriptdata
    TestProc scriptdata
    SpecialCommandCheck = True
    SendToUser "Done Test Proc!", scriptdata


ElseIf LCase(p$) = "debugtime" Then

    If UserParms(1).Value = "on" Then
        SendToUser "Debug Time Enabled", scriptdata
        DebugTime = True
    Else
        SendToUser "Debug Time Disabled", scriptdata
        DebugTime = False
    End If
    
    SpecialCommandCheck = True


ElseIf LCase(p$) = "changename" Then
    
    If NumUserParams < 2 Then
        SendToUser "Usage: changename <userid> <new name>", scriptdata
    Else
        B$ = ""
        For i = 2 To NumUserParams
            B$ = B$ + UserParms(i).Value + " "
        Next i
        B$ = Trim(B$)
        
        Num = FindPlayer(UserParms(1).Value)
        
        If Num = 0 Then
            SendToUser "Player not found.", scriptdata
        Else
            ChangePlayerName Players(Num).UserID, B$
        End If
    End If
    SpecialCommandCheck = True

ElseIf LCase(p$) = "changeclass" Then
    
    If NumUserParams < 2 Then
        SendToUser "Usage: changeclass <userid> <new class>" + vbCrLf + "Classes: 1-9 are Scout to Engineer, 10 is Random, 11 is Civilian", scriptdata
    Else
        B$ = UserParms(2).Value
        
        Num = FindPlayer(UserParms(1).Value)
        
        If Num = 0 Then
            SendToUser "Player not found.", scriptdata
        Else
            ChangePlayerClass Players(Num).UserID, Val(B$)
        End If
    End If
    SpecialCommandCheck = True

ElseIf LCase(p$) = "message" Then
        
    B$ = ""
    For i = 1 To NumUserParams
        B$ = B$ + UserParms(i).Value + " "
    Next i
    B$ = Trim(B$)
    
    B$ = Replace(B$, "\n", Chr(10))
    
    SendRCONCommand "say " + B$, 3
    SpecialCommandCheck = True
        
ElseIf LCase(p$) = "message2" Then
        
    B$ = ""
    For i = 1 To NumUserParams
        B$ = B$ + UserParms(i).Value + " "
    Next i
    B$ = Trim(B$)
    
    B$ = Replace(B$, "\n", Chr(10))
    
    SendRCONCommand "say " + B$, 4
    SpecialCommandCheck = True
        
ElseIf LCase(p$) = "running" Then
        
    SendToUser "Getting list of running scripts... please wait 3 seconds...", scriptdata
    
    FindRunningScriptsTime = Now
    FindRunningScripts = True
    RunningScripts = ""
    
    Dim tmel As Double
    strt = Timer
    Do
        DoEvents
        tmel = Round(Timer - strt, 3)
        If ScriptCheck(scriptdata) Then Exit Function
    Loop Until tmel > 3 Or tmel < 0
        
    FindRunningScripts = False
    
    SendToUser RunningScripts, scriptdata
    
    SendToUser "List retrieved.", scriptdata
    
    SpecialCommandCheck = True
                
        
ElseIf LCase(p$) = "message3" Then
        
    B$ = ""
    For i = 1 To NumUserParams
        B$ = B$ + UserParms(i).Value + " "
    Next i
    B$ = Trim(B$)
    
    B$ = Replace(B$, "\n", Chr(10))
    
    SendRCONCommand "say " + B$, 5
    SpecialCommandCheck = True
        
ElseIf LCase(p$) = "messageflick" Then
        
    B$ = ""
    For i = 1 To NumUserParams
        B$ = B$ + UserParms(i).Value + " "
    Next i
    B$ = Trim(B$)
    
    B$ = Replace(B$, "\n", Chr(10))
    
    SendMessage B$, 1, 255, 1, 1, 1, 1, 255, 1, 6, 0.1, 2, 3, 0.5, 0.2
    SpecialCommandCheck = True
        
ElseIf LCase(p$) = "messagebot" Then
        
    B$ = ""
    For i = 1 To NumUserParams
        B$ = B$ + UserParms(i).Value + " "
    Next i
    B$ = Trim(B$)
    
    B$ = Replace(B$, "\n", Chr(10))
    
    SendMessage B$, 1, 255, 255, 1, 1, 1, 255, 0, 6, 0.1, 2, 3, 0.5, 0.7
    SpecialCommandCheck = True
        
ElseIf LCase(p$) = "saybroad" Then
    
    If NumUserParams < 2 Then
        SendToUser "Usage: saybroad <value> <partial name>" + vbCrLf + "Values: 0 - off, 1 - normal say, 2 - color messages", scriptdata
    Else
        B$ = ""
        For i = 2 To NumUserParams
            B$ = B$ + UserParms(i).Value + " "
        Next i
        B$ = Trim(B$)
        
        For i = 1 To NumPlayers
            If LCase(LeftR(Players(i).Name, Len(B$))) = LCase(B$) Then Num = Num + 1: lastnum = i
            If LCase(B$) = LCase(Players(i).Name) Then specmd = 1: specnm = i
        Next i
        
        If specmd = 1 Then
            Num = 1
            lastnum = specnm
        End If
        
        If Num <> 1 And NumUserParams = 3 Then
            If UserParms(2).Value = "#" Then
                lsst = FindPlayer(UserParms(3).Value)
                If lsst > 0 Then
                    Num = 1: lastnum = lsst
                Else
                    SendToUser "UserID " + UserParms(3).Value + " not found!", scriptdata
                    SpecialCommandCheck = True
                    Exit Function
                End If
            End If
        End If
        
        If Num = 0 Then
            SendToUser "Player not found.", scriptdata
        ElseIf Num > 1 Then
            SendToUser "More than one match, be more specific.", scriptdata
        Else
            nm2 = Int(Val(UserParms(1).Value))
            If nm2 < 0 Or nm2 > 2 Then
                SendToUser "Specify a value between 0 and 2.", scriptdata
            Else
                Players(lastnum).BroadcastType = nm2
                SendToUser Players(lastnum).Name + "'s Broadcasting Mode set to " + Ts(nm2) + ".", scriptdata
            End If
        End If
    End If
    SpecialCommandCheck = True
ElseIf LCase(p$) = "setvar" Then
    
    If NumUserParams < 2 Then
        SendToUser "Usage: setvar <varname> <value>" + vbCrLf + "Note: <value> may contain spaces.", scriptdata
    Else
        B$ = ""
        For i = 2 To NumUserParams
            B$ = B$ + UserParms(i).Value + " "
        Next i
        B$ = Trim(B$)
        
        a$ = Server.BothPath + "\temp.rc"
        If CheckForFile(a$) Then Kill a$
        
        B$ = UserParms(1).Value + " " + Chr(34) + B$ + Chr(34)
        h = FreeFile
Close h
        Open a$ For Append As h
            Print #h, B$
        Close h
        
        'File created, now exec it
        SendRCONCommand "exec temp.rc"
        SendToUser "Command Sent: " + B$, scriptdata
    End If
    SpecialCommandCheck = True
ElseIf LCase(p$) = "dllenabled" Then
    
    If NumUserParams = 0 Then
        If DLLEnabled = True Then SendToUser "The Dll is currently enabled.", scriptdata
        If DLLEnabled = False Then SendToUser "The Dll is currently DISabled.", scriptdata
    Else
        
        If UserParms(1).Value = "on" Then
            DLLEnabled = True
            SendToUser "Dll Enabled.", scriptdata
        ElseIf UserParms(1).Value = "off" Then
            DLLEnabled = False
            SendToUser "Dll DISabled.", scriptdata
        End If
    End If
    SpecialCommandCheck = True

ElseIf LCase(p$) = "checkfordll" Then
    
    CheckForDLL
            
    If DLLEnabled = True Then SendToUser "Checked, The Dll is currently enabled.", scriptdata
    If DLLEnabled = False Then SendToUser "Checked, The Dll is currently DISabled.", scriptdata
    
    SpecialCommandCheck = True
ElseIf LCase(p$) = "debugprint" Then
    
    B$ = "DEBUG PRINTOUT" + vbCrLf
    B$ = B$ + "Number of connected users: " + Ts(NumConnectUsers) + vbCrLf
    For i = 1 To NumConnectUsers
        B$ = B$ + "     " + Ts(i) + "  Index:" + Ts(ConnectUsers(i).Index) + vbCrLf
        B$ = B$ + "     " + Ts(i) + "  IP:" + ConnectUsers(i).IP + vbCrLf
        B$ = B$ + "     " + Ts(i) + "  Usernum:" + Ts(ConnectUsers(i).UserNum) + vbCrLf
        B$ = B$ + "     " + Ts(i) + "  LogLevel:" + Ts(ConnectUsers(i).LogLevel) + vbCrLf
        B$ = B$ + "     " + Ts(i) + "  Name:" + ConnectUsers(i).Name + vbCrLf
    Next i
    
    B$ = B$ + "Log Sort Timer: " + Ts(LogSortTimer) + vbCrLf
    SendToUser B$, scriptdata
    SpecialCommandCheck = True

ElseIf LCase(p$) = "resetmessage" Then

    'ok, now send to server all needed changes
    SendRCONCommand "sa_message_red1 " + Ts(SvMes.Red1)
    SendRCONCommand "sa_message_green1 " + Ts(SvMes.Green1)
    SendRCONCommand "sa_message_blue1 " + Ts(SvMes.Blue1)
    SendRCONCommand "sa_message_red2 " + Ts(SvMes.Red2)
    SendRCONCommand "sa_message_green2 " + Ts(SvMes.Green2)
    SendRCONCommand "sa_message_blue2 " + Ts(SvMes.Blue2)
    SendRCONCommand "sa_message_effect " + Ts(SvMes.Effect)
    SendRCONCommand "sa_message_holdtime " + Ts(SvMes.HoldTime)
    SendRCONCommand "sa_message_fxtime " + Ts(SvMes.FxTime)
    SendRCONCommand "sa_message_fadein " + Ts(SvMes.FadeInTime)
    SendRCONCommand "sa_message_fadeout " + Ts(SvMes.FadeOutTime)
    SendRCONCommand "sa_message_position_x " + Ts(SvMes.X)
    SendRCONCommand "sa_message_position_y " + Ts(SvMes.Y)
    SendRCONCommand "sa_message_dynamic " + Ts(SvMes.Dynamic)
    
    SendToUser "Message Settings Reset.", scriptdata
    SpecialCommandCheck = True
ElseIf LCase(p$) = "doplayerscan" Then
    DoPlayerScan
    SendToUser "Player Scan Complete", scriptdata
    SpecialCommandCheck = True

ElseIf LCase(p$) = "crash" Then
    SendToUser "********** Initiating Server Crash Reset...", scriptdata, True
    
    TempStop = False
    
    ServerCrash
    
    SendToUser "********** Server recovered from crash.", scriptdata, True
    SpecialCommandCheck = True

ElseIf p$ = "setreal" Then
    'gets real name on a player
    If NumUserParams = 0 Then
        SendToUser "Usage: setreal <partial player name>", scriptdata
    Else
        B$ = ""
        For i = 1 To NumUserParams
            B$ = B$ + UserParms(i).Value + " "
        Next i
        B$ = Trim(B$)
                
        For i = 1 To NumPlayers
            If LCase(LeftR(Players(i).Name, Len(B$))) = LCase(B$) Then Num = Num + 1: lastnum = i
            If LCase(B$) = LCase(Players(i).Name) Then specmd = 1: specnm = i
        Next i
        
        If specmd = 1 Then
            Num = 1
            lastnum = specnm
        End If
        
        If Num <> 1 And NumUserParams = 2 Then
            If UserParms(1).Value = "#" Then
                lsst = FindPlayer(UserParms(2).Value)
                If lsst > 0 Then
                    Num = 1: lastnum = lsst
                Else
                    SendToUser "UserID " + UserParms(2).Value + " not found!", scriptdata
                    SpecialCommandCheck = True
                    Exit Function
                End If
            End If
        End If
        
        If Num = 0 Then
            SendToUser "Player not found.", scriptdata
        ElseIf Num > 1 Then
            SendToUser "More than one match, be more specific.", scriptdata
        Else
            If Players(lastnum).RealName <> "" Then
                ChangePlayerName Players(lastnum).UserID, Players(lastnum).RealName
            End If
            If Players(lastnum).RealName = "" Then
                If Players(lastnum).EntryName <> "" Then
                    LastTalk = 1
                    ChangePlayerName Players(lastnum).UserID, Players(lastnum).EntryName
                Else
                    SendToUser "Player " + Players(lastnum).Name + " is not known.", scriptdata
                End If
            End If
        End If
    End If
    SpecialCommandCheck = True

ElseIf p$ = "lookupreal" Then
    'reports when this player was last seen
    If NumUserParams = 0 Then
        SendToUser "Usage: lookupreal <partial realplayer name>", scriptdata
    Else
        B$ = ""
        For i = 1 To NumUserParams
            B$ = B$ + UserParms(i).Value + " "
        Next i
        B$ = Trim(B$)
                
        For i = 1 To NumRealPlayers
            If LCase(LeftR(RealPlayers(i).RealName, Len(B$))) = LCase(B$) Then Num = Num + 1: lastnum = i: lst$ = lst$ + RealPlayers(i).RealName + " - " + Format(CDate(RealPlayers(lastnum).LastTime), "ddd, mmm d yyyy, hh:mm:ss AMPM") + vbCrLf
        Next i
        
        Scret = "0"
        If Num = 0 Then
            SendToUser "Real Player not found.", scriptdata
        ElseIf Num > 1 Then
            SendToUser Ts(Num) + " matches: " + vbCrLf + lst$ + vbCrLf + "Be more specific.", scriptdata
            Scret = Ts(-Num)
        Else
            'report the date
            If RealPlayers(lastnum).LastTime = 0 Then
                SendRCONCommand "say It is not known when " + RealPlayers(lastnum).RealName + " was last seen."
                Scret = "2"
            Else
                Dim tmpdate As Date
                tmpdate = Now - CDate(RealPlayers(lastnum).LastTime)
                B$ = Ts(Int(CDbl(tmpdate))) + " days, " + Ts(Hour(tmpdate)) + " hours, " + Ts(Minute(tmpdate)) + " minutes, " + Ts(Second(tmpdate)) + " seconds"
                
                SendRCONCommand "say Player " + RealPlayers(lastnum).RealName + " was last seen on " + Format(CDate(RealPlayers(lastnum).LastTime), "ddd, mmm d yyyy, hh:mm:ss AMPM"), , 1
                SendRCONCommand "say Which was " + B$ + " ago!"
                
                Scret = "1"
            End If
        End If
    End If
    
    SpecialCommandCheck = True

ElseIf p$ = "talkto" Then 'Talk with the user name first
    If NumUserParams < 2 Then
        SendToUser "Usage: talkto <user id> <message>", scriptdata
    Else
        Msg$ = ""
        For i = 2 To NumUserParams
            Msg$ = Msg$ + UserParms(i).Value + " "
        Next i
        
        TalkToPlayer Val(UserParms(1).Value), "<PRIVATE FROM " + scriptdata.UserName + "> " + Msg$
        
        Num = FindPlayer(UserParms(1).Value)
        If Num > 0 Then
            nm$ = "<PRIVATE " + scriptdata.UserName + " TO " + Players(Num).Name + ">"
            lennm = Len(nm$)
            nm$ = nm$ + " " + Msg$
            DoTalk nm$, CInt(lennm), True
        End If
    End If
    SpecialCommandCheck = True
    
ElseIf p$ = "kill" Then 'Talk with the user name first
    If NumUserParams <> 1 Then
        SendToUser "Usage: kill <user id>", scriptdata
    Else
        RKillPlayer Val(UserParms(1).Value)
    End If
    SpecialCommandCheck = True

ElseIf p$ = "telekill" Then 'Kill this teleporter
    If NumUserParams <> 1 Then
        SendToUser "Usage: telekill <teleporter name>", scriptdata
    Else
        B$ = UserParms(1).Value
        
        Num = FindTele(B$)
               
        If Num = 0 Then
            SendToUser "Teleport " + B$ + " not found!", scriptdata
        Else
            
            For i = Num To NumTele - 1
                Tele(i).Angle = Tele(i + 1).Angle
                Tele(i).X = Tele(i + 1).X
                Tele(i).Y = Tele(i + 1).Y
                Tele(i).Z = Tele(i + 1).Z
                Tele(i).Name = Tele(i + 1).Name
            Next i
            NumTele = NumTele - 1
            ReDim Preserve Tele(0 To NumTele)
            
            SendToUser "Teleport " + B$ + " deleted!", scriptdata
        End If
    End If
    SpecialCommandCheck = True

ElseIf p$ = "teleport" Then 'Teleport a player
    If NumUserParams <> 2 Then
        SendToUser "Usage: teleport <id> <teleporter name>", scriptdata
    Else
        
        B$ = UserParms(2).Value
        B2$ = UserParms(1).Value
        
        If TeleportPlayer(Val(B2$), B$) = False Then
            SendToUser "Teleporter " + B$ + " not found!", scriptdata
        End If
    End If
    SpecialCommandCheck = True

ElseIf p$ = "telename" Then 'Teleport a player
    If NumUserParams <> 2 Then
        SendToUser "Usage: telename <original name> <new name>", scriptdata
    Else
        
        B$ = UserParms(1).Value
        B2$ = UserParms(2).Value
        
        Num = FindTele(B$)
        
        If Num > 0 Then
            Tele(Num).Name = B2$
            SendToUser "Teleporter " + B$ + " renamed to " + B2$ + "!", scriptdata
        Else
            SendToUser "Teleporter " + B$ + " not found!", scriptdata
        End If
        
    End If
    SpecialCommandCheck = True

ElseIf p$ = "teleportto" Then 'Teleport a player
    If NumUserParams <> 4 Then
        SendToUser "Usage: teleport <id> <x> <y> <z>", scriptdata
    Else
        
        B2$ = UserParms(1).Value
        X1 = Val(UserParms(2).Value)
        Y1 = Val(UserParms(3).Value)
        z1 = Val(UserParms(4).Value)
        
        TeleportPlayerTo Val(B2$), CInt(X1), CInt(Y1), CInt(z1)
        
    End If
    SpecialCommandCheck = True


ElseIf p$ = "telelist" Then 'Teleporter list
    
    
    B$ = "Teleporters in map: " + vbCrLf
    For i = 1 To NumTele
        B$ = B$ + Tele(i).Name + " - (" + Ts(Tele(i).X) + ", " + Ts(Tele(i).Y) + ", " + Ts(Tele(i).Z) + ") - " + Ts(Tele(i).Angle) + " degrees" + vbCrLf
    Next i
    B$ = B$ + Ts(NumTele) + " teleporters."
    SendToUser B$, scriptdata
    SpecialCommandCheck = True

ElseIf p$ = "telelistto" Then 'Teleporter list
    
    
    B$ = "Teleporters in map: " + Chr(10)
    For i = 1 To NumTele
        B$ = B$ + Tele(i).Name + Chr(10)
    Next i
    Vars.UserIsTCP = False
    B$ = B$ + Ts(NumTele) + " teleporters."
    SendToUserIP UserParms(1).Value, Val(UserParms(2).Value), B$
    SpecialCommandCheck = True


ElseIf p$ = "telesave" Then 'Teleporter list
    
    SaveTeleFile
    SendToUser Ts(NumTele) + " Teleporters Saved for map " + Vars.Map, scriptdata
    SpecialCommandCheck = True

ElseIf p$ = "teleload" Then 'Teleporter list
    
    LoadTeleFile
    SendToUser Ts(NumTele) + " Teleporters Loaded for map " + Vars.Map, scriptdata
    SpecialCommandCheck = True

ElseIf p$ = "teleadd" Then 'Teleporter list
    If NumUserParams <> 5 Then
        SendToUser "Usage: teleadd <teleporter name> <angle> <x> <y> <z>", scriptdata
    Else

        Num = FindTele(nm$)
        
        If Num = 0 And UserParms(1).Value <> "" Then 'safe to add
            NumTele = NumTele + 1
            ReDim Preserve Tele(0 To NumTele)
        
            Tele(NumTele).Name = UserParms(1).Value
            Tele(NumTele).Angle = Val(UserParms(2).Value)
            
            'get co-ords out of co-ords var
            Tele(NumTele).X = Val(UserParms(3).Value)
            Tele(NumTele).Y = Val(UserParms(4).Value)
            Tele(NumTele).Z = Val(UserParms(5).Value)
            
            AddToLogFile "TELEADD: Teleporter " + nm$ + " added to map " + Vars.Map + " by " + scriptdata.UserName
            SendToUser "Teleporter Added", scriptdata
        Else
            SendToUser "Cannot add teleporter, because a teleporter with this name already exists.", scriptdata
        End If
    End If
    
    SpecialCommandCheck = True
    
ElseIf p$ = "maparraysave" Then

    SaveMapDataFile
    SpecialCommandCheck = True
    
    SendToUser "Map Array Saved!", scriptdata
    
ElseIf p$ = "maparrayload" Then

    LoadMapDataFile
    SpecialCommandCheck = True
    
    SendToUser "Map Array Loaded!", scriptdata
    
    
ElseIf p$ = "maparrayclear" Then

    For X = 0 To 64
        For Y = 0 To 64
            MapArray(X, Y) = 0
        Next Y
    Next X
        
    SendToUser "Map Array Cleared!", scriptdata
    
    SpecialCommandCheck = True
    
ElseIf p$ = "abortscript" Then
    
    ' stop the script execution
    scriptdata.ExitNow = True
    
ElseIf p$ = "stopallscripts" Then

    StopAllScripts = True
    
    SendToUser "Scripts Stopped", scriptdata
    SpecialCommandCheck = True

ElseIf p$ = "clearadminchat22" Then
    
    For i = 1 To 20
        LastChats(i) = ""
    Next i
    
    For i = 1 To NumConnectUsers
        SendPacket "A2", "", ConnectUsers(i).Index
    Next i
    
    SendToUser "AdminChat Cleared", scriptdata
    SpecialCommandCheck = True
    
ElseIf p$ = "resettimer2" Then

    Form1.Timer2.Enabled = False
    DoEvents
    Form1.Timer2.Enabled = True
    
    SendToUser "TIMER2 RESET!", scriptdata

ElseIf p$ = "timer2" Then

    SendToUser "Last run at " + Format(LastTimer2, "dd/mm/yyyy hh:mm:ss") + " doing " + LastTimer2What, scriptdata

ElseIf p$ = "handleentry" Then
    HandleEntry
ElseIf p$ = "resetudp" Then
    'reset the UDP
    
    If DebugMode Then LastCalled = "Resetting UDP..."
    
    
    Form1.UDP1.Close
    SetPorts

ElseIf p$ = "startallscripts" Then

    StopAllScripts = False
    
    SendToUser "Scripts Started", scriptdata
    SpecialCommandCheck = True
    
    
ElseIf p$ = "logbufferlen" Then

    SendToUser "Log Buffer is " + Ts(Len(LogBuffer)), scriptdata
    SpecialCommandCheck = True

ElseIf p$ = "logbufferclear" Then

    SendToUser "Log Buffer is " + Ts(Len(LogBuffer)) + ", cleared.", scriptdata
    LogBuffer = ""
    SpecialCommandCheck = True

ElseIf LCase(p$) = "debugvars" Then
    
    B$ = "DEBUG PRINTOUT - VARS" + vbCrLf
   
    B$ = B$ + "Vars.TimeCounter: " + Ts(Vars.TimeCounter) + vbCrLf
    B$ = B$ + "AdminSpeechBuffer: " + Ts(AdminSpeechBuffer) + vbCrLf
    B$ = B$ + "EventTimer: " + Ts(EventTimer) + vbCrLf
    B$ = B$ + "LogWatchTimer: " + Ts(LogWatchTimer) + vbCrLf
    B$ = B$ + "Vars.MapTimeLeft: " + Ts(Vars.MapTimeLeft) + vbCrLf
    B$ = B$ + "MapCounter: " + Ts(MapCounter) + vbCrLf
    B$ = B$ + "LogSortTimer: " + Ts(LogSortTimer) + vbCrLf
    B$ = B$ + "Vars.MapTimeElapsed: " + Ts(Vars.MapTimeElapsed) + vbCrLf
    B$ = B$ + "Len(LogBuffer): " + Ts(Len(LogBuffer)) + vbCrLf
    
    SendToUser B$, scriptdata
    SpecialCommandCheck = True


ElseIf LCase(p$) = "devoice" Then
    
    If NumUserParams = 0 Then
        SendToUser "Usage: devoice <partial name match | # userid>", scriptdata
    Else
        B$ = ""
        For i = 1 To NumUserParams
            B$ = B$ + UserParms(i).Value + " "
        Next i
        B$ = Trim(B$)
        For i = 1 To NumPlayers
            If LCase(LeftR(Players(i).Name, Len(B$))) = LCase(B$) Then Num = Num + 1: lastnum = i
        Next i
        
        If Num <> 1 And NumUserParams = 2 Then
            If UserParms(1).Value = "#" Then
                lsst = FindPlayer(UserParms(2).Value)
                If lsst > 0 Then
                    Num = 1: lastnum = lsst
                Else
                    SendToUser "UserID " + UserParms(2).Value + " not found!", scriptdata
                    SpecialCommandCheck = True
                    Exit Function
                End If
            End If
        End If
        
        If Num = 0 Then
            SendToUser "Player not found.", scriptdata
        ElseIf Num > 1 Then
            SendToUser "More than one match, be more specific.", scriptdata
        Else
            
            PlayerDeVoice Players(lastnum).UserID
        End If
    End If
    SpecialCommandCheck = True

ElseIf LCase(p$) = "revoice" Then
    
    If NumUserParams = 0 Then
        SendToUser "Usage: revoice <partial name match | # userid>", scriptdata
    Else
        B$ = ""
        For i = 1 To NumUserParams
            B$ = B$ + UserParms(i).Value + " "
        Next i
        B$ = Trim(B$)
        For i = 1 To NumPlayers
            If LCase(LeftR(Players(i).Name, Len(B$))) = LCase(B$) Then Num = Num + 1: lastnum = i
        Next i
        
        If Num <> 1 And NumUserParams = 2 Then
            If UserParms(1).Value = "#" Then
                lsst = FindPlayer(UserParms(2).Value)
                If lsst > 0 Then
                    Num = 1: lastnum = lsst
                Else
                    SendToUser "UserID " + UserParms(2).Value + " not found!", scriptdata
                    SpecialCommandCheck = True
                    Exit Function
                End If
            End If
        End If
        
        If Num = 0 Then
            SendToUser "Player not found.", scriptdata
        ElseIf Num > 1 Then
            SendToUser "More than one match, be more specific.", scriptdata
        Else
            
            PlayerReVoice Players(lastnum).UserID
        End If
    End If
    SpecialCommandCheck = True

ElseIf p$ = "setgrid" Then

    If NumUserParams <> 3 Then
    
        SendToUser "Usage: setgrid <x> <y> <value>", scriptdata
    
    Else
    
        X = Val(UserParms(1).Value)
        Y = Val(UserParms(2).Value)
        Num = Val(UserParms(3).Value)
        
        If X >= 0 And X <= 64 And Y >= 0 And Y <= 64 And Num >= -4096 And Num <= 4096 Then
            MapArray(X, Y) = Num
        Else
            SendToUser "Invalid Values!", scriptdata
        End If
    
    
    End If
    
    SpecialCommandCheck = True

ElseIf p$ = "sendicq" Then

    If NumUserParams < 2 Then
    
        SendToUser "Usage: sendicq <admin name> <message>", scriptdata
    
    Else
        
        B$ = ""
        For i = 2 To NumUserParams
            B$ = B$ + UserParms(i).Value + " "
        Next i
        B$ = Trim(B$)
    
        a$ = UserParms(1).Value
        
        'find admin
        For i = 1 To NumUsers
            If UCase(Users(i).Name) = UCase(a$) Then j = i: Exit For
        Next i
    
        If j = 0 Then
            SendToUser "Admin not found.", scriptdata
        Else
            If Users(j).ICQ = "" Then
                SendToUser "Admins ICQ Unknown", scriptdata
            Else
            
                SendICQMessage Users(j).ICQ, "Message From: " + scriptdata.UserName + vbCrLf + B$
                SendToUser "Message Sent!", scriptdata
            End If
        End If
        
    
    End If
    
    SpecialCommandCheck = True

ElseIf p$ = "sendicqto" Then

    If NumUserParams < 2 Then
    
        SendToUser "Usage: sendicqto <UIN> <message>", scriptdata
    
    Else
        
        B$ = ""
        For i = 2 To NumUserParams
            B$ = B$ + UserParms(i).Value + " "
        Next i
        B$ = Trim(B$)
    
        a$ = UserParms(1).Value
        
        SendICQMessage a$, "Message From: " + scriptdata.UserName + vbCrLf + B$
        SendToUser "Message Sent!", scriptdata
    
    End If
    
    SpecialCommandCheck = True


ElseIf LCase(p$) = "addpoints" Then
    
    If NumUserParams <> 2 Then
        SendToUser "Usage: addpoints <userid> <amount to add>", scriptdata
    Else
        Num = FindPlayer(UserParms(1).Value)
        
        If Num = 0 Then
            SendToUser "Player not found.", scriptdata
        Else
            If Len(UserParms(2).Value) > 8 Then UserParms(2).Value = "0"
            SetPoints Num, GetPoints(Num) + Val(UserParms(2).Value)
        End If
    End If
    SpecialCommandCheck = True

ElseIf LCase(p$) = "rempoints" Then
    
    If NumUserParams <> 2 Then
        SendToUser "Usage: rempoints <userid> <amount to remove>", scriptdata
    Else
        Num = FindPlayer(UserParms(1).Value)
        
        If Num = 0 Then
            SendToUser "Player not found.", scriptdata
        Else
            SetPoints Num, GetPoints(Num) - Val(UserParms(2).Value)
        End If
    End If
    SpecialCommandCheck = True
    
ElseIf LCase(p$) = "setpoints" Then
    
    If NumUserParams <> 2 Then
        SendToUser "Usage: setpoints <userid> <amount>", scriptdata
    Else
        Num = FindPlayer(UserParms(1).Value)
        
        If Num = 0 Then
            SendToUser "Player not found.", scriptdata
        Else
            If Len(UserParms(2).Value) > 8 Then UserParms(2).Value = "0"
            SetPoints Num, Val(UserParms(2).Value)
        End If
    End If
    SpecialCommandCheck = True

ElseIf LCase(p$) = "annpoints" Then
    
    If NumUserParams <> 1 Then
        SendToUser "Usage: annpoints <userid>", scriptdata
    Else
        Num = FindPlayer(UserParms(1).Value)
        
        If Num = 0 Then
            SendToUser "Player not found.", scriptdata
        Else
            SendRCONCommand "say " + General.AutoAdminName + " Player " + Players(Num).Name + " has " + Ts(GetPoints(Num)) + " points!"
        End If
    End If
    SpecialCommandCheck = True
   
ElseIf p$ = "getsendinfo" Then

    SendActualRcon "sa_sendinfo 1"
    SendActualRcon SA_CHECK
    SendToUser "Info Requested.", scriptdata
    
    SpecialCommandCheck = True
ElseIf p$ = "stopserverassist" Then
    
    If CheckForFile(App.Path + "\serverold.exe") = True Then Kill App.Path + "\serverold.exe"
    
    EndProgram 1
    SendToUser "Stopping Server Assistant...", scriptdata

ElseIf p$ = "shell" Then
    
    SendToUser "Shelling: " + UserParms(1).Value, scriptdata
    Shell UserParms(1).Value
    SpecialCommandCheck = True

ElseIf p$ = "showglobalvars" Then
    
    
    For i = 1 To NumUserVars
        B$ = B$ + "  " + UserVars(i).Name + " : " + UserVars(i).Value + Chr(10)
    Next i
    B$ = B$ + "Total " + Ts(NumUserVars) + " variables."
    
    SendToUser B$, scriptdata
    
    SpecialCommandCheck = True

ElseIf p$ = "showlocalvars" Then
    
    
    For i = 1 To scriptdata.NumVars
        B$ = B$ + "  " + scriptdata.VarNames(i) + " : " + scriptdata.VarValues(i) + Chr(10)
    Next i
    B$ = B$ + "Total " + Ts(scriptdata.NumVars) + " variables."
    
    SendToUser B$, scriptdata
    
    SpecialCommandCheck = True

ElseIf p$ = "alertadmins" Then
    

    B$ = ""
    For i = 1 To NumUserParams
        B$ = B$ + UserParms(i).Value + " "
    Next i
    B$ = Trim(B$)
    
    SendToUser "Alerting admins...", scriptdata
    
    AlertAdmins B$
    
    SendToUser "Admins Alerted!", scriptdata
    
    SpecialCommandCheck = True

ElseIf p$ = "getteamnames" Then
    SendToUser "Getting Team names...", scriptdata
    GetTeamNames
    a$ = "Team 1: " + TeamNames(1) + Chr(10)
    a$ = a$ + "Team 2: " + TeamNames(2) + Chr(10)
    a$ = a$ + "Team 3: " + TeamNames(3) + Chr(10)
    a$ = a$ + "Team 4: " + TeamNames(4)
    
    SendToUser "Retrieved:" + Chr(10) + a$, scriptdata
    
    SpecialCommandCheck = True



ElseIf p$ = "scanforamafiles" Then
    SendToUser "Scanning for AMA files...", scriptdata
    
    ScanForAMAFiles
    
    SendToUser "Done!", scriptdata
    
    SpecialCommandCheck = True


End If

' Finally check for varS:

If SpecialCommandCheck = False Then
    SpecialCommandCheck = SpecialVarCheck(p$, UserParms, NumUserParams, scriptdata, NoSayMode, Scret)
End If

End Function


Function SpecialVarCheck(p$, UserParms() As typParams, NumUserParams, scriptdata As typScriptData, Optional NoSayMode As Boolean, Optional Scret As String) As Boolean

If LCase(p$) = "maxkicks" Then Var$ = Ts(General.MaxKicks)
If LCase(p$) = "maxmsg" Then Var$ = Ts(General.MaxKicks)
If LCase(p$) = "maxtime" Then Var$ = Ts(General.MaxTime)
If LCase(p$) = "maxkickvotes" Then Var$ = Ts(General.MaxKickVotes)
If LCase(p$) = "kickbantime" Then Var$ = Ts(General.BanTime)
If LCase(p$) = "VotePercent" Then Var$ = Ts(General.VotePercent)
If LCase(p$) = "point_double_kv_at" Then Var$ = Ts(PointData.DoubleKickVotesAt)
If LCase(p$) = "point_double_mv_at" Then Var$ = Ts(PointData.DoubleMapVotesAt)
If LCase(p$) = "point_kv_cost" Then Var$ = Ts(PointData.KickVotesCost)
If LCase(p$) = "point_spamkick_cost" Then Var$ = Ts(PointData.SpamKickCosts)
If LCase(p$) = "point_kicked_cost" Then Var$ = Ts(PointData.KickedCosts)
If LCase(p$) = "point_played_min" Then Var$ = Ts(PointData.JoiningAdds)
If LCase(p$) = "autoaddrealdays" Then Var$ = Ts(Val(General.AutoAddRealDays))
If LCase(p$) = "autoaddrealtimes" Then Var$ = Ts(Val(General.AutoAddRealTimes))
If LCase(p$) = "samespamnum" Then Var$ = Ts(Val(General.SameSpamNum))
If LCase(p$) = "samespamtime" Then Var$ = Ts(Val(General.SameSpamTime))

If Var$ <> "" Then  'Return the value.
    If NumUserParams = 0 Then
        SpecialVarCheck = True
        Scret = Var$
        If scriptdata.UserName <> "<SERVER>" Then
        
            SendToUser Chr(34) + LCase(p$) + Chr(34) + " is " + Chr(34) + Var$ + Chr(34), scriptdata
        
        End If
    Else
        ' is SETTING the value.
        Num = Val(UserParms(1).Value)
        
        If LCase(p$) = "maxkicks" Then General.MaxKicks = Num
        If LCase(p$) = "maxmsg" Then General.MaxKicks = Num
        If LCase(p$) = "maxtime" Then General.MaxTime = Num
        If LCase(p$) = "maxkickvotes" Then General.MaxKickVotes = Num
        If LCase(p$) = "kickbantime" Then General.BanTime = Num
        If LCase(p$) = "VotePercent" Then General.VotePercent = Num
        If LCase(p$) = "point_double_kv_at" Then PointData.DoubleKickVotesAt = Num
        If LCase(p$) = "point_double_mv_at" Then PointData.DoubleMapVotesAt = Num
        If LCase(p$) = "point_kv_cost" Then PointData.KickVotesCost = Num
        If LCase(p$) = "point_spamkick_cost" Then PointData.SpamKickCosts = Num
        If LCase(p$) = "point_kicked_cost" Then PointData.KickedCosts = Num
        If LCase(p$) = "point_played_min" Then PointData.JoiningAdds = Num
        If LCase(p$) = "autoaddrealdays" Then General.AutoAddRealDays = Ts(Num)
        If LCase(p$) = "autoaddrealtimes" Then General.AutoAddRealTimes = Ts(Num)
        If LCase(p$) = "samespamnum" Then General.SameSpamNum = Ts(Num)
        If LCase(p$) = "samespamtime" Then General.SameSpamTime = Ts(Num)
        SpecialVarCheck = True
        
        If scriptdata.UserName <> "<SERVER>" Then
        
            SendToUser Chr(34) + LCase(p$) + Chr(34) + " changed to " + Chr(34) + Var$ + Chr(34), scriptdata
        
        End If
        
    End If
End If

End Function

Function TeleportPlayer(UserID As Integer, TeleName As String) As Boolean

If DLLEnabled = False Then Exit Function
If Vars.ClanBattle = True Then Exit Function


Num = FindTele(TeleName)
num2 = FindPlayer(Ts(UserID))

TeleportPlayer = True
If Num = 0 Then TeleportPlayer = False: Exit Function


SendRCONCommand "sa_teleport_usernum " + Ts(UserID)
SendRCONCommand "sa_teleport_x " + Ts(Tele(Num).X)
SendRCONCommand "sa_teleport_y " + Ts(Tele(Num).Y)
SendRCONCommand "sa_teleport_z " + Ts(Tele(Num).Z)

SendActualRcon SA_CHECK

'set pos
If num2 > 0 Then
    Players(num2).Pos.X = Tele(Num).X
    Players(num2).Pos.Y = Tele(Num).Y
    Players(num2).Pos.Z = Tele(Num).Z
End If

End Function

Function TeleportPlayerTo(UserID As Integer, X As Integer, Y As Integer, Z As Integer) As Boolean

If DLLEnabled = False Then Exit Function
If Vars.ClanBattle = True Then Exit Function

num2 = FindPlayer(Ts(UserID))

SendRCONCommand "sa_teleport_usernum " + Ts(UserID)
SendRCONCommand "sa_teleport_x " + Ts(X)
SendRCONCommand "sa_teleport_y " + Ts(Y)
SendRCONCommand "sa_teleport_z " + Ts(Z)
SendActualRcon SA_CHECK

'set pos
If num2 > 0 Then
    Players(num2).Pos.X = X
    Players(num2).Pos.Y = Y
    Players(num2).Pos.Z = Z
End If


End Function

Function PlayerDeVoice(UserID As Integer) As Boolean
If DLLEnabled = False Then Exit Function

num2 = FindPlayer(Ts(UserID))

SendRCONCommand "sa_shutupon " + Ts(UserID)
SendActualRcon SA_CHECK

'set pos
If num2 > 0 Then
    Players(num2).ShutUp = True
    SendToWatchers "SERVER", "", "*-*-* Player  " + Players(num2).Name + " was devoiced.", 0, 0
    'SendRCONCommand "say " + General.AutoAdminName + " Player  " + Players(num2).Name + " was devoiced."
End If




End Function

Function PlayerReVoice(UserID As Integer) As Boolean

If DLLEnabled = False Then Exit Function

num2 = FindPlayer(Ts(UserID))

SendRCONCommand "sa_shutupoff " + Ts(UserID)
SendActualRcon SA_CHECK

'set pos
If num2 > 0 Then
    Players(num2).ShutUp = False
    
    SendToWatchers "SERVER", "", "*-*-* Player  " + Players(num2).Name + " was revoiced.", 0, 0
'
'    SendRCONCommand "say " + General.AutoAdminName + " Player  " + Players(num2).Name + " was revoiced."

End If

End Function

Function PlayerVoiceToggle(UserID As Integer) As Boolean

If DLLEnabled = False Then Exit Function

num2 = FindPlayer(Ts(UserID))

'set pos
If num2 > 0 Then
    If Players(num2).ShutUp = True Then
        PlayerReVoice UserID
    Else
        PlayerDeVoice UserID
    End If
End If

End Function

Function FindTele(Name As String) As Integer

For i = 1 To NumTele
    If LCase(Tele(i).Name) = LCase(Name) Then j = i: Exit For
Next i

If j > 0 Then FindTele = j

End Function

Function OLDExecuteScript(p$) As Boolean
'This sub checks for and executes the command in Vars.Command and stuff like that.

'First see if the command is one that is here.

'If SpecialCommandCheck(p$) Then OLDExecuteScript = True: Exit Function
'
'
'For I = 1 To NumCommands
'    If UCase(p$) = UCase(Commands(I).Name) Then j = I: Exit For
'Next I
'
'If j = 0 Then
'    AddToLog "    No command " + Chr(34) + p$ + Chr(34) + " found." + vbCrLf
'    ExecuteScript = False
'    Exit Function
'End If
'
'ExecuteScript = True
'
'
'
'If NumParams <> Commands(j).NumParams And Commands(j).MustHave = 1 Then
'    'MsgBox "Error! Wrong number of parameters!"
'    AddToLog "    Incorrect number of parameters!" + vbCrLf
'    SendToUser "Incorrect # of parameters for command " + p$ + "!" + vbCrLf + "Expecting " + Ts(Commands(j).NumParams) + ", got " + Ts(NumParams) + "!"
'    Exit Function
'End If
'
'fle$ = Chr(13) + Chr(10) + Commands(j).Exec + vbCrLf + vbCrLf + vbCrLf
'
'AddToLog "    Executing Script File " + Commands(j).Name + vbCrLf
'
'y = 1
'o = 1
'Do
'    'Extract a line of text
'    o = y
'    If y <> Len(fle$) + 1 Then y = InStr(y + 1, fle$, vbCrLf)
'
'    'Handle the last line events, etc
'    If y = Len(fle$) + 1 Then
'        y = 0
'    ElseIf y = 0 Then
'        y = Len(fle$) + 1
'    End If
'
'
'    If y > 0 And y > (o + 2) Then
'
'        'Get the command
'        cmd$ = Mid(fle$, o + 2, y - o)
'
'        cmd$ = Trim(cmd$)
'
'        If Left(cmd$, 2) <> "//" Then
'
'            If Right(cmd$, 2) = Chr(13) + Chr(10) Then cmd$ = Left(cmd$, Len(cmd$) - 2)
'
'            cmd$ = Trim(cmd$)
'            'MsgBox ">" + cmd$ + "<"
'
'            ExecuteCommand cmd$
'
'        End If
'
'    End If
'
'Loop Until y = 0
'

End Function

Function ExecuteScript(p$, UserParms() As typParams, NumUserParams As Integer, Scret As String, scriptdata As typScriptData, Optional NoSayMode As Boolean) As Boolean
If DebugMode Then LastCalled = "ExecuteScript"
If scriptdata.ExitNow = True Then Exit Function

'This sub checks for and executes the command in Vars.Command and stuff like that.
'First see if the command is one that is here.

'The return
Scret = ""
If p$ <> "serversay" Then
    If SpecialCommandCheck(p$, UserParms, NumUserParams, scriptdata, NoSayMode, ret$) Then
        ExecuteScript = True
        Scret = ret$
        Exit Function
    End If
End If
        
If ScriptCheck(scriptdata) Then Exit Function

DoEvents

For i = 1 To NumCommands
    If UCase(p$) = UCase(Commands(i).Name) Then j = i: Exit For
Next i

If j = 0 Then 'no such script found
    ExecuteScript = False
    Exit Function
End If

ExecuteScript = True

If NumUserParams <> Commands(j).NumParams And Commands(j).MustHave = 1 Then
    'MsgBox "Error! Wrong number of parameters!"
    SendToUser "Incorrect # of parameters for command " + p$ + "!" + vbCrLf + "Expecting " + Ts(Commands(j).NumParams) + ", got " + Ts(NumUserParams) + "!", scriptdata
    Exit Function
End If

'make a new scriptdata variable, and copy over the needed stuff.
Dim NewScriptData As typScriptData
CopyScriptData scriptdata, NewScriptData
NewScriptData.StartedName = "ExecuteScript: " + Commands(j).Exec

ret$ = ExecTextScript(Commands(j).Exec, UserParms, NumUserParams, NewScriptData)
Scret = ret$

If NewScriptData.LogToFile = True Then
    
    h = FreeFile
Close h
    Open App.Path + "\scriptlog.log" For Append As h
        Print #h, Date$ + " " + Time$ + " : " + p$ + ", " + Ts(NumUserParams) + ", took: " + Ts(Round(Timer - scriptdata.TimeStarted, 4))
    Close h

End If

End Function


Function ScriptCheck(scriptdata As typScriptData) As Boolean

' sees if it's time to end this script
If StopAllScripts Then ScriptCheck = False: Exit Function
If scriptdata.ExitNow Then ScriptCheck = False: Exit Function


' see if the script has been running longer than it's supposed to

If Int(Timer - scriptdata.TimeStarted) > 60 Then
    
    AddToLogFile "!!! ERROR !!! Script running longer than 60 seconds! Name: " & scriptdata.StartedName & ", Time Started: " & scriptdata.TimeStarted & ", Stared by: " & scriptdata.UserName & ", NumVars: " & scriptdata.NumVars
    
    ScriptCheck = False: Exit Function
End If

If FindRunningScripts Then
    
    If FindRunningScriptsTime > scriptdata.LastRunningCheck Then
        
        RunningScripts = RunningScripts & "Name: " & scriptdata.StartedName & ", Time Started: " & scriptdata.TimeStarted & ", Stared by: " & scriptdata.UserName & ", NumVars: " & scriptdata.NumVars & vbCrLf
        scriptdata.LastRunningCheck = FindRunningScriptsTime
        
    End If
End If




End Function

Function CopyScriptData(OldScriptData As typScriptData, NewScriptData As typScriptData, Optional mde As Boolean)

    NewScriptData.Index = OldScriptData.Index
    If mde Then NewScriptData.LastIf = OldScriptData.LastIf
    NewScriptData.TimeStarted = OldScriptData.TimeStarted
    NewScriptData.UserIP = OldScriptData.UserIP
    NewScriptData.UserIsTCP = OldScriptData.UserIsTCP
    NewScriptData.UserName = OldScriptData.UserName
    NewScriptData.UserPort = OldScriptData.UserPort
    NewScriptData.IsRCON = OldScriptData.IsRCON
      
    
    If mde Then
        NewScriptData.NumVars = OldScriptData.NumVars
        ReDim NewScriptData.VarNames(0 To OldScriptData.NumVars)
        ReDim NewScriptData.VarValues(0 To OldScriptData.NumVars)
        
        For i = 1 To OldScriptData.NumVars
            NewScriptData.VarNames(i) = OldScriptData.VarNames(i)
            NewScriptData.VarValues(i) = OldScriptData.VarValues(i)
        Next i
    End If
    
End Function

Function ExecTextScript(fle22$, UserParms() As typParams, NumUserParams As Integer, scriptdata As typScriptData) As String
If DebugMode Then LastCalled = "ExecTextScript"

Fle$ = fle22$
Fle$ = Fle$ + vbCrLf
Fl$ = Fle$

'trash any comments
e = 0
'do  'replaced DO with FOR
For jkk = 1 To 10000000

    e = InStrQuote(e + 1, Fl$, "//")
    If e > 0 Then
        f = InStrQuote(e + 1, Fl$, vbCrLf)
        'extract this
        If f > e Then
                
            fl1$ = ""
            If e > 1 Then fl1$ = Left(Fl$, e - 1)
            fl2$ = ""
            If e < Len(Fl$) Then fl2$ = Right(Fl$, Len(Fl$) - f - 1)
                       
            Fl$ = fl1$ + fl2$
               
        End If
    End If
    If e = 0 Then Exit For
Next jkk
If jkk = 10000000 Then AddToLogFile "!!! ERROR !!! INFINITE LOOP: Occured at " + LastCalled

'Loop Until e = 0

Fle$ = Fl$

'trash the ENTER PRESSES -- easier to exec
Fle$ = Replace(Fle$, vbCrLf, "")

'AddToLog "    Executing Script File " + Commands(j).Name + vbCrLf


Dim Brck As Integer

e = 0
'Do
'do  'replaced DO with FOR
For jkk = 1 To 10000000

    'lets look for a line
    If Brck = 0 Then f = e
chck:
    e1 = InStr(e + 1, Fle$, ";")
    e2 = InStr(e + 1, Fle$, "{")
    e3 = InStr(e + 1, Fle$, Chr(34))
    e4 = InStr(e + 1, Fle$, "}")
    
    flg = 0
    e = e1
    If e = 0 Then e = 100000000
    If e2 < e And e2 > 0 Then e = e2: flg = 1
    If e4 < e And e4 > 0 Then e = e4: flg = 2
    
    If e3 < e And e3 <> 0 Then
        e = InStr(e3 + 1, Fle$, Chr(34))
        jkkk = jkkk + 1
        If jkkk > 1000000 Then Exit For
        
        GoTo chck
    End If
       
    If flg = 1 Then Brck = Brck + 1
    If flg = 2 Then Brck = Brck - 1
    
    If Brck = 0 And e <> 100000000 Then
    
        'all set, e contains pos of the ";" or the "{" whichever comes first
        
        cmd$ = Mid(Fle$, f + 1, e - f)
        ret$ = ""
        ret$ = ExecuteCommand(cmd$, UserParms, NumUserParams, scriptdata)
        
        'if something was returned, exit immediatly
        If ret$ <> "" Then Exit For
    End If
    
    If ScriptCheck(scriptdata) Then Exit Function
    If scriptdata.ExitNow = True Then Exit Function
    DoEvents

'Loop Until e = 100000000
    If e = 100000000 Then Exit For
Next jkk
If jkk = 10000000 Then AddToLogFile "!!! ERROR !!! INFINITE LOOP: Occured at ExecTextScript!!"

ExecTextScript = ret$

End Function


Function ExecuteCommand(cmd$, UserParms() As typParams, NumUserParams As Integer, scriptdata As typScriptData) As String
If DebugMode Then LastCalled = "ExecuteCommand"

'Executes a script command.

cmd$ = Trim(cmd$)

If cmd$ = "" Then Exit Function

'get rid of the semicolon at the end
If Right(cmd$, 1) = ";" Then cmd$ = Left(cmd$, Len(cmd$) - 1)

'Samples:
'define map1;
'OR
'if (abc = jhi) {eatme;}
'OR
'if (abc = "muu") kill;

'First, extract the actual command.
e = InStrQuote(1, cmd$, " ")
                                        
If e = 0 Then 'No paramerers.
    'example:
    'startit;
    
    ret$ = ExecFinal(cmd$, UserParms, NumUserParams, scriptdata)
Else 'There are parameters.
        
    'how it MIGHT look: (one possibility per line)
    '1- startit(abc, def, ghi);                -- THIS is a command to start a script
    '2- abc = startit(jef);                    -- THIS is a command to start a script and put the result into the var "abc"
    '3- if (jef) {blah; blah;}                 -- IF block with more than one command
    '4- if (jef) startit(123);                 -- IF block with only one command
       
    'NOW we must RECOGNIZE the TYPE of command we have here (see above)
    
    'see if its type 2
    er$ = Replace(cmd$, " ", "")
    e1 = InStr(1, er$, "=")
    e2 = InStr(1, er$, Chr(34))
    If e2 = 0 Then e2 = 1000000000
    
    If e1 = e And e1 < e2 Then 'yes it is type 2
    
        'get the variable we are supposed to be filling
        Var$ = Left(cmd$, e - 1)
        'we got something like this:
        'abc = startit(jef);
        'so now lets try to exec the command
        
        'using the new thing:
        f = InStr(1, cmd$, "=")
        fr$ = Trim(Right(cmd$, Len(cmd$) - f - 1))
                
        ret$ = ExecFormula(fr$, UserParms, NumUserParams, scriptdata)
        
        ret$ = TrimQuotes(ret$)
        
        
        'stick that into the var :P
        
        
        SetVar Var$, ret$, scriptdata
    
    ElseIf LeftR(LCase(cmd$), 2) = "if" Then    'the IF situation
                   
        r$ = DoIf(cmd$, UserParms, NumUserParams, scriptdata)
        
    ElseIf LeftR(LCase(cmd$), 6) = "elseif" Then     'the ELSEIF situation, only if the last if or elseif attempt failed
                   
        If scriptdata.LastIf <> 0 Then r$ = DoIf(cmd$, UserParms, NumUserParams, scriptdata)
        
    ElseIf LeftR(LCase(cmd$), 4) = "else" Then    'the ELSE situation, only if the last if or elseif attempt failed
                       
        If scriptdata.LastIf <> 0 Then
            'this isnt handled by DoIf, because it can only have two possibilities:
            'else stop();
            'or
            'else {   blah blah    }
            
            'extract the command contained in the var... and execute it! hahaha!
            commd$ = Trim(Right(cmd$, Len(cmd$) - 4))
             
            'see if its with or without { }
            If Left(commd$, 1) = "{" Then 'trash the first and last...
                commd$ = Mid(commd$, 2, Len(commd$) - 2)
                
                mylastif = scriptdata.LastIf
                r$ = ExecTextScript(commd$, UserParms, NumUserParams, scriptdata)
                scriptdata.LastIf = mylastif

            Else
                'just a simple command
                'LOOK!!! RECURSION!! COOL, EH?
                nne$ = ExecuteCommand(commd$, UserParms, NumUserParams, scriptdata)
            End If
        End If
        scriptdata.LastIf = 0
    ElseIf LeftR(LCase(cmd$), 6) = "global" Or LeftR(LCase(cmd$), 6) = "define" Then     'creating a new variable... this is a special case and can't be handled by the other subs :P
        'format:
        'define varname;
        Varname$ = LCase(Trim(Right(cmd$, Len(cmd$) - 6)))
        
        'check if it exists...
        
        For i = 1 To NumUserVars
            If LCase(UserVars(i).Name) = Varname$ Then j = i: Exit For
        Next i
        
        If j = 0 Then
            'doesnt exist... create it
            NumUserVars = NumUserVars + 1
            j = NumUserVars
            
            ReDim Preserve UserVars(0 To j)
            UserVars(j).Name = Varname$
            UserVars(j).Value = ""
        End If
    
    ElseIf LeftR(LCase(cmd$), 3) = "dim" Then     'creating a new LOCAL variable...
        'format:
        'define varname;
        Varname$ = LCase(Trim(Right(cmd$, Len(cmd$) - 3)))
        
        'check if it exists globally...
        
        For i = 1 To NumUserVars
            If LCase(UserVars(i).Name) = Varname$ Then j = i: Exit For
        Next i
        
        If j = 0 Then
        
            For i = 1 To scriptdata.NumVars
                If LCase(scriptdata.VarNames(i)) = Varname$ Then j = i: Exit For
            Next i
        
            If j = 0 Then
        
                'doesnt exist... create it
                scriptdata.NumVars = scriptdata.NumVars + 1
                j = scriptdata.NumVars
                
                ReDim Preserve scriptdata.VarNames(0 To j)
                ReDim Preserve scriptdata.VarValues(0 To j)
                scriptdata.VarNames(j) = Varname$
            End If
        End If
    
    ElseIf LeftR(LCase(cmd$), 7) = "destroy" Then     'destroying a variable... this removes it from memory
                   
        'format:
        'destroy varname;
        
        Varname$ = LCase(Trim(Right(cmd$, Len(cmd$) - 7)))
        
        'check if it exists...
        
        For i = 1 To NumUserVars
            If LCase(UserVars(i).Name) = Varname$ Then j = i: Exit For
        Next i
        
        If j <> 0 Then
            'exists... now delete it
            For i = j To NumUserVars - 1
                UserVars(i).Name = UserVars(i + 1).Name
                UserVars(i).Value = UserVars(i + 1).Value
            Next i
       
            NumUserVars = NumUserVars - 1
            ReDim Preserve UserVars(0 To NumUserVars)
        End If
    Else
        
        'no, its type 1
        ret$ = ExecFinal(cmd$, UserParms, NumUserParams, scriptdata)
    End If
End If

AddToLog "        Executing Command " + Vars.ScriptCommand + vbCrLf
'ExecuteScriptCommand Vars.ScriptCommand

'returning some info... get the info and QUIT!
'unless the command was return, clear what was returned :)

If LeftR(LCase(cmd$), 6) <> "return" Then ret$ = ""
ExecuteCommand = ret$

End Function

Function LeftR(String1 As String, length) As String
'just a non-bug left command
If length <= Len(String1) Then LeftR = Left(String1, length)

End Function

Function RightR(String1 As String, length) As String
'just a non-bug left command
If length <= Len(String1) Then RightR = Right(String1, length)

End Function



Function DoIf(cmd$, UserParms() As typParams, NumUserParams As Integer, scriptdata As typScriptData) As String
If DebugMode Then LastCalled = "DoIf"

'Proccesses IF, ELSEIF, and ELSE

'LASTIF vars
'0 - there was no last if, or it was TRUE
'1 - the last if was FALSE
      
'format of an if:
'if (abc = def) or (ghi != jkl) and (bbb > ccc) {commands;}
'OR
'if (abc = def) or (ghi != jkl) command;

'things that go in the middle:
'=      equal to
'!=     not equal to
'>      greater than
'<      less than
'>= =>  greater than or equal to
'<= =<  less than or equal to

'statements:
'xx OR xx           either or both must be true
'xx AND xx          both must be true

'order of extraction:
'vars - check for trueness
'final trueness statement - assemble
'if true, extract commands and execute
'if false, set lastif and exit

'part 1 - extract and check vars
Dim Brck As Integer




d$ = cmd$

StartAgain:
G = InStrQuote(1, d$, "{")
e = 0
fff = 0
Brck = 0

'do  'replaced DO with FOR
For jkk = 1 To 10000000

    'lets look for a bracket
    If Brck = 1 And fff = 0 Then f = e: fff = 1
chck:
    e2 = InStr(e + 1, d$, "(")
    e3 = InStr(e + 1, d$, Chr(34))
    e4 = InStr(e + 1, d$, ")")
    
    flg = 1
    e = e2
    If e = 0 Then e = 100000000
    If e4 < e And e4 > 0 Then e = e4: flg = 2
    
    If e > G And G <> 0 Then Exit For
    
    If e3 < e And e3 <> 0 Then
        e = InStr(e3 + 1, d$, Chr(34))
        GoTo chck
    End If
       
       
    If flg = 1 Then Brck = Brck + 1
    If flg = 2 Then Brck = Brck - 1
    
    If Brck = 0 And e <> 100000000 Then
            
        cm$ = Mid(d$, f + 1, e - f - 1)
        fff = 0
        If Mid(d$, f - 1, 1) <> " " Then GoTo nextone
        
        'got the inside of the brackets, split it up between the equal signs
                   
        'pass to the centreif sub which will do this for us
        rr = DoCentreIf(cm$, UserParms, NumUserParams, scriptdata)
        
        'now place this where the old one was
        
        If rr = True Then retrn$ = "1"
        If rr = False Then retrn$ = "0"
        
        d1$ = ""
        If f > 1 Then d1$ = Left(d$, f - 1)
        d2$ = ""
        If e < Len(d$) Then d2$ = Right(d$, Len(d$) - e)
        d$ = d1$ + retrn$ + d2$
        lastlen = Len(d1$ + retrn$)
        GoTo StartAgain
        
'        Debug.Print cmd$
        
    End If
nextone:

    If e = 100000000 Then Exit For
Next jkk
If jkk = 10000000 Then AddToLogFile "!!! ERROR !!! INFINITE LOOP: Occured at " + LastCalled

'Loop Until e = 100000000

'now we have a statement that looks like this:

'if 1 and 0 and 1 and 1 {.....
'or
'if 1 and 0 and 1 and 1 startit;

'extract the important part...
e = InStr(1, d$, " ")
checkstr$ = LCase(Replace(Mid(d$, e + 1, lastlen - e), " ", ""))

'looks like:
'1 and 0 or 1 and 1

'start extracting
e = 0
f = 0
reloop:

'do  'replaced DO with FOR
For jkk = 1 To 10000000
    'get the first number
    e = 0
    e1 = InStr(e + 1, checkstr$, "1")
    e2 = InStr(e + 1, checkstr$, "0")
    
    e = e1
    If e = 0 Then e = 100000000
    If e2 <= e And e2 <> 0 Then e = e2
    
    
    'get the second number
    G = 0
    G1 = InStr(e + 1, checkstr$, "1")
    G2 = InStr(e + 1, checkstr$, "0")
    
    G = G1
    If G = 0 Then G = 100000000
    If G2 <= G And G2 <> 0 Then G = G2
    
      
    If e > 0 And G > 0 And G <> 100000000 And e <> 100000000 Then
    
        chck$ = Mid(checkstr$, e + 1, G - e - 1)
        'terms:
        t1 = Val(Mid(checkstr$, e, 1))
        t2 = Val(Mid(checkstr$, G, 1))
        
        Ans = 0
        If chck$ = "and" And t1 = 1 And t2 = 1 Then Ans = 1
        If chck$ = "or" And t1 = 0 And t2 = 1 Then Ans = 1
        If chck$ = "or" And t1 = 1 And t2 = 0 Then Ans = 1
        If chck$ = "or" And t1 = 1 And t2 = 1 Then Ans = 1
        If chck$ = "xor" And t1 = 0 And t2 = 1 Then Ans = 1
        If chck$ = "xor" And t1 = 1 And t2 = 0 Then Ans = 1
        
        'now that we have the "ans" stick it back in
        d1$ = ""
        If e > 1 Then d1$ = Left(checkstr$, e - 1)
        d2$ = ""
        If G < Len(checkstr$) Then d2$ = Right(checkstr$, Len(checkstr$) - G)
        checkstr$ = d1$ + Ts(Ans) + d2$
        GoTo reloop
    End If

'Loop Until e = 100000000 Or G = 100000000

    If e = 100000000 Or G = 100000000 Then Exit For
Next jkk
If jkk = 10000000 Then AddToLogFile "!!! ERROR !!! INFINITE LOOP: Occured at " + LastCalled


'FINALLY we have just ONE statement that says either "1" or "0"


If checkstr$ = "1" Then
    'TRUE!!
       
    'extract the command contained in the var... and execute it! hahaha!
    commd$ = Trim(Right(d$, Len(d$) - lastlen))
     
    'see if its with or without { }
    If Left(commd$, 1) = "{" Then 'trash the first and last...
        commd$ = Mid(commd$, 2, Len(commd$) - 2)
        
        mylastif = scriptdata.LastIf
        ret$ = ExecTextScript(commd$, UserParms, NumUserParams, scriptdata)
        scriptdata.LastIf = mylastif
        
    Else
        'just a simple command
        nne$ = ExecuteCommand(commd$, UserParms, NumUserParams, scriptdata)
    End If
    
    scriptdata.LastIf = 0
Else
    'FALSE!!
    'Set the lastif var
    scriptdata.LastIf = 1
    
    'And thats it... its up to the next sub to handle the rest.
End If

End Function

Function DoCentreIf(cm$, UserParms() As typParams, NumUserParams As Integer, scriptdata As typScriptData) As Boolean
If DebugMode Then LastCalled = "DoCentreIf"

'just takes
'abc = def
'and returns TRUE or FALSE

'things that go in the middle:      type
'=      equal to                    1
'!=     not equal to                2
'>      greater than                3
'<      less than                   4
'>= =>  greater than or equal to    5
'<= =<  less than or equal to       6

'first get the centre divider
e1 = InStrQuote(1, cm$, "=")
e2 = InStrQuote(1, cm$, "!=")
e3 = InStrQuote(1, cm$, ">")
e4 = InStrQuote(1, cm$, "<")
e5 = InStrQuote(1, cm$, "<=")
e6 = InStrQuote(1, cm$, "=>")
e7 = InStrQuote(1, cm$, "=<")
e8 = InStrQuote(1, cm$, ">=")

'find out what it is
If e1 Then Typ = 1: e = e1: l = 1
If e2 Then Typ = 2: e = e2: l = 2
If e3 Then Typ = 3: e = e3: l = 1
If e4 Then Typ = 4: e = e4: l = 1
If e5 Then Typ = 6: e = e5: l = 2
If e6 Then Typ = 5: e = e6: l = 2
If e7 Then Typ = 6: e = e7: l = 2
If e8 Then Typ = 5: e = e8: l = 2

'split it up
c1$ = Trim(Left(cm$, e - 1))
c2$ = Trim(Right(cm$, Len(cm$) - e - 1))

'now evaluate the expressions
c1$ = TrimQuotes(ExecFormula(c1$, UserParms, NumUserParams, scriptdata))
c2$ = TrimQuotes(ExecFormula(c2$, UserParms, NumUserParams, scriptdata))

'now see if its true

Dim Ans As Boolean
Ans = False

If Typ = 1 And c1$ = c2$ Then Ans = True
If Typ = 2 And c1$ <> c2$ Then Ans = True
If Typ = 3 And Val(c1$) > Val(c2$) Then Ans = True
If Typ = 4 And Val(c1$) < Val(c2$) Then Ans = True
If Typ = 5 And (Val(c1$) > Val(c2$) Or Val(c1$) = Val(c2$)) Then Ans = True
If Typ = 6 And (Val(c1$) < Val(c2$) Or Val(c1$) = Val(c2$)) Then Ans = True

DoCentreIf = Ans

End Function

Function TrimQuotes(String1 As String) As String
If DebugMode Then LastCalled = "TrimQuotes"

'just trashes the quotes on the outside

a$ = String1
If Len(String1) > 1 Then
    Do
        dne = 0
        If Left(a$, 1) = Chr(34) Then a$ = Right(a$, Len(a$) - 1): dne = 1
        If Right(a$, 1) = Chr(34) And Len(a$) > 1 Then a$ = Left(a$, Len(a$) - 1): dne = 1
        If a$ = Chr(34) Then a$ = "": dne = 1
    Loop Until dne = 0 Or a$ = ""
End If

TrimQuotes = a$

End Function

Function InStrQuote(Start, String1 As String, String2 As String) As Integer
If DebugMode Then LastCalled = "InStrQuote"

'works JUST like InStr(), except this one only returns the requested character if it ISNT in a quote

e = Start - 1
'do  'replaced DO with FOR
For jkk = 1 To 10000000

    f = e
agn:
    e1 = InStr(e + 1, String1, String2)
    e2 = InStr(e + 1, String1, Chr(34))
    e = e1
    If e2 < e1 And e2 <> 0 Then
        e = InStr(e2 + 1, String1, Chr(34))
        jkk = jkk + 1
        If jkk > 1000000 Then AddToLogFile "!!! ERROR !!! INFINITE LOOP: Occured at " + LastCalled: scriptdata.ExitNow = True: Exit Function
        GoTo agn
    End If
                
    'got it
    If e > 0 And e > f Then
        InStrQuote = e
        Exit Function
    End If
    
    If StopAllScripts Then Exit Function
    DoEvents
    
'Loop Until e = 0

    If e = 0 Then Exit For
Next jkk
If jkk = 10000000 Then AddToLogFile "!!! ERROR !!! INFINITE LOOP: Occured at " + LastCalled

InStrQuote = 0

End Function

Function BracketCount(CharNum, String1 As String) As Integer
If DebugMode Then LastCalled = "BracketCount"

'Counts the bracket level of this character

Dim Brck As Integer

e = 0
fff = 0
Brck = 0

'do  'replaced DO with FOR
For jkk = 1 To 10000000

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
    
    
    If StopAllScripts Then Exit Function
    DoEvents

'Loop Until e = 100000000 Or e >= CharNum
    If e = 100000000 Or e >= CharNum Then Exit For
Next jkk
If jkk = 10000000 Then AddToLogFile "!!! ERROR !!! INFINITE LOOP: Occured at " + LastCalled


BracketCount = Brck

End Function


Function ExecFinal(cmd$, UserParms() As typParams, NumUserParams As Integer, scriptdata As typScriptData) As String
If DebugMode Then LastCalled = "ExecFinal"

If ScriptCheck(scriptdata) Then Exit Function
If scriptdata.ExitNow = True Then Exit Function

DoEvents

'params var

Dim Parms() As typParams
ReDim Parms(1 To 200)

'FINALLY, this will take:
'start(abc, def, "ghi")                 -- execs script
'OR
'abc                        -- simply returns the variables value
'OR
'
'and extract all the parameters, etc etc, and execute the script, and return the result

'FIRST check if its a number
If Ts(Val(cmd$)) = cmd$ Then
    ExecFinal = Ts(Val(cmd$))
    Exit Function
End If

If "0" + Ts(Val(cmd$)) = cmd$ Then
    ExecFinal = Ts(Val(cmd$))
    Exit Function
End If

Num = 0
e = InStrQuote(1, cmd$, "(")
If e Then 'we seem to have CASE 1
    'get the command
    
    If e > 1 Then cm$ = Left(cmd$, e - 1)
    cm$ = Trim(cm$)
    'cm$ now has "start"
    
    'now see what kind of parameters we have:
    f = InStrRev(cmd$, ")")
    
    If f > e Then
        prms$ = Mid(cmd$, e + 1, f - e - 1)
        'now prms$ is something like: abc, def, "ghi", heg + may, "hey, whats up?"
        'formats like this:           |-|  |-|  |---|  |-------|  |--------------|
        'evaluate each term:
        e = 0
        
        'do  'replaced DO with FOR
        For jkk = 1 To 10000000
        f = e
agn:
            Do
                e1 = InStr(e + 1, prms$, ",")
                br = BracketCount(e1, prms$)
                If br > 0 And e1 > 0 Then e = e1
            Loop Until e1 = 0 Or br = 0
            
            e2 = InStr(e + 1, prms$, Chr(34))
            If f = Len(prms$) + 1 Then e1 = -1
            If e1 = 0 Then e1 = Len(prms$) + 1
            
            
            e = e1
            If e2 < e1 And e2 <> 0 Then
                e = InStr(e2 + 1, prms$, Chr(34))
                GoTo agn
            End If
                        
            'finally got it
            If e > 0 And e > f Then
                pr$ = Trim(Mid(prms$, f + 1, e - f - 1))
                        
                'evaluate this parameter
                       
                ev$ = ExecFormula(pr$, UserParms, NumUserParams, scriptdata)
                Num = Num + 1
                'now we have the parameters value, stick it into the params thing
                ev$ = TrimQuotes(ev$)
                
                'it could have the code for CHR(34) in it
                
                ev$ = Replace(ev$, Chr(255), Chr(34))
                                
                Parms(Num).Value = ev$
                
            End If
            If ScriptCheck(scriptdata) Then Exit Function
            If scriptdata.ExitNow = True Then Exit Function
            DoEvents

            If e = -1 Then Exit For
        Next jkk
        If jkk = 10000000 Then AddToLogFile "!!! ERROR !!! INFINITE LOOP: Occured at " + LastCalled
            
        'Loop Until e = -1
    End If

    'that was that
Else
    'it could also be a variable...
    'let executescriptcommand handle that case :P
    
    cm$ = cmd$
End If

'now run it

ret$ = ExecuteScriptCommand(cm$, Parms, Num, UserParms, NumUserParams, scriptdata)
ExecFinal = ret$

End Function

Function GetVar(Varname As String, scriptdata As typScriptData) As String
If DebugMode Then LastCalled = "Getvar"

Varname = LCase(Varname)

'SERVER VARS:

If Varname = "hlpath" Then GetVar = Server.HLPath
'If Varname = "numparams" Then GetVar = Ts(NumParams)
If Varname = "gamedir" Then GetVar = Server.GamePath
If Varname = "rconpass" Then GetVar = Server.RCONPass
If Varname = "map" Then GetVar = Vars.Map
If Varname = "userip" Then GetVar = Vars.UserIP
If Varname = "userport" Then GetVar = Ts(Vars.UserPort)
If Varname = "username" Then GetVar = scriptdata.UserName
If Varname = "path" Then GetVar = Server.BothPath
If Varname = "numusers" Then GetVar = Ts(NumPlayers)
If Varname = "hostname" Then GetVar = Server.HostName
If Varname = "localip" Then GetVar = Server.LocalIP
If Varname = "votedmap" Then GetVar = Vars.VotedMap
If Varname = "clanbattle" Then GetVar = Ts(CInt(Vars.ClanBattle))
If Varname = "time" Then GetVar = Time$
If Varname = "timeform" Then GetVar = Format(Time$, "hh:mm:ss AMPM")
If Varname = "date" Then GetVar = Date$
If Varname = "dateform" Then GetVar = Format(Date$, "dddd, mmm dd, yyyy")
If Varname = "gameport" Then GetVar = Server.ServerPort
If Varname = "timeleft" Then GetVar = Ts(Vars.MapTimeLeft)
If Varname = "version" Then GetVar = Ts(App.Major) + "." + Ts(App.Minor) + "." + Ts(App.Revision)
If Varname = "pvversion" Then GetVar = PVVersion
If Varname = "maxplayers" Then GetVar = Ts(Vars.MaxPlayers)
If Varname = "nl" Then GetVar = Chr(10)
If Varname = "quot" Then GetVar = Chr(34)
If Varname = "useristcp" And scriptdata.UserIsTCP = True Then GetVar = "1"
If Varname = "useristcp" And scriptdata.UserIsTCP = False Then GetVar = "0"
If Varname = "pi" Then GetVar = "3.141592654"
If Varname = "e" Then GetVar = "2.718281828"


GetVar = Replace(GetVar, Chr(34), Chr(255))
If GetVar <> "" Then Exit Function

For i = 1 To NumUserVars
    If LCase(UserVars(i).Name) = Varname Then j = i: Exit For
Next i

If j = 0 Then
    
    'search local vars
    
    For i = 1 To scriptdata.NumVars
        If LCase(scriptdata.VarNames(i)) = Varname Then j = i: Exit For
    Next i

    If j = 0 Then
        GetVar = Chr(0)
        Exit Function
    Else
        GetVar = scriptdata.VarValues(j)
        GetVar = Replace(GetVar, Chr(34), Chr(255))
    End If
Else
    GetVar = UserVars(j).Value
    GetVar = Replace(GetVar, Chr(34), Chr(255))
End If


End Function

Function SetVar(Varname As String, Value As String, scriptdata As typScriptData) As Boolean
If DebugMode Then LastCalled = "Setvar"

Varname = LCase(Varname)

For i = 1 To NumUserVars
    If LCase(UserVars(i).Name) = Varname Then j = i: Exit For
Next i


SetVar = True
If j = 0 Then
    
    For i = 1 To scriptdata.NumVars
        If LCase(scriptdata.VarNames(i)) = Varname Then j = i: Exit For
    Next i
    
    If j = 0 Then
        If scriptdata.NoAutoCreate = False Then
            'doesnt exist... create it
            scriptdata.NumVars = scriptdata.NumVars + 1
            j = scriptdata.NumVars
            
            ReDim Preserve scriptdata.VarNames(0 To j)
            ReDim Preserve scriptdata.VarValues(0 To j)
            scriptdata.VarNames(j) = Varname
            scriptdata.VarValues(j) = Value
        Else
    
            SetVar = False
            Exit Function
        End If
    Else
        scriptdata.VarValues(i) = Value
    End If
Else
    UserVars(j).Value = Value
End If



End Function

Function ExecFormula(fr$, UserParms() As typParams, NumUserParams As Integer, scriptdata As typScriptData) As String
If DebugMode Then LastCalled = "ExecFormula"

If scriptdata.ExitNow = True Then Exit Function

'this is an amazing function. it will take a line like this:
'
'  abc + (start(muu) + fudge + 10) * def
'OR
'  start(stank, fudge) + deg
'
'and will replace all the vars with the proper values, and execute all the scripts, and finally
'add it all together and return the value

'first lets trash all of the spaces and leave just the important stuff

e = 0
'do  'replaced DO with FOR
For jkk = 1 To 10000000
    
    f = e
agn:
    'e1 = InStr(e + 1, fr$, " ")
    'e2 = InStr(e + 1, fr$, Chr(34))
    
    e1 = InStrQuote(e + 1, fr$, " ")
    
    If f = Len(fr$) + 1 Then e1 = -1
    If e1 = 0 Then e1 = Len(fr$) + 1
    e = e1
    'If e2 < e1 And e2 <> 0 Then
    '    e = InStr(e2 + 1, fr$, Chr(34))
    '    GoTo agn
    'End If
                
    'finally got it
    If e > 0 And e > f Then
        pr$ = Trim(Mid(fr$, f + 1, e - f - 1))
        d$ = d$ + pr$
    End If
'Loop Until e = -1

    If e = -1 Then Exit For
Next jkk
If jkk = 10000000 Then AddToLogFile "!!! ERROR !!! INFINITE LOOP: Occured at " + LastCalled


fr$ = d$

'now we have only the line without any spaces EXCEPT IN QUOTES left

'VALID seperators are:
'+      addition... both strings and numbers
'-      subraction... only numbers
'*      multiplication... only numbers
'/      division... only numbers
'^      power... only numbers

'SAMPLE line:
'abc+def*(mach4(muu,"ruu fuu")+juu)-ghi

'NOW we must separate the BRACKETS!

Dim Brck As Integer

StartAgain:
e = 0
fff = 0
Brck = 0

'do  'replaced DO with FOR
For jkk = 1 To 10000000
    'lets look for a line
    If Brck = 1 And fff = 0 Then f = e: fff = 1
chck:
    If scriptdata.ExitNow = True Then Exit Function
    e2 = InStr(e + 1, d$, "(")
    e3 = InStr(e + 1, d$, Chr(34))
    e4 = InStr(e + 1, d$, ")")
    
    flg = 1
    e = e2
    If e = 0 Then e = 100000000
    If e4 < e And e4 > 0 Then e = e4: flg = 2
    
    If e3 < e And e3 <> 0 Then
        jkkk = jkkk + 1
        e = InStr(e3 + 1, d$, Chr(34))
        
        If jkkk > 1000000 Then Exit For
        GoTo chck
    End If
       
    If flg = 1 Then Brck = Brck + 1
    If flg = 2 Then Brck = Brck - 1
    
    If Brck = 0 And e <> 100000000 Then
            
        cmd$ = Mid(d$, f + 1, e - f - 1)
        fff = 0
        If f > 1 Then
            tst$ = Mid(d$, f - 1, 1)
            If tst$ <> "&" And tst$ <> "+" And tst$ <> "-" And tst$ <> "*" And tst$ <> "/" And tst$ <> "^" Then GoTo nextone
        End If
        'got the inside of the brackets, call this function again
            
'        if instr(1, cmd$, "+") = 0 and instr(1, cmd$, "-") = 0 and
            
        retrn$ = ExecFormula(cmd$, UserParms, NumUserParams, scriptdata)
        
        'now place this where the old one was
        retrn$ = TrimQuotes(retrn$)
        
        
        d1$ = ""
        If f > 1 Then d1$ = Left(d$, f - 1)
        d2$ = ""
        If e < Len(d$) Then d2$ = Right(d$, Len(d$) - e)
        d$ = d1$ + retrn$ + d2$
        GoTo StartAgain
        
'        Debug.Print cmd$
        
    End If
nextone:

'Loop Until e = 100000000
    If e = 100000000 Then Exit For
Next jkk
If jkk = 10000000 Then AddToLogFile "!!! ERROR !!! INFINITE LOOP: Occured at " + LastCalled: scriptdata.ExitNow = True



'finally all the brackets are out of the function, now lets add the old bastard together!

ret$ = ExecFormulaFinal(d$, UserParms, NumUserParams, scriptdata)
ExecFormula = ret$


End Function

Function ExecFormulaFinal(d$, UserParms() As typParams, NumUserParams As Integer, scriptdata As typScriptData) As String
If DebugMode Then LastCalled = "ExecFormulaFinal"

If scriptdata.ExitNow = True Then Exit Function
'VALID seperators are:
'&      sticks two strings together...
'+      addition... only numbers
'-      subraction... only numbers
'*      multiplication... only numbers
'/      division... only numbers
'^      power... only numbers

'SAMPLE line:
'abc+def*mach4+juu-ghi

'there are NO MORE BRACKETS in this formula anymore!
'EXCEPT for brackets meaning the start of a sub, like start(abc, def)

'seperate it by the signs:
'just start looping thru
buff$ = ""
For i = 1 To Len(d$)
    If scriptdata.ExitNow = True Then Exit Function
    c$ = Mid(d$, i, 1)
    If c$ = Chr(34) And inquot = 0 Then
        inquot = 1
    ElseIf c$ = Chr(34) And inquot = 1 Then
        inquot = 0
    End If
    
    If c$ = "(" And inquot = 0 Then brack = brack + 1
    If c$ = ")" And inquot = 0 Then brack = brack - 1
        
    If (c$ = "+" Or c$ = "-" Or c$ = "*" Or c$ = "/" Or c$ = "&" Or c$ = "^") And inquot = 0 And brack = 0 Then
        'If c$ = "-" And I = 1 Then
        '    buff$ = buff$ + c$
        'Else
            'get the value of this and stick it in
            If Len(buff$) >= 1 Then
                If Left(buff$, 1) = Chr(34) And Right(buff$, 1) = Chr(34) Then
                    'here we have something that is in quotes
                    'we just want to replace \n with enter
                    buff$ = Replace(buff$, " \n", vbCrLf)
                    
                    a$ = a$ + buff$ + c$
                    buff$ = ""
                Else
            
                    f$ = ExecFinal(buff$, UserParms, NumUserParams, scriptdata)
                    buff$ = ""
                    a$ = a$ + Chr(34) + f$ + Chr(34) + c$
                End If
            Else
                buff$ = buff$ + c$
            End If
        'Else
        
        'End If
    Else
        buff$ = buff$ + c$
    End If
Next i

'final go
If Len(buff$) >= 1 Then
    If Left(buff$, 1) = Chr(34) And Right(buff$, 1) = Chr(34) Then
        
        a$ = a$ + buff$
        buff$ = ""
    Else

        f$ = ExecFinal(buff$, UserParms, NumUserParams, scriptdata)
        buff$ = ""
        a$ = a$ + Chr(34) + f$ + Chr(34)
    End If
End If

'FINALLY we have something thats ready for evaluation:
'"12"+"19"+"49"-"78"
'or "abc"&"def"

'do the math in order...
a$ = DoTerm("^", a$)
a$ = DoTerm("*", a$)
'a$ = DoTerm("/", a$)
a$ = DoTerm("+", a$)
a$ = DoTerm("-", a$)
a$ = DoTerm("&", a$)

ExecFormulaFinal = Chr(34) + a$ + Chr(34)

End Function

Function DoTerm(trm$, a$) As String
If DebugMode Then LastCalled = "DoTerm"


'DO a term:
e = 0
'do  'replaced DO with FOR
For jkk = 1 To 10000000
    'find the term
    If e = -1 Then e = 0
    olde = e
    e = InStrQuote(e + 1, a$, trm$)
    
    currtrm$ = trm$
    If trm$ = "*" Then
        e2 = InStrQuote(olde + 1, a$, "/")
        If e2 > 0 And (e2 < e Or e = 0) Then e = e2: currtrm$ = "/"
        
    End If
'
'    If trm$ = "+" Then
'        e2 = InStrQuote(olde + 1, a$, "-")
'        If e2 > 0 And e2 < e Then e = e2: currtrm$ = "-"
'    End If
    
    If e > 0 Then
        'find the term thats ONE before...
        f1 = InStrRev(a$, Chr(34), e - 2)
        'and the one thats right after
        G1 = InStr(e + 2, a$, Chr(34))
        
        If f1 = 0 Then f1 = 1
        If G1 = 0 Then G1 = Len(a$)
        
        If f1 < e And G1 > e Then
            'get the first and second term
            t1$ = Mid(a$, f1, e - f1)
            t2$ = Mid(a$, e + 1, G1 - e)
            
            'strip the quotes
            t1$ = TrimQuotes(t1$)
            t2$ = TrimQuotes(t2$)
            
            'do the math
            If currtrm$ = "*" Then Ans$ = Ts(Val(t1$) * Val(t2$))
            If currtrm$ = "/" And Val(t2$) <> 0 Then Ans$ = Ts(Val(t1$) / Val(t2$))
            If currtrm$ = "-" Then Ans$ = Ts(Val(t1$) - Val(t2$))
            If currtrm$ = "+" Then Ans$ = Ts(Val(t1$) + Val(t2$))
            If currtrm$ = "^" Then Ans$ = Ts(Val(t1$) ^ Val(t2$))
            If currtrm$ = "&" Then Ans$ = t1$ + t2$
            
            'finally replace it all out
            a1$ = ""
            If f1 > 1 Then a1$ = Left(a$, f1 - 1)
            a2$ = ""
            If G1 < Len(a$) Then a2$ = Right(a$, Len(a$) - G1)
            
            a$ = a1$ + Chr(34) + Ans$ + Chr(34) + a2$
            e = -1
        End If
    End If

'Loop Until e = 0

    If e = 0 Then Exit For
Next jkk
If jkk = 10000000 Then AddToLogFile "!!! ERROR !!! INFINITE LOOP: Occured at " + LastCalled


DoTerm = a$


End Function

Function ExecuteScriptCommand(c$, Parms() As typParams, NumParms, UserParms() As typParams, NumUserParams As Integer, scriptdata As typScriptData) As String
If DebugMode Then LastCalled = "ExecuteScriptCommand"
If scriptdata.ExitNow = True Then Exit Function
If ScriptCheck(scriptdata) Then Exit Function
DoEvents

On Error GoTo errocc

'This sub will RUN the script file with all it's parameters.

'FIRST Check if the given par is a VAR
a$ = GetVar(c$, scriptdata)
If a$ <> Chr(0) Then 'so it is
    ExecuteScriptCommand = a$
    Exit Function
End If

'Now check if its a SCRIPT

Dim NewUserParms() As typParams
ReDim NewUserParms(1 To 200)

For i = 1 To NumParms
    NewUserParms(i).Value = Parms(i).Value
Next i

'try to execute
ae = ExecuteScript(c$, NewUserParms, CInt(NumParms), Scret$, scriptdata)
If ae Then ret$ = Scret$

'finally check if its a built-in command
c$ = LCase(c$)

If Not ae Then
    If c$ = "msgbox" Then 'a messagebox
        'MsgBox Parms(1).Value
        MsgBoxToUser Parms(1).Value, scriptdata
    
    ElseIf c$ = "msgbox2" Then 'a messagebox
        Debug.Print Parms(1).Value
        'MsgBox Parms(1).Value
    
    ElseIf c$ = "numparams" Then 'num of parms
        ret$ = Ts(NumUserParams)
        'MsgBox Parms(1).Value
        
    ElseIf c$ = "broadcast" Then 'broadcast
        G$ = Parms(1).Value
        SendRCONCommand "say " + G$
   

    
    ElseIf c$ = "messagefull" Then 'messaging in colour
        
        If NumParms = 14 Then
        
            SendMessage Parms(1).Value, 1, Val(Parms(2).Value), Val(Parms(3).Value), Val(Parms(4).Value), Val(Parms(5).Value), Val(Parms(6).Value), Val(Parms(7).Value), Val(Parms(8).Value), Val(Parms(9).Value), Val(Parms(10).Value), Val(Parms(11).Value), Val(Parms(12).Value), Val(Parms(13).Value), Val(Parms(14).Value)
        
        End If
    
    ElseIf c$ = "param" Then 'asking for a parameter
        dd = Val(Parms(1).Value)
        If dd >= 1 And dd <= NumUserParams Then ret$ = UserParms(dd).Value
    
    ElseIf c$ = "return" Then 'return
        ret$ = Parms(1).Value
    
    ElseIf c$ = "logtofile" Then 'return
        scriptdata.LogToFile = True
        
    'TIME FUNCTIONS
    
    ElseIf c$ = "minutes" Then 'get number of minutes in this amount of seconds
        dd = Val(Parms(1).Value)
        ret$ = Ts(dd \ 60)
    
    ElseIf c$ = "seconds" Then 'get number of seconds after minutes
        dd = Val(Parms(1).Value)
        ret$ = Ts(dd Mod 60)
    
    ElseIf c$ = "addtimer" Then 'adds a second timer
        dd = Val(Parms(2).Value)
        G$ = Parms(1).Value
        
        If dd >= 0 And G$ <> "" And SetVar(G$, "1", scriptdata) = True Then
            'add a timer
            kk = 0
            For i = 1 To NumTimerVars
                If LCase(G$) = LCase(TimerVars(i)) Then kk = i
            Next i
            
            If kk = 0 Then
                NumTimerVars = NumTimerVars + 1
                ReDim Preserve TimerVars(0 To NumTimerVars)
                TimerVars(NumTimerVars) = G$
            End If
            SetVar G$, Ts(dd), scriptdata

        End If
    ElseIf c$ = "importents" Then
        If NumParms = 2 Then
            h = FreeFile
Close h
            G$ = ""
            If CheckForFile(Parms(1).Value) Then
                Open Parms(1).Value For Binary As h
                    Do Until EOF(h)
                        G$ = G$ + Input(65000, #h)
                    Loop
                Close h
                
                SetBSPEnts Parms(2).Value, G$
            End If
        End If
    
    ElseIf c$ = "exportents" Then
        If NumParms = 2 Then
            h = FreeFile
Close h
            G$ = ""
            If CheckForFile(Parms(1).Value) Then Kill Parms(1).Value
            
            If CheckForFile(Parms(2).Value) Then
                            
                G$ = GetBSPEnts(Parms(2).Value)
            
                Open Parms(1).Value For Binary As h
                    Put #h, , G$
                Close h
            End If
        End If
    
    ElseIf c$ = "checkfornamebind" Then
        'see if someone is using one of those damned %i binds.
        If NumParms = 1 Then
        
        
            ret$ = "0"
            j = 0
            For i = 1 To NumPlayers
                If InStr(1, Parms(1).Value, Players(i).Name) Then j = i: Exit For
            Next i
            
            ret$ = Ts(j)
    
        End If
    'STRING FUNCTIONS
    ElseIf c$ = "kicksomeone" Then
        KickSomeOne scriptdata
    
    ElseIf c$ = "reverse" Then
        'reverse this string
        G$ = Parms(1).Value
        ret$ = StrReverse(G$)
    
    ElseIf c$ = "runformula" Then
    
        Dim NewScriptData As typScriptData
        NewScriptData.TimeStarted = Timer
        
        Dim NewParms() As typParams
        
        G$ = UCase(Parms(1).Value)
        
        NewScriptData.StartedName = "runformula: " + G$
        
        If Parms(2).Value <> "1" Then
            For i = Asc("A") To Asc("Z")
                G$ = Replace(G$, Chr(i), "")
            Next i
        Else
            G$ = Replace(LCase(G$), "rcon", "")
        End If
                
        ret$ = TrimQuotes(ExecFormula(G$, NewParms, 0, NewScriptData))
        
    
    ElseIf c$ = "rnd" Then 'asking for a character
        ret$ = Ts(Rnd)
    
    ElseIf c$ = "randomize" Then 'asking for a character
        Randomize
    
    ElseIf c$ = "nocreate" Then 'asking for a character
        scriptdata.NoAutoCreate = True
    
    ElseIf c$ = "chr" Then 'asking for a character
        dd = Val(Parms(1).Value)
        If dd >= 0 And dd < 255 Then ret$ = Chr(dd)
    
    ElseIf c$ = "hexify" Then 'hex
        G1$ = Parms(1).Value
        ret$ = ReadyForDLL(G1$)
   
    ElseIf c$ = "makepoints" Then 'hex
        
        MakeRealPlayersPointList Parms(1).Value, Val(Parms(2).Value)
     
    ElseIf c$ = "addtoadminchat" Then 'adding to admin chat
        G1$ = Parms(1).Value
   
        SendChatToAll G1$, RGB(255, 255, 255), "<SERVER>", Time$
        
    ElseIf c$ = "setevent" Then 'setting an event
        
        G$ = Parms(1).Value ' event name
        G1$ = Parms(2).Value ' what to do
        dd = Val(Parms(3).Value) ' the value
          
        'find event
        For i = 1 To NumEvents
            If LCase(Events(i).Name) = LCase(G$) Then j = i: Exit For
        Next i
        
        If j > 0 Then
            Dim nextrun As Date
            If G1$ = "minutes" Then nextrun = Now + CDate(dd / 60 / 24)
            If G1$ = "hours" Then nextrun = Now + CDate(dd / 24)
            If G1$ = "days" Then nextrun = Now + CDate(dd)
            If G1$ = "seconds" Then nextrun = Now + CDate(dd / 60 / 60 / 24)
            
            If nextrun <> 0 Then
                Events(j).FirstCheck = nextrun
            End If
        End If
   
    ElseIf c$ = "dotalk" Then 'takin
        G$ = Parms(1).Value
        DoTalk G$
   
    ElseIf c$ = "specialadminspeech" Then 'takin
        G$ = Parms(1).Value
        G1$ = Parms(2).Value
        G2$ = Parms(3).Value
        
        ret$ = SpecialAdminSpeech(G$, G1$, us$)
   
   
    ElseIf c$ = "log" Then 'asking for a character
        dd = Val(Parms(1).Value)
        ret$ = Ts(Log(dd))
   
   
    ElseIf c$ = "sin" Then 'math function
        dd = Val(Parms(1).Value)
        If dd >= 0 And dd < 255 Then ret$ = Ts(Sin(dd / (180 / 3.14159)))
    
    ElseIf c$ = "cos" Then 'math function
        dd = Val(Parms(1).Value)
        If dd >= 0 And dd < 255 Then ret$ = Ts(Cos(dd / (180 / 3.14159)))
    
    ElseIf c$ = "tan" Then 'math function
        dd = Val(Parms(1).Value)
        If dd >= 0 And dd < 255 Then ret$ = Ts(Tan(dd / (180 / 3.14159)))
    
    ElseIf c$ = "arctan" Then 'math function
        dd = Val(Parms(1).Value)
        If dd >= 0 And dd < 255 Then ret$ = Ts(Atn(dd) * (180 / 3.14159))
    
    ElseIf c$ = "sqrt" Then 'square root
        dd = Val(Parms(1).Value)
        If dd < 0 Then
            
            ret$ = Ts(Sqr(-dd)) + "i"
        
        Else
            ret$ = Ts(Sqr(dd))
    
        End If
    
    ElseIf c$ = "modnum" Then 'Mod Code for BillDoor
        dd = Val(Parms(1).Value)
        dd2 = Val(Parms(2).Value)
        ret$ = Ts(dd Mod dd2)
     
    ElseIf c$ = "round" Then 'rounding
        dd = Val(Parms(1).Value)
        dd2 = Val(Parms(2).Value)
        ret$ = Ts(Round(dd, dd2))
        
    ElseIf c$ = "asc" Then 'ascii of a character
        G$ = Parms(1).Value
        If G$ <> "" Then ret$ = Ts(Asc(G$))
    
    ElseIf c$ = "dllmode" Then 'is dll on
        ret$ = Str(DLLEnabled)
    
    ElseIf c$ = "lcase" Then 'lower case
        G$ = Parms(1).Value
        If G$ <> "" Then ret$ = LCase(G$)
    
    ElseIf c$ = "ucase" Then 'lower case
        G$ = Parms(1).Value
        If G$ <> "" Then ret$ = UCase(G$)

    ElseIf c$ = "trim" Then 'trim the string
        G$ = Parms(1).Value
        If G$ <> "" Then ret$ = Trim(G$)

    ElseIf c$ = "ltrim" Then 'trim the left side
        G$ = Parms(1).Value
        If G$ <> "" Then ret$ = LTrim(G$)

    ElseIf c$ = "rtrim" Then 'trim the right side
        G$ = Parms(1).Value
        If G$ <> "" Then ret$ = RTrim(G$)

    ElseIf c$ = "val" Then 'trim the right side
        G$ = Parms(1).Value
        If G$ <> "" Then ret$ = Ts(Val(G$))

    ElseIf c$ = "int" Then 'trim the right side
        G$ = Parms(1).Value
        If G$ <> "" Then ret$ = Ts(Int(Val(G$)))
        
    ElseIf c$ = "mid" Then 'get a character
        G$ = Parms(1).Value
        n1 = Val(Parms(2).Value)
        n2 = Val(Parms(3).Value)
        If n2 <= 0 Then n2 = 1
        If G$ <> "" And n1 + n2 - 1 <= Len(G$) And n1 > 0 Then ret$ = Mid(G$, n1, n2)
    
    ElseIf c$ = "left" Then 'get characters on the left
        G$ = Parms(1).Value
        n1 = Parms(2).Value
        If G$ <> "" And n1 <= Len(G$) And n1 > 0 Then ret$ = Left(G$, n1)

    ElseIf c$ = "right" Then 'get characters on the right
        G$ = Parms(1).Value
        G1 = Parms(2).Value
        If G$ <> "" And G1 <= Len(G$) And G1 > 0 Then ret$ = Right(G$, G1)

    ElseIf c$ = "len" Then 'length of the string
        G$ = Parms(1).Value
        ret$ = Ts(Len(G$))
    
    ElseIf c$ = "instr" Then 'check location in string
        G1 = Val(Parms(1).Value)
        G2$ = Parms(2).Value
        g3$ = Parms(3).Value
        If G1 > 0 And G1 <= Len(G2$) Then ret$ = Ts(InStr(G1, G2$, g3$))

    ElseIf c$ = "wait" Then 'Wait X amount of seconds
        G1 = Val(Parms(1).Value)
        If G1 > 0 Then
            
            Dim tmel As Double
            strt = Timer
            Do
                DoEvents
                tmel = Round(Timer - strt, 3)
                If ScriptCheck(scriptdata) Then Exit Function
            Loop Until tmel > G1 Or tmel < 0
        End If

    ElseIf c$ = "instrrev" Then 'trim the left side
        G1 = Parms(1).Value
        G2$ = Parms(2).Value
        g3$ = Parms(3).Value
        If G1 = 0 Then G1 = -1
        If G1 > 0 And G1 <= Len(G2$) Then ret$ = Ts(InStrRev(G2$, g3$, G1))
    
    ElseIf c$ = "replacestring" Then 'replace string
        s1$ = Parms(1).Value
        G2$ = Parms(2).Value
        g3$ = Parms(3).Value
        ret$ = Replace(s1$, G2$, g3$)

    
    ElseIf c$ = "namematch" Then
        B$ = ""
        For i = 1 To NumParms
            B$ = B$ + Parms(i).Value + " "
        Next i
        B$ = Trim(B$)
        For i = 1 To NumPlayers
            If LCase(LeftR(Players(i).Name, Len(B$))) = LCase(B$) Then Num = Num + 1: lastnum = i
            If LCase(B$) = LCase(Players(i).Name) Then specmd = 1: specnm = i
        Next i
        
        If specmd = 1 Then
            Num = 1
            lastnum = specnm
        End If
        
        
        If Num <> 1 And NumParms = 2 Then
            If Parms(1).Value = "#" Then
                lsst = FindPlayer(Parms(2).Value)
                If lsst > 0 Then
                    Num = 1: lastnum = lsst
                Else
                    ret$ = "0"
                End If
            End If
        End If
        
        If Num = 0 Then
            ret$ = "0"
        ElseIf Num > 1 Then
            ret$ = Ts(-Num)
        Else
            ret$ = Ts(Players(lastnum).UserID)
        End If
    
    'END STRING FUNCTIONS
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'FILE AND DIRECTORY OPERATIONS
    
    ElseIf c$ = "filelines" Then
        
        'Opens a file, and for each line in the file, executes a script once, passing it the line and the line number.
        If NumParms = 2 Then
            a$ = Parms(1).Value
            B$ = Parms(2).Value
            
            If CheckForFile(a$) Then
                Num = 0
                h = FreeFile
Close h
                ret$ = ""
                Open a$ For Input As h
                    Do While Not (EOF(h))
                        Line Input #h, G$
                        Num = Num + 1
                                    
                        ExecFunctionScript2 B$, 2, scriptdata, G$, Ts(Num)
                    Loop
                Close h
            Else
                ret$ = "-1"
                AddToLog "            File " + a$ + " not found!" + vbCrLf
            End If
            
        End If
    
    
    
    
    
    
    ElseIf c$ = "makefile" Then
        If NumParms = 2 Then
            'This routine creates a file (param1) and places in it param2
            'If file exists it deletes it
            a$ = Parms(1).Value
            B$ = Parms(2).Value
            
            If CheckForFile(a$) Then Kill a$
            
            h = FreeFile
Close h
            Open a$ For Append As h
                Print #h, B$
            Close h
            AddToLog "            File " + a$ + " created!" + vbCrLf
        End If
    ElseIf c$ = "filesize" Then
        If NumParms = 1 Then
            'This routine returns the size of a file
            
            a$ = Parms(1).Value
            
            If CheckForFile(a$) Then
                ret$ = Ts(FileLen(a$))
            Else
                AddToLog "            File " + a$ + " not found!" + vbCrLf
                ret$ = ""
            End If
        End If
    
    ElseIf c$ = "getvarfromfile" Then
        If NumParms = 2 Then
            'This routine checks for a file
            
            a$ = Parms(1).Value
            B$ = Parms(2).Value
            
            If CheckForFile(a$) Then
                ret$ = GetVarFromFile(a$, B$)
            Else
                ret$ = ""
            End If
        End If
    ElseIf c$ = "makefilex" Then
        If NumParms = 2 Then
            'This routine creates a file (param1) and places in it param2
            'If file exists it doesnt do it
            a$ = Parms(1).Value
            B$ = Parms(2).Value
            
            If CheckForFile(a$) Then
            Else
                h = FreeFile
Close h
                Open a$ For Append As h
                    Print #h, B$
                Close h
                AddToLog "            File " + a$ + " created!" + vbCrLf
            End If
            
        End If
    ElseIf c$ = "addtofile" Then
        If NumParms = 2 Then
            'This routine creates a file or opens an existing file (param1) and places in it param2
            'If file exists it doesnt do it
            a$ = Parms(1).Value
            B$ = Parms(2).Value
            
            h = FreeFile
Close h
            Open a$ For Append As h
                Print #h, B$
            Close h
            AddToLog "            Added to file " + a$ + "!" + vbCrLf
    
        End If

    ElseIf c$ = "copyall" Then
        If NumParms = 2 Then
            'This routine copies all files from folder param1 to folder param2.
            a$ = Parms(1).Value 'Must NOT have final slash
            B$ = Parms(2).Value
            
            If Right(a$, 1) = "\" Then a$ = Left(a$, Len(a$) - 1)
            If Right(B$, 1) = "\" Then B$ = Left(B$, Len(B$) - 1)
            
            d$ = Dir(a$ + "\*.*")
            
            i = 0
            Do Until d$ = ""
                FileCopy a$ + "\" + d$, B$ + "\" + d$
                d$ = Dir
                i = i + 1
            Loop
            AddToLog "            " + Ts(i) + " files copied from " + a$ + " to " + B$ + vbCrLf
        End If
    
    ElseIf c$ = "copyfile" Then
        If NumParms = 2 Then
            'This routine copies file1 to file2
            a$ = Parms(1).Value
            B$ = Parms(2).Value
            
            If Dir(a$) <> "" Then
            
                ' make sure folder exists
                
                f = InStrRev(B$, "\")
                
                If f > 0 Then fold$ = Left(B$, f)
                
                If Dir(B$, vbDirectory) <> "" Then
            
                    FileCopy a$, B$
                End If
            End If
        End If

    ElseIf c$ = "delete" Then
        If NumParms = 1 Then
            'This routine will delete file param1
            a$ = Parms(1).Value
    
            If CheckForFile(a$) Then
                Kill a$
                AddToLog "            File " + a$ + " successfully deleted." + vbCrLf
            Else
                AddToLog "            File " + a$ + " not found!" + vbCrLf
            End If
            
        End If
    ElseIf c$ = "mapcycle" Then
        If NumParms > 0 Then
            'This routine changes the map cycle. It needs at least 1 parameter.
            
            a$ = Server.BothPath + "\mapcycle.txt"
            
            If CheckForFile(a$) Then Kill a$
            
            h = FreeFile
Close h
            Open a$ For Append As h
                For i = 1 To NumParms
                    Print #h, Parms(i).Value
                Next i
            Close h
            AddToLog "            Mapcycle changed! " + Ts(NumParms) + " maps in cycle!" + vbCrLf
    
        End If
    ElseIf c$ = "mkdir" Then
        If NumParms = 1 Then
            'This creates a directory
            a$ = Parms(1).Value
                    
            If Right(a$, 1) = "\" Then a$ = Left(a$, Len(a$) - 1)
            If Dir(a$, vbDirectory) = "" Then MkDir a$
    
            AddToLog "            Directory " + a$ + " created!" + vbCrLf
    
        End If
    ElseIf c$ = "renamefile" Then
        If NumParms = 2 Then
            'Renames one file to another name, if the first file exists and the second one doesnt.
            a$ = Parms(1).Value
            B$ = Parms(2).Value
            
            If CheckForFile(a$) Then
                If Not CheckForFile(B$) Then
                    Name a$ As B$
                    AddToLog "            File " + a$ + " renamed to " + a$ + "successfully!" + vbCrLf
                Else
                    AddToLog "            Error! File " + B$ + " exists!" + vbCrLf
                End If
            Else
                AddToLog "            Error! File " + a$ + " not found!" + vbCrLf
            End If
        End If
    ElseIf c$ = "addtolog" Then
        If NumParms = 1 Then
            'Adds the param1 to the log file.
            AddToLog "            Added to log: " + Chr(34) + Parms(1).Value + Chr(34) + vbCrLf
        End If
    ElseIf c$ = "checkforfile" Then
        If NumParms = 1 Then
            a$ = Parms(1).Value
            'Checks to see if a file is present
            If CheckForFile(a$) Then
                ret$ = "1"
            Else
                ret$ = "0"
            End If
        End If
    
    ElseIf c$ = "filedatetime" Then
        If NumParms = 1 Then
            a$ = Parms(1).Value
            'Checks to see if a file is present
            If CheckForFile(a$) Then
                ret$ = Format$(FileDateTime(a$), "DD/MM/YYYY HH:MM:SS")
            Else
                ret$ = "0"
            End If
        End If
        
    ElseIf c$ = "removefromfile" Then
        If NumParms = 2 Then
            'removes any line starting with param2 from file param1
            
            a$ = Parms(1).Value
            B$ = Parms(2).Value
            
            If CheckForFile(a$) Then
                
                h = FreeFile
Close h
                
                n$ = ""
                Open a$ For Input As h
                    Do While Not (EOF(h))
                        Line Input #h, G$
                        If LCase(LeftR(Trim(G$), Len(B$))) <> LCase(B$) Then n$ = n$ + G$ + vbCrLf
                    Loop
                Close h
            
                'all done, replace original
                Kill a$
                Open a$ For Append As h
                    Print #h, n$
                Close h
            
            Else
                AddToLog "            File " + a$ + " not found!" + vbCrLf
            End If
            
        End If
    ElseIf c$ = "getfile" Then
        If NumParms = 1 Then
            'gets the entire string of the file and returns it
            a$ = Parms(1).Value
            
            If CheckForFile(a$) Then
                
                h = FreeFile
Close h
                ret$ = ""
                Open a$ For Binary As h
                    ret$ = ret$ + Input(65000, #h)
                Close h
            Else
                AddToLog "            File " + a$ + " not found!" + vbCrLf
            End If
            
        End If
        
    ElseIf c$ = "makedirlist" Then
        'scans a dir, and puts all the files matching param2 into a list in another file
        'param1 = file to write
        'param2 = dir to scan, including types
        'param3 = write extensions? 1 = yes
        
        If NumParms = 3 Then
            
            a$ = Parms(1).Value
            B$ = Parms(2).Value
            G$ = Parms(3).Value
            
            gg$ = Dir(B$)
            
            d$ = ""
            Do While gg$ <> ""
                
                If Val(G$) = 1 Then
                    e = InStr(1, gg$, ".")
                    If e > 1 Then
                        gg$ = Left(gg$, e - 1)
                    End If
                End If
                
                d$ = d$ + gg$ + vbCrLf
                gg$ = Dir
            Loop
            
            If CheckForFile(a$) Then Kill a$
                    
            h = FreeFile
Close h
            Open a$ For Append As h
                Print #h, d$
            Close h
            
        End If
        
    ElseIf c$ = "mergefiles" Then
        If NumParms = 2 Then
            'gets the entire string of the file and returns it
            a$ = Parms(1).Value
            B$ = Parms(2).Value
            
            If CheckForFile(a$) And CheckForFile(B$) Then
                
                h = FreeFile
Close h
                ret$ = ""
                Open B$ For Binary As h
                    G$ = G$ + Input(65000, #h)
                Close h
                
                Open a$ For Append As h
                    Print #h, G$
                Close h
            Else
                AddToLog "            File " + a$ + " not found!" + vbCrLf
            End If
            
        End If
    
    ElseIf c$ = "getfileline" Then
        If NumParms = 2 Then
            'gets line number param(2) in the file
            a$ = Parms(1).Value
            B$ = Parms(2).Value
            
            If CheckForFile(a$) Then
                Num = 0
                h = FreeFile
Close h
                ret$ = ""
                Open a$ For Input As h
                    Do While Not (EOF(h))
                        Line Input #h, G$
                        Num = Num + 1
                        If Num = Val(B$) Then ret$ = G$: Exit Do
                    Loop
                    
                Close h
            Else
                ret$ = "-1"
                AddToLog "            File " + a$ + " not found!" + vbCrLf
            End If
            
        End If
    
    ElseIf c$ = "getfilenumlines" Then
        If NumParms = 1 Then
            'gets number of lines in file
            a$ = Parms(1).Value
            
            
            If CheckForFile(a$) Then
                Num = 0
                h = FreeFile
Close h
                ret$ = ""
                Open a$ For Input As h
                    Do While Not (EOF(h))
                        Line Input #h, G$
                        Num = Num + 1
                        
                    Loop
                    
                Close h
                ret$ = Ts(Num)
            Else
                AddToLog "            File " + a$ + " not found!" + vbCrLf
            End If
            
        End If
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'SERVER OPERATIONS

    
    ElseIf c$ = "rcon" Then
        If NumParms = 1 Then
            'This routine sends an RCON command to the server.
            a$ = Parms(1).Value
            SendRCONCommand a$
            
            AddToLog "            RCON Command " + Chr(34) + a$ + Chr(34) + " sent!" + vbCrLf
        End If
    
    ElseIf c$ = "combine" Then
        If NumParms = 2 Then
            'combines parameters
            gg1 = Val(Parms(1).Value)
            gg2 = Val(Parms(2).Value)
            
            aa$ = ""
            If NumParms = 3 Then aa$ = Parms(3).Value
            If aa$ = "" Then aa$ = " "
            
            If gg1 <= NumUserParams And gg2 <= NumUserParams And gg1 <= gg2 And gg1 > 0 Then
            
                a$ = ""
                For i = gg1 To gg2
                    a$ = a$ + UserParms(i).Value + aa$
                Next i
            
                ret$ = a$
            End If
        End If
    
    
    
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'USER OPERATIONS

    
    ElseIf c$ = "sendtouser" Then
        If NumParms = 1 Then
            'Will send text directly to the users console (who activated this)
            a$ = Parms(1).Value
            SendToUser a$, scriptdata
        End If
    
    ElseIf c$ = "sendtodebug" Then
        If NumParms = 1 Then
            'Will send text directly to the users console (who activated this)
            a$ = Parms(1).Value
            SendToDebug a$, scriptdata
        End If
    
    ElseIf c$ = "sendtousername" Then
        If NumParms = 2 Then
            'Will send this text to the player with this name
            a$ = Parms(1).Value
            B$ = Parms(2).Value
                       
            For i = 1 To NumPlayers
                If LCase(Players(i).Name) = LCase(a$) Then j = i: Exit For
            Next i
                
            If j = 0 Then
                AddToLog "Player " + a$ + " not found on server!"
            Else
                If Players(j).IP = "" And Players(j).Port > 0 Then
                    AddToLog "Player " + a$ + "'s IP is unknown!"
                Else
                    SendToUserIP Players(j).IP, Players(j).Port, B$
                End If
            End If
        End If
    
     ElseIf c$ = "sendtouserid" Then
        If NumParms = 2 Then
            'Will send this text to the player with this userid
            a$ = Parms(2).Value
            G1 = Val(Parms(1).Value)
                       
            For i = 1 To NumPlayers
                If Players(i).UserID = G1 Then j = i: Exit For
            Next i
                
            If j = 0 Then
                AddToLog "Player with userid " + Ts(G1) + " not found on server!"
            Else
                If Players(j).IP = "" And Players(j).Port > 0 Then
                    AddToLog "Player with userid " + Ts(G1) + "'s IP is unknown!"
                Else
                    SendToUserIP Players(j).IP, Players(j).Port, a$
                End If
            End If
        End If
     
     ElseIf c$ = "sendtoall" Then
        If NumParms = 1 Then
            'Will send this text to the player with this userid
            a$ = Parms(1).Value
            
            For i = 1 To NumPlayers
                If Players(i).IP = "" And Players(i).Port > 0 Then
                Else
                    SendToUserIP Players(i).IP, Players(i).Port, a$
                End If
            Next i
        End If
    
    
     ElseIf c$ = "getuseridfromname" Then
        If NumParms = 1 Then
            'Gets the user id of the player with this name
            B$ = Parms(1).Value
                       
            For i = 1 To NumPlayers
                If LCase(Players(i).Name) = LCase(B$) Then j = i: Exit For
            Next i
                
            If j = 0 Then
                AddToLog "Player " + B$ + " not found on server!"
                ret$ = "0"
            Else
                ret$ = Ts(Players(j).UserID)
            End If
        End If
        
     ElseIf c$ = "getuseridfromip" Then
        If NumParms = 1 Then
            'Gets the user id of the player with this IP
            B$ = Parms(1).Value
                       
            For i = 1 To NumPlayers
                If LCase(Players(i).IP) = LCase(B$) Then j = i: Exit For
            Next i
                
            If j = 0 Then
                AddToLog "Player with IP " + B$ + " not found on server!"
                ret$ = "0"
            Else
                ret$ = Ts(Players(j).UserID)
            End If
        End If
     
     ElseIf c$ = "getuseridfromuid" Then
        If NumParms = 1 Then
            'Gets the user id of the player with this UniqueID
            B$ = Parms(1).Value
                       
            For i = 1 To NumPlayers
                If LCase(Players(i).UniqueID) = LCase(B$) Then j = i: Exit For
            Next i
                
            If j = 0 Then
                AddToLog "Player with UniqueID " + B$ + " not found on server!"
                ret$ = "0"
            Else
                ret$ = Ts(Players(j).UserID)
            End If
        End If
    
     ElseIf c$ = "getuserinfo" Then
        If NumParms = 2 Then
            'Gets various info
                                
            j = FindPlayer(Parms(1).Value)

            If j = 0 Then
                ret$ = "0"
                AddToLog "Player with userid " + Ts(G1) + " not found on server!"
            Else
                If Val(Parms(2).Value) = 1 Or Parms(2).Value = "class" Then ret$ = Ts(Players(j).Class)
                If Val(Parms(2).Value) = 2 Or Parms(2).Value = "connect" Then ret$ = CInt(Players(j).ConnectOnly)
                If Val(Parms(2).Value) = 3 Or Parms(2).Value = "entryname" Then ret$ = Players(j).EntryName
                If Val(Parms(2).Value) = 4 Or Parms(2).Value = "ip" Then ret$ = Players(j).IP
                If Val(Parms(2).Value) = 5 Or Parms(2).Value = "name" Then ret$ = Players(j).Name
                If Val(Parms(2).Value) = 6 Or Parms(2).Value = "port" Then ret$ = Ts(Players(j).Port)
                If Val(Parms(2).Value) = 7 Or Parms(2).Value = "realname" Then ret$ = Players(j).RealName
                If Val(Parms(2).Value) = 8 Or Parms(2).Value = "team" Then ret$ = Ts(Players(j).Team)
                If Val(Parms(2).Value) = 9 Or Parms(2).Value = "uniqueid" Then ret$ = Players(j).UniqueID
                If Val(Parms(2).Value) = 10 Or Parms(2).Value = "num" Then ret$ = Ts(j)
                If Val(Parms(2).Value) = 11 Or Parms(2).Value = "time" Then ret$ = Ts(Second(Now - Players(j).TimeJoined) + (60 * Minute(Now - Players(j).TimeJoined)) + (60 * 60 * Hour(Now - Players(j).TimeJoined)))
                If Val(Parms(2).Value) = 12 Or Parms(2).Value = "x" Then ret$ = Ts(Players(j).Pos.X)
                If Val(Parms(2).Value) = 13 Or Parms(2).Value = "y" Then ret$ = Ts(Players(j).Pos.Y)
                If Val(Parms(2).Value) = 14 Or Parms(2).Value = "z" Then ret$ = Ts(Players(j).Pos.Z)
                If Val(Parms(2).Value) = 15 Or Parms(2).Value = "shutup" Then ret$ = Ts(CInt(Players(j).ShutUp))
                If Val(Parms(2).Value) = 16 Or Parms(2).Value = "numkickvotes" Then ret$ = Ts(Players(j).NumKickVotes)
                If Val(Parms(2).Value) = 17 Or Parms(2).Value = "warn" Then ret$ = Ts(Players(j).Warn)
                If Val(Parms(2).Value) = 18 Or Parms(2).Value = "points" Then ret$ = Ts(GetPoints(j))
                If Val(Parms(2).Value) = 19 Or Parms(2).Value = "flag1" Then ret$ = Ts(GetFlag(j, 1))
                If Val(Parms(2).Value) = 20 Or Parms(2).Value = "flag2" Then ret$ = Ts(GetFlag(j, 2))
                If Val(Parms(2).Value) = 21 Or Parms(2).Value = "flag3" Then ret$ = Ts(GetFlag(j, 3))
                If Val(Parms(2).Value) = 22 Or Parms(2).Value = "flag4" Then ret$ = Ts(GetFlag(j, 4))
                
            End If
        End If
     
     ElseIf c$ = "setwarn" Then
        If NumParms = 2 Then
            'sets WARN flag
                                
            j = FindPlayer(Parms(1).Value)

            If j = 0 Then
                ret$ = "0"
                AddToLog "Player with userid " + Ts(G1) + " not found on server!"
            Else
                Players(j).Warn = Val(Parms(2).Value)
            End If
        End If
          
     ElseIf c$ = "getuser" Then
        If NumParms = 1 Then
            'Gets the user id of user number param1 (used to just get any user)
            G1 = Val(Parms(1).Value)
                       
            If G1 > NumPlayers Or G1 <= 0 Then
                ret$ = "0"
                AddToLog "Invalid user number (" + Ts(G1) + ")"
            Else
                ret$ = Ts(Players(G1).UserID)
            End If
        End If

  



    End If
End If

ret$ = Replace(ret$, Chr(34), Chr(255))

ExecuteScriptCommand = ret$

Exit Function
errocc:
ErrorReport Err.Number, Err.Description + ", " + Err.Source


End Function

Function GetPoints(Num) As Long

    'Retrieves the number of points this player has
    
    j = RealNameSearch2(Players(Num).UniqueID)
    
    If j > 0 Then
        GetPoints = Val(RealPlayers(j).Points)
    End If

End Function

Function GetFlag(Num, FlagNum) As Long

    'Retrieves the number of points this player has
    
    j = RealNameSearch2(Players(Num).UniqueID)
    
    If j > 0 Then
        
        
        'flag1 = 6, flag2 = 7, flag3 = 8, flag4 = 9
        nm = FlagNum + 5
        
        If CheckBit2(RealPlayers(j).Flags, nm) Then
            GetFlag = 1
        Else
            GetFlag = 0
        End If
    End If

End Function


Sub SetPoints(Num, Amt)

    'Retrieves the number of points this player has
    
    j = RealNameSearch2(Players(Num).UniqueID)
    
    If j > 0 Then
        If Amt > 99999999 Then Amt = 0
        
        RealPlayers(j).Points = Ts(Amt)
    End If

End Sub


Sub AddToLog(a$)


'Form6.Text1 = Form6.Text1 + Indent + a$
'If Len(Form6.Text1) > 3000 Then Form6.Text1 = Right(Form6.Text1, 2900)
'Form6.Text1.SelStart = Len(Form6.Text1)


End Sub

Function ExecuteScriptParams(p$, scriptdata As typScriptData, Optional NoSayMode As Boolean) As Boolean
If DebugMode Then LastCalled = "ExecuteScriptParams"

'takes a script exec command passed from RCON or whatever and interprits it

Dim UserParms() As typParams
ReDim UserParms(1 To 200)

ExecuteScriptParams = True

p$ = Trim(p$)

p$ = Replace(p$, "\q", Chr(34))

e = InStr(1, p$, " ")
                                        
                                        
If e = 0 Then 'No paramerers.
    Vars.Command = p$
    NumParams = 0
Else
    'There are parameters.
    'Start extracting them.
    
    'First the command:
    Vars.Command = Left(p$, e - 1)
    
    'Now the others:
    NumParams = 0
    
    
    i = 0
    'do  'replaced DO with FOR
    For jkk = 1 To 10000000
        G = e
        e = InStr(e + 1, p$, " ")
        If e = 0 Then e = Len(p$)
        
        i = i + 1
        r$ = Mid(p$, G + 1, e - G)
        r$ = Trim(r$)
                                                       
        'Now replace special commands in the parameter.
        'r$ = DoVars(r$)
        r$ = Replace(r$, "\\", Chr(245))
        r$ = Replace(r$, "\s", " ")
        r$ = Replace(r$, "\t", "~")
        r$ = Replace(r$, "\[", "(")
        r$ = Replace(r$, "\]", ")")
        r$ = Replace(r$, "\a", ";")
        r$ = Replace(r$, "\c", ",")
        r$ = Replace(r$, Chr(245), "\")
        
        UserParms(i).Value = r$
        
        'DEBUG STUFF
        
    'Loop Until e >= Len(p$)

        If e >= Len(p$) Then Exit For
    Next jkk
    If jkk = 10000000 Then AddToLogFile "!!! ERROR !!! INFINITE LOOP: Occured at " + LastCalled
    
    Num = i
    
    'MsgBox o$
End If

'TIME TO EXECUTE!
If Vars.Command <> "" Then
    AddToLog "Starting " + Vars.Command + vbCrLf
                                                
    abc = ExecuteScript(Vars.Command, UserParms, CInt(Num), nne$, scriptdata, NoSayMode)
    
    If abc = False Then
        
        nne$ = ExecuteCommand(p$, UserParms, CInt(Num), scriptdata)
        ExecuteScriptParams = False
    End If
End If

End Function
Function DoVars(r$) As String
If DebugMode Then LastCalled = "DoVars"

r$ = Replace(r$, Chr(34), "")
r$ = Replace(r$, "\\", "\")
r$ = Replace(r$, "\n", vbCrLf)
r$ = Replace(r$, "_", " ")
r$ = Replace(r$, "\q", Chr(34))
r$ = Replace(r$, "\t", "~")
r$ = Replace(r$, "\s", ";")
r$ = Replace(r$, "%hlpath%", Server.HLPath)
r$ = Replace(r$, "%numparams%", NumParams)
r$ = Replace(r$, "%gamedir%", Server.GamePath)
r$ = Replace(r$, "%password%", Server.RCONPass)
r$ = Replace(r$, "%map%", Vars.Map)
r$ = Replace(r$, "%userip%", Vars.UserIP)
r$ = Replace(r$, "%userport%", Vars.UserPort)
r$ = Replace(r$, "%username%", Vars.UserName)
r$ = Replace(r$, "%newestlog%", NewLastLog)

'r$ = ReplaceUserVars(r$)

DoVars = r$

End Function
Function DoVars2(r$) As String
If DebugMode Then LastCalled = "DoVars2"

'r$ = Replace(r$, Chr(34), "")
r$ = Replace(r$, "\\", "\")
r$ = Replace(r$, "\n", vbCrLf)
'r$ = Replace(r$, "_", " ")
r$ = Replace(r$, "\q", Chr(34))
r$ = Replace(r$, "\t", "~")
r$ = Replace(r$, "\s", ";")
r$ = Replace(r$, "%hlpath%", Server.HLPath)
r$ = Replace(r$, "%numparams%", NumParams)
r$ = Replace(r$, "%gamedir%", Server.GamePath)
r$ = Replace(r$, "%password%", Server.RCONPass)
r$ = Replace(r$, "%map%", Vars.Map)
r$ = Replace(r$, "%userip%", Vars.UserIP)
r$ = Replace(r$, "%userport%", Vars.UserPort)
r$ = Replace(r$, "%username%", Vars.UserName)
r$ = Replace(r$, "%newestlog%", NewLastLog)

'r$ = ReplaceUserVars(r$)

DoVars2 = r$

End Function


Function ReplaceParams(ByVal Txt As String)

from_str = "%param"
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
            pos2 = InStr(Pos + 1, Txt, "%")
            q = pos2 - Pos + 1
            
            w$ = Mid(Txt, Pos + 6, pos2 - Pos - 6)
            X = Val(w$)
                        
            If X = 0 Then
                to_str = ""
            Else
                to_str = Params(X)
            End If
                        
            new_txt = new_txt & Left$(Txt, Pos - 1) & to_str
            Txt = Mid$(Txt, Pos + q)
        End If
    Loop

    ReplaceParams = new_txt
End Function

Sub SaveCommands(Optional Num As Integer)
If DebugMode Then LastCalled = "SaveCommands - " + Ts(Num)

'Saves commands to file

s = 1
e = 18

h = FreeFile
Close h

If Num > 0 Then s = Num: e = Num

For i = s To e
    If CheckForFile(Data(i)) Then Kill Data(i)
    If i <> 6 And i <> 1 And i <> 11 And i <> 2 Then Open Data(i) For Binary As h

        Select Case i
        Case 1
            'Put #h, , NumCommands
            'Put #h, , Commands
        Case 2
            ''Put #h, , NumUsers
            'Put #h, , Users
        Case 3
            Put #h, , NumKickBans
            Put #h, , KickBans
        Case 4
            Put #h, , NumClans
            Put #h, , Clans
        Case 5
            Put #h, , NumKills
            Put #h, , KillList
        Case 6
            'Put #h, , NumRealPlayers
            'Put #h, , RealPlayers
        Case 7
            Put #h, , Web
            Put #h, , General
        Case 8
            Put #h, , NumEvents
            Put #h, , Events
        Case 9
            Put #h, , NumMessages
            Put #h, , Messages
        Case 10
            Put #h, , NumSpeech
            Put #h, , Speech
        Case 11
            'Put #h, , NumSwears
            'Put #h, , Swears
        Case 12
            Put #h, , NumRealPlayers
            Put #h, , RealPlayers
        Case 13
            Put #h, , NumUsers
            Put #h, , Users
        Case 14
            Put #h, , PointData
        Case 15
            Put #h, , NumCommands
            Put #h, , Commands
        Case 16
            Put #h, , NumSwears
            Put #h, , Swears
        Case 17
            Put #h, , NumAdminBMP
            Put #h, , AdminBMP
        Case 18
            Put #h, , ServerStart
        End Select
    Close h
Next i
'text for ASSISTANT.CFG
a$ = "// Avatar-X's Server Assistant Configuration File" + vbCrLf _
+ "// Fill this out if its the FIRST TIME running the server" + vbCrLf + "// These changes can also be made remotly" + vbCrLf + vbCrLf _
+ "// Half-Life Path" + vbCrLf + "// Example: hl_path " + Chr(34) + "c:\sierra\half-life" + Chr(34) + vbCrLf + vbCrLf _
+ "hl_path " + Chr(34) + Server.HLPath + Chr(34) + vbCrLf + vbCrLf + "// Game directory that you are using" + vbCrLf _
+ "// Example: game_dir " + Chr(34) + "tfc" + Chr(34) + vbCrLf + vbCrLf + "game_dir " + Chr(34) + Server.GamePath + Chr(34) + vbCrLf _
+ vbCrLf + "// Port of game server" + vbCrLf + "// Example: game_port " + Chr(34) + "27015" + Chr(34) + vbCrLf _
+ vbCrLf + "game_port " + Chr(34) + Server.ServerPort + Chr(34) + vbCrLf + vbCrLf _
+ "// IP of the server" + vbCrLf + "// Example: server_ip " + Chr(34) + "24.12.12.12" + Chr(34) + vbCrLf + vbCrLf _
+ "server_ip " + Chr(34) + Server.LocalIP + Chr(34) + vbCrLf + vbCrLf + "// RCON port on which the Assistant listens. CANNOT BE THE SAME AS GAME_PORT!!" + vbCrLf _
+ "// Example: rcon_listen_port " + Chr(34) + "25500" + Chr(34) + vbCrLf + vbCrLf + "rcon_listen_port " + Chr(34) + Server.RconListenPort + Chr(34) + vbCrLf _
+ vbCrLf + "// TCP port on which clients can connect" + vbCrLf + "// Example: local_tcp_port " + Chr(34) + "26000" + Chr(34) + vbCrLf _
+ vbCrLf + "local_tcp_port " + Chr(34) + Server.LocalConnectPort + Chr(34)

Dim DataFile4 As String
DataFile4 = App.Path + "\assistant.cfg"

If CheckForFile(DataFile4) Then Kill DataFile4
Open DataFile4 For Binary As h
    Put #h, , a$
Close h


End Sub

Function SendICQMessage(UINTo As String, Body As String)

Dim body2 As String

For i = 1 To Len(Body)
    a$ = Mid(Body, i, 1)
    e = Asc(a$)
    If (e >= 65 And e <= 90) Or (e >= 97 And e <= 122) Or a$ = " " Or (e >= 48 And e <= 57) Then
        If a$ = " " Then
            body2 = body2 + "+"
        Else
            body2 = body2 + a$
        End If
    Else
        c$ = Hex(e)
        If Len(c$) = 1 Then c$ = "0" + c$
        body2 = body2 + "%" + c$
    End If
Next i

Form1.TCP2.Close

ICQURL = "http://wwp.icq.com/scripts/WWPMsg.dll?from=Server&fromemail=None&subject=Server+Message&body=" + body2 + "&to=" + UINTo
Form1.TCP2.Connect

End Function


Function LoadCommands() As Boolean
If DebugMode Then LastCalled = "LoadCommands"

h = FreeFile
Close h

If CheckForFile(Data(11)) Then
    
    Dim OldSwears() As typOldSwearWords
    
    Open Data(11) For Binary As h
        Get #h, , NumSwears
        ReDim OldSwears(0 To NumSwears)
        Get #h, , Swears
    Close h
    
    ReDim Swears(0 To NumSwears)
    
    For i = 1 To NumSwears
        Swears(i).BadWord = OldSwears(i).BadWord
        Swears(i).Flags = OldSwears(i).Flags
    Next i
    
    If CheckForFile(Data(11) + "_old") Then Kill Data(11) + "_old"
    Name Data(11) As Data(11) + "_old"
    
    AddToLogFile "SERVER: Renamed " + Data(11) + " to " + Data(11) + "_old !"
    
End If

If CheckForFile(Data(1)) Then
    'old script file
    
    Dim OldCommands() As comds
    ReDim OldCommands(1 To 200)
    
    Open Data(1) For Binary As h
        'scripts
        Get #h, , NumCommands
        Get #h, , OldCommands
    Close h
    
    ReDim Commands(0 To NumCommands)
    
    For i = 1 To NumCommands
    
        Commands(i).NumButtons = OldCommands(i).NumButtons
        ReDim Commands(i).Buttons(0 To OldCommands(i).NumButtons)
        
        For j = 1 To OldCommands(i).NumButtons
    
            Commands(i).Buttons(j).ButtonName = OldCommands(i).Buttons(j).ButtonName
            Commands(i).Buttons(j).ButtonText = OldCommands(i).Buttons(j).ButtonText
            Commands(i).Buttons(j).OptionOff = OldCommands(i).Buttons(j).OptionOff
            Commands(i).Buttons(j).OptionOn = OldCommands(i).Buttons(j).OptionOn
            Commands(i).Buttons(j).Type = OldCommands(i).Buttons(j).Type
    
        Next j
        
        Commands(i).Exec = OldCommands(i).Exec
        Commands(i).MustHave = OldCommands(i).MustHave
        Commands(i).Name = OldCommands(i).Name
        Commands(i).NumParams = OldCommands(i).NumParams
        Commands(i).ScriptName = OldCommands(i).ScriptName
        Commands(i).Group = "No Group"
        Commands(i).AutoMakeVars = True
        
    Next i
    
    If CheckForFile(Data(1) + "_old") Then Kill Data(1) + "_old"
    Name Data(1) As Data(1) + "_old"
    
    AddToLogFile "SERVER: Renamed " + Data(1) + " to " + Data(1) + "_old !"
    
End If


If CheckForFile(Data(2)) Then
    
    Dim UsersOld(1 To 400) As typUsersOld
    
    Open Data(2) For Binary As h
        'scripts
        Get #h, , NumUsers
        Get #h, , UsersOld
    Close h

    'convert
    For i = 1 To NumUsers
        Users(i).Allowed = UsersOld(i).Allowed
        Users(i).Flags = UsersOld(i).Flags
        Users(i).Name = UsersOld(i).Name
        Users(i).PassWord = UsersOld(i).PassWord
    Next i
    
    If CheckForFile(Data(2) + "_old") Then Kill Data(2) + "_old"
    Name Data(2) As Data(2) + "_old"
    
    AddToLogFile "SERVER: Renamed " + Data(2) + " to " + Data(2) + "_old !"

End If

If CheckForFile(Data(3)) Then
    Open Data(3) For Binary As h
        Get #h, , NumKickBans
        Get #h, , KickBans
    Close h
End If
If CheckForFile(Data(4)) Then
    Open Data(4) For Binary As h
        'clan protection
        Get #h, , NumClans
        Get #h, , Clans
    Close h
End If
If CheckForFile(Data(5)) Then
    Open Data(5) For Binary As h
        Get #h, , NumKills
        ReDim KillList(0 To NumKills)
        Get #h, , KillList
    Close h
End If


If CheckForFile(Data(6)) Then ' OLD REALPLAYERS

    Dim RealPlayersOld() As typRealPlayerOLD

    Open Data(6) For Binary As h
        Get #h, , NumRealPlayers
        ReDim RealPlayersOld(0 To NumRealPlayers)
        Get #h, , RealPlayersOld
    Close h
    
    'convert
    ReDim RealPlayers(0 To NumRealPlayers)
    For i = 1 To NumRealPlayers
        RealPlayers(i).Flags = RealPlayersOld(i).Flags
        RealPlayers(i).RealName = RealPlayersOld(i).RealName
        RealPlayers(i).LastName = RealPlayersOld(i).LastName
        RealPlayers(i).LastTime = CDbl(RealPlayersOld(i).LastTime)
        RealPlayers(i).UniqueID = RealPlayersOld(i).UniqueID
    Next i
    
    If CheckForFile(Data(6) + "_old") Then Kill Data(6) + "_old"
    Name Data(6) As Data(6) + "_old"
    
    AddToLogFile "SERVER: Renamed " + Data(6) + " to " + Data(6) + "_old !"
End If

If CheckForFile(Data(7)) Then
    Open Data(7) For Binary As h
        Get #h, , Web
        Get #h, , General
    Close h
End If
If CheckForFile(Data(8)) Then
    Open Data(8) For Binary As h
        Get #h, , NumEvents
        ReDim Events(0 To NumEvents)
        Get #h, , Events
    Close h
End If

If CheckForFile(Data(9)) Then
    Open Data(9) For Binary As h
        'scripts
        Get #h, , NumMessages
        ReDim Messages(0 To NumMessages)
        Get #h, , Messages
    Close h
End If

If CheckForFile(Data(10)) Then
    Open Data(10) For Binary As h
        Get #h, , NumSpeech
        ReDim Speech(0 To NumSpeech)
        Get #h, , Speech
    Close h
End If

If CheckForFile(Data(12)) Then
    Open Data(12) For Binary As h
        'scripts
        Get #h, , NumRealPlayers
        ReDim RealPlayers(0 To NumRealPlayers)
        Get #h, , RealPlayers
    Close h
End If

If CheckForFile(Data(13)) Then ' NEW users
    Open Data(13) For Binary As h
        Get #h, , NumUsers
        Get #h, , Users
    Close h
End If
If CheckForFile(Data(14)) Then ' NEW users
    Open Data(14) For Binary As h
        Get #h, , PointData
    Close h
End If

If CheckForFile(Data(15)) Then
    Open Data(15) For Binary As h
        'scripts - new
        Get #h, , NumCommands
        ReDim Commands(0 To NumCommands)
        Get #h, , Commands
    Close h
End If

If CheckForFile(Data(16)) Then
    
    Open Data(16) For Binary As h
        Get #h, , NumSwears
        ReDim Swears(0 To NumSwears)
        Get #h, , Swears
    Close h
    
End If

If CheckForFile(Data(17)) Then
    
    Open Data(17) For Binary As h
        Get #h, , NumAdminBMP
        ReDim AdminBMP(0 To NumAdminBMP)
        Get #h, , AdminBMP
    Close h
    
End If

If CheckForFile(Data(18)) Then
    Open Data(18) For Binary As h
        Get #h, , ServerStart
    Close h
End If


'Load database file
'LoadDB

Dim DataFile4 As String
DataFile4 = App.Path + "\assistant.cfg"

f$ = GetVarFromFile(DataFile4, "hl_path")
If f$ <> "" Then Server.HLPath = f$

f$ = GetVarFromFile(DataFile4, "game_dir")
If f$ <> "" Then Server.GamePath = f$

f$ = GetVarFromFile(DataFile4, "game_type")
If f$ = "cstrike" Then Server.GameMode = 2

f$ = GetVarFromFile(DataFile4, "game_port")
If f$ <> "" Then Server.ServerPort = Val(f$)

f$ = GetVarFromFile(DataFile4, "server_ip")
If f$ <> "" Then Server.LocalIP = f$

f$ = GetVarFromFile(DataFile4, "rcon_listen_port")
If f$ <> "" Then Server.RconListenPort = Val(f$)

f$ = GetVarFromFile(DataFile4, "local_tcp_port")
If f$ <> "" Then Server.LocalConnectPort = Val(f$)


End Function

Function ReadyForDLL(ms$) As String

'If General.  = 0 Then
'
'    Msg$ = ms$
'    Msg$ = Replace(Msg$, " ", "|")
'    Msg$ = Replace(Msg$, "'", "`")
'    Msg$ = Replace(Msg$, ",", "~")
'    Msg$ = Replace(Msg$, "(", "*")
'    Msg$ = Replace(Msg$, ")", "\")
'    Msg$ = Replace(Msg$, ":", "^")
'    Msg$ = Replace(Msg$, "{", "*")
'    Msg$ = Replace(Msg$, "}", "\")
'Else

    
    Msg$ = "&" + MakeHex(Replace(ms$, "%", "%%"))
'End If

ReadyForDLL = Msg$

End Function

Function MakeHex(Str As String) As String

For i = 1 To Len(Str)
    
    G$ = Hex$(Asc(Mid(Str, i, 1)))
    If Len(G$) = 1 Then G$ = "0" + G$
    ff$ = ff$ + G$

Next i

MakeHex = ff$

End Function

Sub SendRCONCommand(c$, Optional mde As Integer, Optional Last As Integer)
If DebugMode Then LastCalled = "SendRconCommand"
    
    On Error GoTo errocc
    
    If UCase(LeftR(c$, 4)) = "SAY " And DLLEnabled = True And (CInt(mde) = 0 Or mde >= 3) Then
        'replace with the other one :)
        
        If mde = 0 Then
            mm$ = "<SERVER> " + RightR(c$, Len(c$) - 4)
            If Len(mm$) > 96 Then mm$ = LeftR(mm$, 96)
            mm$ = ReadyForDLL(mm$)
            
            If LastTalk = 0 Then LastTalk = 1
            mm$ = "sa_talk" + Ts(LastTalk) + " " + mm$
            LastTalk = LastTalk + 1
            If LastTalk >= 5 Then LastTalk = 1
           
            SendActualRcon mm$
            If Last = 0 Then SendActualRcon SA_CHECK
            
            SendToWatchers "SERVER", "", RightR(c$, Len(c$) - 4), 0, 0
        End If
        
        If mde >= 3 Then
            mm$ = RightR(c$, Len(c$) - 4)
                        
            If MesLastY = 0 Then MesLastY = 0.25
            If mde = 3 Then SendMessage mm$, 1, 255, 255, 1, 1, 1, 255, 2, 6, 2, 0.02, 1, -1, MesLastY
            If mde = 4 Then SendMessage mm$, 1, 1, 255, 1, 255, 1, 1, 2, 6, 2, 0.02, 1, -1, MesLastY
            If mde = 5 Then SendMessage mm$, 1, 1, 255, 255, 100, 100, 1, 2, 6, 2, 0.02, 1, -1, MesLastY
            'MesLastY = MesLastY + 0.03
            If MesLastY > 0.45 Then MesLastY = 0.2
        End If

    Else
    
        SendActualRcon c$
        
        If UCase(LeftR(c$, 4)) = "SAY " And mde <> 1 Then
            SendToWatchers "SERVER", "", RightR(c$, Len(c$) - 4), 0, 0
        End If
    
    End If
    
        
    Exit Sub
errocc:
    ErrorReport Err.Number, Err.Description + ", " + Err.Source

End Sub

Sub SendActualRcon(cmd$)

' add to the list of waiting RCON commands


n = UBound(WaitingCommands)

n = n + 1

ReDim Preserve WaitingCommands(0 To n)
WaitingCommands(n) = cmd$

DoAfterChallenge = True
GetChallenge


'If ChallengeNum <> "" Then
'
'    hed$ = Chr(255) + Chr(255) + Chr(255) + Chr(255) + "rcon " + ChallengeNum + " " + Chr(34) + Server.RCONPass + Chr(34) + " "
'
'    'rcon number \"password\" rconcommands
'
'    'cm$ = Replace(cmd$, " ", Chr(34) + " " + Chr(34))
'    cm$ = cmd$
'    'Form1.UDP1.SendData hed$ + " " + Chr(34) + Server.RCONPass + Chr(34) + " " + Chr(34) + cm$ + Chr(34) + Chr(255) + Chr(255) + Chr(255) + Chr(255)
'    Form1.UDP1.SendData hed$ + cm$ + Chr(255) + Chr(255) + Chr(255) + Chr(255)
'
'End If

End Sub

Sub GetChallenge()

    
    hed$ = Chr(255) + Chr(255) + Chr(255) + Chr(255)
    
    cmd$ = hed$ + "challenge rcon" + Chr(10)
                
    Form1.UDP1.SendData cmd$
    

End Sub

Sub AfterChallenge()

If Server.RCONPass = "" Then
    Server.RCONPass = "1"
    AlertAdmins "WARNING -- RCON PASS IS BLANK FOR SOME FUCKED UP REASON"
    AddToLogFile "ERROROROROROROR: BLANK RCON!"
    Exit Sub

End If

If Server.RCONPass = "1" Then
    Exit Sub
End If

If DoAfterChallenge Then




    'send all commands in the command buffer
    
    n = UBound(WaitingCommands)
    
    
    For i = 1 To n
    
        hed$ = Chr(255) + Chr(255) + Chr(255) + Chr(255) + "rcon " + ChallengeNum + " " + Chr(34) + Server.RCONPass + Chr(34) + " "
    
        'rcon number "password" rconcommands
        Form1.UDP1.SendData hed$ + WaitingCommands(i) + Chr(255) + Chr(255) + Chr(255) + Chr(255)
    Next i

    'clear array
    ReDim WaitingCommands(0 To 0)

    DoAfterChallenge = False
End If


End Sub


Function GetRCON(a$) As String
    If DebugMode Then LastCalled = "GetRCON"

    LastRCON = ""
    TimerVR = 10
    SendRCONCommand a$
    Form1.Timer1.Enabled = True
    Do
        DoEvents
    Loop Until TimerVR <= 0 Or LastRCON <> ""
    Form1.Timer1.Enabled = False
    
    GetRCON = LastRCON

End Function

Function CompileData() As String

'This function compiles into a single string all the data to be sent to the remote server
'which is mainly just the commands.

'format:
'(1)(password)(255)(number of commands)(255)(command name)(255)(number of params)(255)(paramsrequired)(255)(command exec)(255)(command name)... etc

'B$ = Chr(255)
'a$ = Chr(1) + Server.Password + B$ + Ts(NumCommands) + B$

'For i = 1 To NumCommands
'    a$ = a$ + Commands(i).Name + B$
'    a$ = a$ + Ts(Commands(i).NumParams) + B$
'    a$ = a$ + Ts(Commands(i).MustHave) + B$
'    a$ = a$ + Commands(i).Exec + B$
'Next i
'
'CompileData = a$

End Function

Function DeCompileData(a$) As Boolean

'This does the reverse of the CompileData function.
'It takes text sent to the server via UDP and converts it back.

'format:
'(number of commands)(255)(command name)(255)(number of params)(255)(command exec)(255)(command name)... etc
'
'B$ = Chr(255)
'
'a$ = Right(a$, Len(a$) - 1)
'
'e = InStr(1, a$, B$)
'pass$ = Mid(a$, 1, e - 1)
'
'If UCase(pass$) <> UCase(Server.Password) Then
'    AddToLog "Someone attempted to upload with the incorrect password!" + vbCrLf
'    FileBuffer = ""
'    Exit Function
'End If
'
'f = e
'e = InStr(f + 1, a$, B$)
'num = Val(Mid(a$, f + 1, e - f - 1))
'
'For i = 1 To num
'    g = e
'    e = InStr(e + 1, a$, B$)
'    Commands(i).Name = Mid(a$, g + 1, e - g - 1)
'    g = e
'    e = InStr(e + 1, a$, B$)
'    Commands(i).NumParams = Val(Mid(a$, g + 1, e - g - 1))
'    g = e
'    e = InStr(e + 1, a$, B$)
'    Commands(i).MustHave = Val(Mid(a$, g + 1, e - g - 1))
'    g = e
'    e = InStr(e + 1, a$, B$)
'    Commands(i).Exec = Mid(a$, g + 1, e - g - 1)
'Next i
'
'NumCommands = num
'
'DeCompileData = True

End Function



Function DeCompileEXE(DataFile$) As Boolean
'
'a$ = DataFile$
'
''r$ = Form1.Caption
''convert hex to dec
'rrr = Len(a$)
'c$ = ""
'
'For i = 0 To Len(a$) Step 2
'    j = i + 1
'
'    B$ = Mid(a$, j, 2)
'
'    'add
'    c$ = c$ + Chr(Hex2Dec(B$))
'    If i Mod 2000 = 0 Then
''        Form1.Caption = "De-encoding EXE: " + Ts(Int((I / Len(a$)) * 100)) + "%"
'        DoEvents
'
'    End If
'Next i
'
''Form1.Caption = r$
'
'd$ = App.Path + "\" + App.EXEName + "_new.exe"
'If Dir(d$) <> "" Then Kill d$
'
'Open d$ For Binary As #1
'    Put #1, , c$
'Close #1
'
'
'
'DeCompileEXE = True

End Function




Sub HandleEntry()
If DebugMode Then LastCalled = "HandleEntry"

'this sub will handle the log buffer - extract the first line and handle it.

a$ = LogBuffer


'look for the beginning string, which is "(255)(255)(255)(255)l"
Start$ = Chr(255) + Chr(255) + Chr(255) + Chr(255) + "l"
endstr$ = Chr(0)


Dim statpos
statpos = 0


e = InStr(1, a$, Start$)
f = InStr(e + 1, a$, endstr$)

If e > 0 And f > e Then
    'extract the line
        
    r$ = Mid(a$, e, f - e + 2)
    If Len(r$) > 5 Then r$ = Right(r$, Len(r$) - 5)
    LogBuffer = Right(LogBuffer, Len(LogBuffer) - f)
    statpos = 1
    If Len(r$) > 3 Then
        If UCase(Left(r$, 4)) = "OG L" Then
    
            tmestart = Timer
            ReadLog r$
            statpos = 2
            If DebugTime Then
                h = FreeFile
                Close h
                
                Open App.Path + "\debugtime.log" For Append As h
                    Print #h, Date$ + " " + Time$ + " READLOG: took: " + Ts(Round(Timer - tmestart, 4)) + " sec, was: " + r$
                Close h
            End If
            
            If InStr(1, r$, Chr(34) + " say") Or InStr(1, r$, "Rcon") Then
                
                r$ = Replace(r$, Chr(10), vbCrLf)
                
            End If
            statpos = 3
        Else
            LastRCON = LastRCON + r$
            
            If NumRconMonitors > 0 Then
                'ensure this ISNT the answer someone else is waiting for
                                
                fg$ = Replace(r$, Chr(34) + " is " + Chr(34), " ")
                fg$ = Replace(fg$, Chr(34), "")
                                
                e = InStr(1, fg$, " ")
                If e > 1 And NumRconMonitors > 1 Then
                    fg$ = Left(fg$, e - 1)
                    fg$ = Trim(LCase(fg$))
                    
                    If InStr(1, RconMonitors(1).Command, fg$) = 0 Then
                        'its not the answer #1 was waiting for, find who's it is
                        j = 0
                        For i = 2 To NumRconMonitors
                            If InStr(1, RconMonitors(i).Command, fg$) Then j = i: Exit For 'its this persons
                        Next i
                        
                        If j > 1 Then
                            'remove all ppl before this person
                            For i = 1 To j - 1
                                RemoveFirstMonitor
                            Next i
                            'all set, now this person can get the answer he deserves
                        End If
                    End If
                End If
                
                'Check monitors
                If RconMonitors(1).IsTCP = False Then
                    SendToUserIP RconMonitors(1).IP, RconMonitors(1).Port, r$
                ElseIf RconMonitors(1).Index <> 0 Then
                    SendPacket "TY", Replace(r$, Chr(10), vbCrLf), RconMonitors(1).Index
                End If
                
                'remove
                RemoveFirstMonitor
            End If
        End If
    End If
    
Else
    statpos = 5
End If

If Len(LogBuffer) > 10000 Then

    AddToLogFile "!! STATPOS: got to " + Ts(statpos) + " and its: " + LeftR(LogBuffer, 100)
    LogBuffer = ""
End If

End Sub

Sub MainLoop()

On Error GoTo errocc

' The main loop

Do

    Form1.TimerGo
DoEvents
Loop


Exit Sub
errocc:
ErrorReport Err.Number, Err.Description + ", " + Err.Source


End Sub

Sub RemoveFirstMonitor()

NumRconMonitors = NumRconMonitors - 1
If NumRconMonitors > 1 Then
    For i = 1 To NumRconMonitors
        RconMonitors(i).Command = RconMonitors(i + 1).Command
        RconMonitors(i).Port = RconMonitors(i + 1).Port
        RconMonitors(i).IP = RconMonitors(i + 1).IP
        RconMonitors(i).Index = RconMonitors(i + 1).Index
        RconMonitors(i).IsTCP = RconMonitors(i + 1).IsTCP
        RconMonitors(i).TimeSent = RconMonitors(i + 1).TimeSent
    Next i
End If


ReDim Preserve RconMonitors(0 To NumRconMonitors)

End Sub

Function CheckIfInStr(Allowed As String, c$) As Boolean

Dim Allw As String
Allw = Allowed

'checks inside the ALLOWED text to see if this command is allowed


'loop thru allowed commands


Allw = vbCrLf + Allw
Y = 1
o = 1
'do  'replaced DO with FOR
For jkk = 1 To 10000000
    'Extract a line of text
    o = Y
    If Y <> Len(Allw) + 1 Then Y = InStr(Y + 1, Allw, vbCrLf)

    'Handle the last line events, etc
    If Y = Len(Allw) + 1 Then
        Y = 0
    ElseIf Y = 0 Then
        Y = Len(Allw) + 1
    End If

    If Y > 0 And Y > (o + 2) Then

        'Get the command
        cmd$ = Mid(Allw, o + 2, Y - o)
        cmd$ = Trim(cmd$)
        If Right(cmd$, 2) = Chr(13) + Chr(10) Then cmd$ = Left(cmd$, Len(cmd$) - 2)
        cmd$ = Trim(cmd$)

        'Do a pattern match
        
        If LCase(c$) Like LCase(cmd$) Then CheckIfInStr = True: Exit Function
    End If
'Loop Until Y = 0
    If Y = 0 Then Exit For
Next jkk
If jkk = 10000000 Then AddToLogFile "!!! ERROR !!! INFINITE LOOP: Occured at " + LastCalled






End Function


Sub HandleUserRcon(mde As Boolean, a$, IP$, Prt, Index)
If DebugMode Then LastCalled = "HandleUserRcon"

'handles the first userrcon entry

If mde = False Then
    k = UBound(UserRCONBuffer)
    If k = 0 Then Exit Sub
    
    IP$ = UserRCONBuffer(1).IP
    Prt = UserRCONBuffer(1).Port
    a$ = UserRCONBuffer(1).Command
    
    'remove entry
    
    For i = 1 To k - 1
        UserRCONBuffer(i).IP = UserRCONBuffer(i + 1).IP
        UserRCONBuffer(i).Command = UserRCONBuffer(i + 1).Command
        UserRCONBuffer(i).Port = UserRCONBuffer(i + 1).Port
    Next i
    
    ReDim Preserve UserRCONBuffer(0 To k - 1)
End If


'-------------
'now handle it
'-------------

'CHECK IF ITS A CHALLENGE
'ÿÿÿÿchallenge rcon
'Beep
B$ = "ÿÿÿÿchallenge rcon" + Chr(10)
If LeftR(a$, Len(B$)) = B$ Then  'answer the challenge

    
    Form1.RconUDP.RemoteHost = IP$
    Form1.RconUDP.RemotePort = Prt
    Form1.RconUDP.SendData "ÿÿÿÿchallenge rcon " + ChallengeNum + Chr(10)
    
    Exit Sub
End If


'first scan for the password
'ÿÿÿÿrcon number "password" rconcommands

'a$ = Replace(a$, Chr(34), "")
'a$ = Replace(a$, Chr(0), "")

a$ = Left(a$, Len(a$) - 1)

'extract the password
e = InStr(1, a$, Chr(34))
f = InStr(e + 1, a$, Chr(34))

If e > 0 And f > e Then
    p$ = Mid(a$, e + 1, f - e - 1)
    
    'the command
    
    c$ = Right(a$, Len(a$) - f)
    c$ = Trim(c$)
    
    'Now, replace the characters that got F00KED in the rcon process:
    'c$ = Replace(c$, " ,", ",")
    'c$ = Replace(c$, " (", "(")
    'c$ = Replace(c$, " )", ")")
    'c$ = Replace(c$, " :", ":")
    
End If

'see if the used rcon pass is actually one for clan joining...



For i = 1 To NumClans
    If LCase(Clans(i).JoinPass) = LCase(p$) Then j = i: Exit For
Next i

If j <> 0 Then
    AddClanPlayer IP$, Prt, j
    Exit Sub
End If

j = 0
For i = 1 To NumUsers
    If LCase(Users(i).PassWord) = LCase(p$) Then j = i: Exit For
Next i

If j = 0 Then 'password not found
    SendToUserIP IP$, Prt, "Your password profile was not found.", CInt(Index)
    Exit Sub
End If

'found, now lets get the requested info

If CheckBit(j, 12) = False Then
    B$ = "Sorry, " + Users(j).Name + "! You aren't allowed to use RCON!"
    SendToUserIP IP$, Prt, B$, CInt(Index)
    Exit Sub
End If
'ensure the command is allowed

cc$ = Trim(LCase(c$))
e = InStr(1, c$, " ")
If e > 0 Then cc$ = Left(cc$, e - 1)

dec = CheckIfInStr(Users(j).Allowed, cc$)

If Users(j).Allowed <> "all" Then
    If dec = False And CheckBit(j, 23) = False Then
        B$ = "Sorry, " + Users(j).Name + "! You aren't allowed to use command " + Chr(34) + cc$ + Chr(34) + "!"
        SendToUserIP IP$, Prt, B$, CInt(Index)
        Exit Sub
    ElseIf dec And CheckBit(j, 23) = True Then
        B$ = "Sorry, " + Users(j).Name + "! You aren't allowed to use command " + Chr(34) + cc$ + Chr(34) + "!"
        SendToUserIP IP$, Prt, B$, CInt(Index)
        Exit Sub
    End If
End If

If Not mde Then AddToLogFile "RCON: User " + Chr(34) + Users(j).Name + Chr(34) + ", IP: " + Chr(34) + IP$ + Chr(34) + ", Command: " + Chr(34) + c$ + Chr(34)
If LeftR(c$, 7) <> "setgrid" And mde Then AddToLogFile "INTRCON: User " + Chr(34) + Users(j).Name + Chr(34) + ", IP: " + Chr(34) + IP$ + Chr(34) + ", Command: " + Chr(34) + c$ + Chr(34)

LastUser = j
Dim NewScriptData As typScriptData

NewScriptData.TimeStarted = Timer
NewScriptData.UserIP = IP$
NewScriptData.UserName = Users(j).Name
NewScriptData.UserPort = Prt
If mde Then
    NewScriptData.UserIsTCP = True
    NewScriptData.Index = Index
End If
NewScriptData.StartedName = "UserRCON: " + c$


asd = ExecuteScriptParams(c$, NewScriptData)


If asd Then Exit Sub

TimerVar2 = 0
Form1.Timer4.Enabled = True

'add a new monitor
NumRconMonitors = NumRconMonitors + 1
ReDim Preserve RconMonitors(0 To NumRconMonitors)

If mde Then
    RconMonitors(NumRconMonitors).IsTCP = True
    RconMonitors(NumRconMonitors).Index = Index
End If

RconMonitors(NumRconMonitors).Port = Prt
RconMonitors(NumRconMonitors).IP = IP$
RconMonitors(NumRconMonitors).Command = cc$
RconMonitors(NumRconMonitors).TimeSent = Timer

SendRCONCommand c$
B$ = "Hello, " + Users(j).Name + "! The requested information follows:"
'SendToUserIP IP$, Prt, B$
SendToUser B$, NewScriptData, True

LastUser = 0

End Sub


Function AddClanPlayer(IP$, Prt, k)
If DebugMode Then LastCalled = "AddClanPlayer"

'adds a new player to the clan
'find the player..

For i = 1 To NumPlayers
    If Players(i).IP = IP$ And Players(i).Port = Prt Then Num = i: Exit For
Next i

If Num = 0 Then
    SendToUserIP IP$, Prt, "Player info not found. Try reconnecting."
    Exit Function
End If

'got the stuff, add the dude
Players(Num).RemoveMe = False

For i = 1 To Clans(k).NumMembers
    If LCase(Clans(k).Members(i).Name) = LCase(Players(Num).Name) Then  'got us a player
        m = i
        Exit For
    End If
Next i

If m = 0 Then
    Clans(k).NumMembers = Clans(k).NumMembers + 1
    m = Clans(k).NumMembers
End If

'add the info
Clans(k).Members(m).LastIP = IP$
Clans(k).Members(m).Name = Players(Num).Name

If InStr(1, Clans(k).Members(m).UIN, Players(Num).UniqueID) = 0 Then
    Clans(k).Members(m).UIN = Clans(k).Members(m).UIN + "; " + Players(Num).UniqueID
End If

'tell the player
G$ = Chr(10) + Chr(10) + "You have been successfully added to clan " + Clans(k).Clan + "." + Chr(10)
G$ = G$ + "You may now reset the RCON_xxx values to whatever you want, and re-join the game, by typing RETRY." + Chr(10)

SendToUserIP IP$, Prt, G$

'save the settings
SaveCommands

'finally, remove the player
RemovePlayer Num

End Function

Sub ChangePlayerName(UserID As Integer, NewName As String)
If DebugMode Then LastCalled = "ChangePlayerName"

'Change the name of a player on the server

If DLLEnabled = False Then Exit Sub

'tell server the ID:
SendRCONCommand "sa_namesetid " + Ts(UserID)

'tell server the new name:
SendRCONCommand "sa_nameset " + ReadyForDLL(NewName)

SendActualRcon SA_CHECK

'w00t! done!

End Sub

Sub ChangePlayerClass(UserID As Integer, NewClass As Integer)
If DebugMode Then LastCalled = "ChangePlayerClass"

'Change the name of a player on the server

If DLLEnabled = False Then Exit Sub

'tell server the ID:
SendRCONCommand "sa_setclassid " + Ts(UserID)

'tell server the new name:
SendRCONCommand "sa_setclass " + Ts(NewClass)

SendActualRcon SA_CHECK

'w00t! done!

End Sub

Sub TalkToPlayer(UserID As Integer, Msg As String)
If DebugMode Then LastCalled = "TalkToPlayer"

'Send a message to only this player

If DLLEnabled = False Then Exit Sub

'tell server the ID:
SendRCONCommand "sa_talktoid " + Ts(UserID)

'tell server the new name:
SendRCONCommand "sa_talkto " + ReadyForDLL(Msg)

SendActualRcon SA_CHECK

'w00t! done!
End Sub

Sub RKillPlayer(UserID As Integer)
If DebugMode Then LastCalled = "RKillPlayer"

'Kill player - PVB ONLY

If DLLEnabled = False Then Exit Sub

'tell server the ID:
SendRCONCommand "sa_rkillid " + Ts(UserID)

SendActualRcon SA_CHECK

'w00t! done!

End Sub

Sub SendMessage(Message As String, Dynamic As Integer, _
Optional R1 As Byte, Optional G1 As Byte, Optional B1 As Byte, _
Optional R2 As Byte, Optional G2 As Byte, Optional B2 As Byte, _
Optional Effect As Integer, Optional HoldTime As Single, Optional FxTime As Single, _
Optional FadeInTime As Single, Optional FadeOutTime As Single, _
Optional X As Single, Optional Y As Single)
If DebugMode Then LastCalled = "SendMessage"

'sa_message_dynamic

'sends a coloured screen message
If DLLEnabled = False Then Exit Sub 'special DLL function only

'Check for missing elements and set defaults.

If R1 = 0 And G1 = 0 And B1 = 0 Then G1 = 255
If X = 0 Then X = -1
If Y = 0 Then Y = 0.75
If SvMes.Channel = 0 Then SvMes.Channel = 1

'ok, now send to server all needed changes
If R1 <> SvMes.Red1 And R1 > 0 Then SendRCONCommand "sa_message_red1 " + Ts(R1): SvMes.Red1 = R1
If G1 <> SvMes.Green1 And G1 > 0 Then SendRCONCommand "sa_message_green1 " + Ts(G1): SvMes.Green1 = G1
If B1 <> SvMes.Blue1 And B1 > 0 Then SendRCONCommand "sa_message_blue1 " + Ts(B1): SvMes.Blue1 = B1
If R2 <> SvMes.Red2 And R2 > 0 Then SendRCONCommand "sa_message_red2 " + Ts(R2): SvMes.Red2 = R2
If G2 <> SvMes.Green2 And G2 > 0 Then SendRCONCommand "sa_message_green2 " + Ts(G2): SvMes.Green2 = G2
If B2 <> SvMes.Blue2 And B2 > 0 Then SendRCONCommand "sa_message_blue2 " + Ts(B2): SvMes.Blue2 = B2
If Effect <> SvMes.Effect Then SendRCONCommand "sa_message_effect " + Ts(Effect): SvMes.Effect = Effect
If HoldTime <> SvMes.HoldTime Then SendRCONCommand "sa_message_holdtime " + Ts(HoldTime): SvMes.HoldTime = HoldTime
If FxTime <> SvMes.FxTime Then SendRCONCommand "sa_message_fxtime " + Ts(FxTime): SvMes.FxTime = FxTime
If FadeInTime <> SvMes.FadeInTime Then SendRCONCommand "sa_message_fadein " + Ts(FadeInTime): SvMes.FadeInTime = FadeInTime
If FadeOutTime <> SvMes.FadeOutTime Then SendRCONCommand "sa_message_fadeout " + Ts(FadeOutTime): SvMes.FadeOutTime = FadeOutTime
If X <> SvMes.X Then SendRCONCommand "sa_message_position_x " + Ts(X): SvMes.X = X
If Y <> SvMes.Y Then SendRCONCommand "sa_message_position_y " + Ts(Y): SvMes.Y = Y

SendRCONCommand "sa_message_channel " + Ts(SvMes.Channel)

SvMes.Channel = SvMes.Channel + 1
If SvMes.Channel > 4 Then SvMes.Channel = 1


e = InStr(1, Message, Chr(10))
If e = 0 Then e = Len(Message)

If e > 64 Then Message = LeftR(Message, 64)

'set message
SendActualRcon "sa_message " + ReadyForDLL(Message)
SendActualRcon SA_CHECK

SendToWatchers "MESSAGE", "", Message, 0, 0

'done!

End Sub

Function AddClanPlayer2(Num, scriptdata As typScriptData)
If DebugMode Then LastCalled = "AddClanPlayer2"

'adds a new player to the clan
'find the player..

If Num = 0 Then
    Exit Function
End If

'find the clan they are supposed to be in
'PROBLEM: What if player is in 2 clans?
'SOLUTION: Add them one at a time

nn$ = Players(Num).Name

kk = 1
checkagain:
k = 0
For i = kk To NumClans
    If InStr(1, nn$, Clans(i).Clan) Then k = i: Exit For
Next i

If k = 0 Then
    If kk = 1 Then
        SendToUser "Player's clan not found in clan list.", scriptdata
    Else
        If Len(cc$) > 2 Then cc$ = LeftR(cc$, Len(cc$) - 2)
        SendToUser "Player is already registered under clan(s) " + cc$ + ". No new clans could be found!", scriptdata
    End If
    Exit Function
End If

For i = 1 To Clans(k).NumMembers
    If LCase(Clans(k).Members(i).Name) = LCase(Players(Num).Name) Then  'got us a player
        m = i
        Exit For
    End If
Next i

If m = 0 Then
    Clans(k).NumMembers = Clans(k).NumMembers + 1
    m = Clans(k).NumMembers
End If

'add the info
Clans(k).Members(m).LastIP = Players(Num).IP
Clans(k).Members(m).Name = Players(Num).Name

If InStr(1, Clans(k).Members(m).UIN, Players(Num).UniqueID) = 0 Then
    Clans(k).Members(m).UIN = Clans(k).Members(m).UIN + "; " + Players(Num).UniqueID
Else
    'The player is already registered under this clan, check for other clans
    kk = k + 1
    cc$ = cc$ + Clans(k).Clan + ", "
    GoTo checkagain
End If

SendToUser "Player " + nn$ + " added to clan " + Clans(k).Clan + ".", scriptdata

End Function

Function SendToUser(a$, scriptdata As typScriptData, Optional mde As Boolean)
If DebugMode Then LastCalled = "SendToUser"

'This will send text to somoenes console, specifically the one to activate this command

If scriptdata.UserIsTCP = False Then
        
    If scriptdata.UserIP = "" Then Exit Function
    If scriptdata.UserPort = 0 Then Exit Function
        
    Form1.UDP3.RemoteHost = scriptdata.UserIP
    Form1.UDP3.RemotePort = scriptdata.UserPort

    If scriptdata.UserIP = "" Then Exit Function
    
    If mde = False Then a$ = "Server Message" + Chr(10) + "------------" + Chr(10) + a$ + Chr(10) + "------------"
    
    a$ = Replace(a$, vbCrLf, Chr(10))
    Form1.UDP3.SendData Chr(255) + Chr(255) + Chr(255) + Chr(255) + Chr(108) + a$ + Chr(10)
Else
    If mde = False Then SendPacket "TY", "Server Message" + vbCrLf + "------------" + vbCrLf + a$ + vbCrLf + "------------", scriptdata.Index
    If mde = True Then SendPacket "TY", a$, scriptdata.Index
End If


End Function

Function SendToDebug(a$, scriptdata As typScriptData)
If DebugMode Then LastCalled = "SendToDebug"

'This will send text to somoenes console, specifically the one to activate this command

If scriptdata.UserIsTCP = True Then
    SendPacket "DB", a$, scriptdata.Index
End If


End Function

Function MsgBoxToUser(a$, scriptdata As typScriptData)
If DebugMode Then LastCalled = "MsgBoxToUser"

'This will pop up a message to someone

If scriptdata.UserIsTCP = True Then
    SendPacket "MS", a$, scriptdata.Index
End If


End Function

Function SendToUserIP(IP As String, Port, Msg As String, Optional Index As Integer)
If DebugMode Then LastCalled = "SendToUserIP"

'This will send text to somoenes console

If Index > 0 Then
    SendPacket "TY", Msg, CInt(Index)
Else

    If IP <> "" And Port > 0 Then

        Form1.RconUDP.RemoteHost = IP
        Form1.RconUDP.RemotePort = Port
        
        Form1.RconUDP.SendData Chr(255) + Chr(255) + Chr(255) + Chr(255) + Chr(108) + Msg '+ Chr(10)
    End If
End If

End Function

Sub HelpFeature(UserParms() As typParams, NumUserParams, scriptdata As typScriptData)
If DebugMode Then LastCalled = "HelpFeature"

'This will help the user out.

If NumUserParams = 0 Then

    'send them a list of all the commands
    
    a$ = "The list of commands: " + vbCrLf
    
    For i = 1 To NumCommands
        If Commands(i).MustHave = 1 Then B$ = "required"
        If Commands(i).MustHave = 0 Then B$ = "suggested"
        
        a$ = a$ + Ts(i) + ": " + Commands(i).Name + " - " + Ts(Commands(i).NumParams) + " params " + B$ + vbCrLf
    
    Next i

    a$ = a$ + vbCrLf + "Type HELP (command or number of command) for more info on that command."
    
    SendToUser a$, scriptdata

Else
    c = Val(UserParms(1).Value)
    dd$ = UserParms(1).Value
    
    If c <= 0 Then
        For i = 1 To NumCommands
            If LCase(Commands(i).Name) = dd$ Then j = i: Exit For
        Next i
                
        If j = 0 Then SendToUser "Command not found!", scriptdata
        c = j
               
    ElseIf c > NumCommands Then
        SendToUser "Command not found!", scriptdata
        c = 0
    End If
        
    If c > 0 Then
        a$ = "Command Help on " + Commands(c).Name + vbCrLf
        If Commands(c).MustHave = 1 Then B$ = "required"
        If Commands(c).MustHave = 0 Then B$ = "suggested"
        
        a$ = a$ + Ts(Commands(c).NumParams) + " params " + B$ + vbCrLf + vbCrLf
        a$ = a$ + "The Code:" + vbCrLf
        a$ = a$ + Replace(Commands(c).Exec, "%", "#")
        
        SendToUser a$, scriptdata

    End If
End If



End Sub

Sub Defaults()
If DebugMode Then LastCalled = "Defaults"

    Server.LocalIP = "127.0.0.1"
    Server.ServerPort = "27015"
    Server.LocalConnectPort = "26000"
    
    
End Sub

Sub SetPorts()
If DebugMode Then LastCalled = "SetPorts"

    'set the Winsock controls' ports
    Form1.UDP1.RemotePort = Val(Server.ServerPort)
    Form1.UDP1.RemoteHost = Server.LocalIP
    Form1.RconUDP.LocalPort = Val(Server.RconListenPort)
    Form1.TCP1(0).LocalPort = Server.LocalConnectPort
    
    'start listening for connections
    Form1.TCP1(0).Listen
    Form1.RconUDP.Bind
    Form1.UDP1.Bind

    
    
End Sub

Sub StartLogWatch()
If DebugMode Then LastCalled = "StartLogWatch"

    GetInfo
    'SendRCONCommand "logaddress " + "24.100.160.29" + " " + Ts(Form1.UDP1.LocalPort)
    SendRCONCommand "logaddress " + Form1.UDP1.LocalIP + " " + Ts(Form1.UDP1.LocalPort)

End Sub


Public Sub Interprit(Txt As String, Index As Integer)
If DebugMode Then LastCalled = "Interprit"

'gets the stuff

'(254)(254)(254)(255)[CODE](255)[NAME](255)[PASSWORD](255)[PARAMS](255)(253)(253)(253)

e = InStr(1, Txt, Chr(255))
f = InStr(e + 1, Txt, Chr(255))

If e > 0 And f > e And f > 0 Then
    'code
    a$ = Mid(Txt, e + 1, f - e - 1)
    

    e = f
    f = InStr(e + 1, Txt, Chr(255))
    
    If e > 0 And f > e And f > 0 Then
        'name
        n$ = Mid(Txt, e + 1, f - e - 1)
    
        e = f
        f = InStr(e + 1, Txt, Chr(255))
        
        If e > 0 And f > e And f > 0 Then
            'pass
            pass$ = Mid(Txt, e + 1, f - e - 1)
        
            e = f
            f = InStrRev(Txt, Chr(255))
            
            If e > 0 And f > e And f > 0 Then
                'params
                p$ = Mid(Txt, e + 1, f - e - 1)
            End If
        End If
    End If
End If

'first check for username and password
For i = 1 To NumUsers
    If LCase(Users(i).Name) = LCase(n$) Then j = i: Exit For
Next i

For i = 1 To NumConnectUsers
    If ConnectUsers(i).Index = Index Then k = i: Exit For
Next i
If j > 0 Then
    ConnectUsers(k).PassWord = Users(j).PassWord
    ConnectUsers(k).UserNum = j
    ConnectUsers(k).Name = Users(j).Name
End If

If j = 0 Then
    SendPacket "MS", "Username not found.", Index
    ConnectUsers(k).RemoveMe = True
    Exit Sub
Else
    If LCase(Users(j).PassWord) <> LCase(pass$) Then
        SendPacket "IC", "", Index
        ConnectUsers(k).RemoveMe = True
        Exit Sub
    End If
    If CheckBit(j, 1) = False Then
        SendPacket "MS", "You are not allowed to connect to the server in this fashion.", Index
        ConnectUsers(k).RemoveMe = True
        Exit Sub
    End If
End If



If a$ = "X1" Then 'first packet
  
    SendPacket "X2", "", Index
    ConnectUsers(k).EncryptedMode = True
      Exit Sub
End If



'***************************************************************************************
'***************************************************************************************
'password checking done
'***************************************************************************************
'***************************************************************************************

Dim NewScriptData As typScriptData
NewScriptData.TimeStarted = Timer
NewScriptData.UserIP = ConnectUsers(k).IP
NewScriptData.UserName = Users(j).Name
NewScriptData.UserIsTCP = True
NewScriptData.Index = Index
NewScriptData.StartedName = "In INTERPRIT"

Vars.UserName = Users(j).Name



If a$ = "HL" Then 'hi message
    SendPacket "HI", "Welcome to the " + Server.HostName + " admin system.", Index
    AddToLogFile "LOGIN: " + Users(j).Name + " logged in."
    'send him last 20 lines, if allowed.
    If CheckBit(j, 17) Then PackageMessages2 Index, Users(j).Name
    PackageConnectPacket Index, CheckBit(j, 11)
    PackageAdminBMPList Index
    PackageMenuScripts Index
    SendUpdate
    'UpdateUsersList
    ExecFunctionScript "spec_adminlogin", 1, Users(j).Name

End If

If a$ = "EU" Then 'client is asking for all of the user info
    'check access:
    If CheckBit(j, 11) Then
        SendUserInfo Index
    Else
        SendPacket "MS", "You are not allowed to edit users.", Index
    End If
End If

If a$ = "EY" Then 'client is giving us more users
    'check access:
    If CheckBit(j, 11) Then
        'is allowed
        InterpritUsers p$
        SendPacket "MS", "User Update Recieved and Installed.", Index
        AddToLogFile "UPDATE: " + Users(j).Name + " updated the users."
    Else
        SendPacket "MS", "You are not allowed to edit users.", Index
    End If
End If

If a$ = "EC" Then 'client is asking for the scripts
    'check access:
    If CheckBit(j, 4) Then
        
        If Val(Replace(ConnectUsers(k).Version, ".", "")) >= 1136 Then
            PackageScripts Index
        Else
            SendPacket "MS", "You need version 1.1.36 or higher to edit scripts." + vbCrLf + "Ask your administrator for the newest client version.", Index
        End If
    Else
        SendPacket "MS", "You are not allowed to view or edit scripts.", Index
    End If
End If

If a$ = "ED" Then 'client is giving us scripts
    'check access:
    If CheckBit(j, 4) Then
        'is allowed
        UnPackageScripts p$
        SendPacket "MS", "Script update Recieved and Installed.", Index
        AddToLogFile "UPDATE: " + Users(j).Name + " updated the scripts."
    Else
        SendPacket "MS", "You are not allowed to edit scripts.", Index
    End If
End If

If a$ = "E1" Then 'client is giving us changed scripts
    'check access:
    If CheckBit(j, 4) Then
        'is allowed
        rt$ = UnPackageChangedScripts(p$)
        SendPacket "MS", "Script update for: " + vbCrLf + rt$ + vbCrLf + "Recieved and Installed.", Index
        AddToLogFile "UPDATE: " + Users(j).Name + " updated scripts " + rt$ + "."
    Else
        SendPacket "MS", "You are not allowed to edit scripts.", Index
    End If
End If

If a$ = "EX" Then 'client is sending dnew EXE
    'check access:
'    If CheckBit(j, 8) Then
'        'is allowed
'        R = DeCompileEXE(p$)
'        If R Then
'            SendPacket "MS", "EXE recieved, server shutting down in order to update.", Index
'            'finally start the thing
'            Shell App.Path + "\starter.exe server_new.exe", vbMinimizedNoFocus
'            End
'        Else
'            SendPacket "MS", "Error occured when decompiling EXE data.", Index
'        End If
'    Else
'        SendPacket "MS", "You are not allowed to upload EXE files.", Index
'    End If
End If

If a$ = "SU" Then 'client is asking for the players on the server
    'check access:
    If CheckBit(j, 13) Then
        SendRCONCommand "users"
        'PackagePlayers Index
    Else
        SendPacket "MS", "You are not allowed to view users on server.", Index
    End If
End If

If a$ = "SK" Then 'client is kicking a server user
    'check access:
    If CheckBit(j, 14) Then
        KickPlayer p$, Index
        AddToLogFile "INTKICK: " + Users(j).Name + " kicked a player."
    Else
        SendPacket "MS", "You are not allowed to kick users on server.", Index
    End If
End If

If a$ = "SB" Then 'client is banning a server user
    'check access:
    If CheckBit(j, 16) Then
        Vars.UserName = Users(j).Name
        
        UnPackageBanPlayer p$, Index, Users(j).Name, j
        'BanPlayer p$, Index, j
        'AddToLogFile "INTBAN: " + Users(j).Name + " banned a player."
    Else
        SendPacket "MS", "You are not allowed to ban users on server.", Index
    End If
End If

If a$ = "CP" Then 'client wants to change password
    'check access:
    If CheckBit(j, 9) Then
        ChangePass p$, j, Index
        AddToLogFile "PASSCHANGE: " + Users(j).Name + " changed their password."
    Else
        SendPacket "MS", "You are not allowed to change your password.", Index
    End If
End If

If a$ = "LL" Then 'client is changing log level
    'check access:
    If CheckBit(j, 10) Then
        ChangeLogLevel p$, k
    Else
        SendPacket "MS", "You are not allowed to monitor the server.", Index
    End If
End If

If a$ = "SY" Then 'client is talking
    'check access:
    If CheckBit(j, 2) Then
        If CheckBit(j, 34) Then
            ExecFunctionScript "spec_chatonlytalk", 2, (Users(j).Name), p$
            DoTalk "<CHAT " + (Users(j).Name) + "> " + p$, Len("<CHAT " + (Users(j).Name) + ">")
        Else
            If p$ <> "" Then
                Chat p$, (Users(j).Name)
            End If
        End If
    Else
        SendPacket "MS", "You are not allowed to chat with the server.", Index
    End If
End If

If a$ = "KB" Then 'client is asking for kick-ban list
    'check access:
    If CheckBit(j, 18) Then
        PackageKickBans Index
    Else
        SendPacket "MS", "You are not allowed to view or edit the kick-ban list.", Index
    End If
End If

If a$ = "KD" Then 'client is giving us the kick-ban list
    'check access:
    If CheckBit(j, 18) Then
        'is allowed
        UnPackageKickBans p$
        SendPacket "MS", "Kick-Ban update Recieved and Installed.", Index
        AddToLogFile "UPDATE: " + Users(j).Name + " updated the kick-ban lists."
    Else
        SendPacket "MS", "You are not allowed to edit the kick-ban list.", Index
    End If
End If

If a$ = "ES" Then 'client is asking for server info
    'check access:
    If CheckBit(j, 5) Then
        PackageServerInfo Index
    Else
        SendPacket "MS", "You are not allowed to view or edit the server info.", Index
    End If
End If

If a$ = "ER" Then 'client is returning server info
    'check access:
    If CheckBit(j, 5) Then
        'is allowed
        UnPackageServerInfo p$
        SendPacket "MS", "Server Info update Recieved, Server must be restarted.", Index
        AddToLogFile "UPDATE: " + Users(j).Name + " updated the server info."
    Else
        SendPacket "MS", "You are not allowed to edit the server info.", Index
    End If
End If

If a$ = "CL" Then 'client is asking for clan list
    'check access:
    If CheckBit(j, 19) Then
        PackageClans Index
    Else
        SendPacket "MS", "You are not allowed to edit clans.", Index
    End If
End If

If a$ = "CM" Then 'client is returning clan list
    'check access:
    If CheckBit(j, 19) Then
        'is allowed
        UnPackageClans p$
        SendPacket "MS", "Clan update recieved and installed.", Index
        AddToLogFile "UPDATE: " + Users(j).Name + " updated the clan lists."
    Else
        SendPacket "MS", "You are not allowed to edit clan lists.", Index
    End If
End If

If a$ = "RC" Then 'client is sending us some rcon!
    'format the packet
    rr$ = "rcon 213123123 " + Chr(34) + pass$ + Chr(34) + " " + p$ + Chr(10)
    HandleUserRcon True, rr$, ConnectUsers(k).IP, 0, Index
End If

If a$ = "SL" Then 'client is asking for speech list
    'check access:
    If CheckBit(j, 20) Then
        PackageSpeech Index
    Else
        SendPacket "MS", "You are not allowed to edit speech list.", Index
    End If
End If

If a$ = "SP" Then 'client is returning speech list
    'check access:
    If CheckBit(j, 20) Then
        'is allowed
        UnPackageSpeech p$
        SendPacket "MS", "Speech update recieved and installed.", Index
        AddToLogFile "UPDATE: " + Users(j).Name + " updated the speech list."
    Else
        SendPacket "MS", "You are not allowed to edit speech lists.", Index
    End If
End If

If a$ = "RP" Then 'client is asking for real player list
    'check access:
    If CheckBit(j, 21) Then
        PackageRealPlayers Index
    Else
        SendPacket "MS", "You are not allowed to edit real player list.", Index
    End If
End If

If a$ = "RR" Then 'client is returning realplayers list
    'check access:
    If CheckBit(j, 21) Then
        'is allowed
        'ensure version is correct...
        
        If Val(Replace(ConnectUsers(k).Version, ".", "")) >= 1116 Then
        
            UnPackageRealPlayers p$
            SendPacket "MS", "Real Players update recieved and installed.", Index
            AddToLogFile "UPDATE: " + Users(j).Name + " updated the real player lists."
        Else
            SendPacket "MS", "You need version 1.1.15 or HIGHER to update the real player list." + vbCrLf + "Obtain a copy from your administrator.", Index
        End If
        
    Else
        SendPacket "MS", "You are not allowed to edit real player lists.", Index
    End If
End If

If a$ = "RA" Then 'client is adding one realplayer
    'check access:
    If CheckBit(j, 21) Then
        'is allowed
        AddRealPlayer p$
        AddToLogFile "ADDREAL: " + Users(j).Name + " added a real player."
    Else
        SendPacket "MS", "You are not allowed to edit real player lists.", Index
    End If
End If

If a$ = "CA" Then 'client is adding one clanplayer
    'check access:
    If CheckBit(j, 19) Then
        'is allowed
        Num = FindPlayer(p$)
        
        AddClanPlayer2 (Num), NewScriptData
        AddToLogFile "ADDCLAN: " + Users(j).Name + " added a clan player."
    Else
        SendPacket "MS", "You are not allowed to edit clans.", Index
    End If
End If


If a$ = "WS" Then 'client is asking for web info
    'check access:
    If CheckBit(j, 22) Then
        PackageWebInfo Index
    Else
        SendPacket "MS", "You are not allowed to edit Web info.", Index
    End If
End If

If a$ = "WI" Then 'client is returning web info
    'check access:
    If CheckBit(j, 22) Then
        'is allowed
        UnPackageWebInfo p$
        SendPacket "MS", "Web Info update recieved and installed.", Index
        AddToLogFile "UPDATE: " + Users(j).Name + " updated the web info."
    Else
        SendPacket "MS", "You are not allowed to edit Web Info", Index
    End If
End If

If a$ = "WL" Then 'client is asking for web colors
    'check access:
    If CheckBit(j, 22) Then
        PackageWebColors Index
    Else
        SendPacket "MS", "You are not allowed to edit Web Colors.", Index
    End If
End If

If a$ = "WC" Then 'client is returning web info
    'check access:
    If CheckBit(j, 22) Then
        'is allowed
        UnPackageWebColors p$
        SendPacket "MS", "Web Color update recieved and installed.", Index
        AddToLogFile "UPDATE: " + Users(j).Name + " updated the web colors."
    Else
        SendPacket "MS", "You are not allowed to edit Web Colors", Index
    End If
End If

If a$ = "GW" Then 'client is asking for general info
    'check access:
    If CheckBit(j, 24) Then
        PackageGeneral Index
    Else
        SendPacket "MS", "You are not allowed to edit the General Info.", Index
    End If
End If

If a$ = "GI" Then 'client is returning general info
    'check access:
    If CheckBit(j, 24) Then
        'is allowed
        UnPackageGeneral p$
        SendPacket "MS", "General Info update recieved and installed.", Index
        AddToLogFile "UPDATE: " + Users(j).Name + " updated the general info."
    Else
        SendPacket "MS", "You are not allowed to edit the General Info", Index
    End If
End If

If a$ = "SS" Then 'client is starting a script
    'check access:
    If CheckBit(j, 0) Then
        'is allowed

        
        UnPackageButtonScript p$, NewScriptData
        AddToLogFile "BUTSCRIPT: " + Users(j).Name + " started a button script."
    Else
        SendPacket "MS", "You are not allowed to run Button Scripts", Index
    End If
End If

If a$ = "BS" Then 'client is asking for button scripts
    'check access:
    If CheckBit(j, 0) Then
        'is allowed
        PackageButtonScripts Index
    Else
        SendPacket "MS", "You are not allowed to run Button Scripts", Index
    End If
End If

If a$ = "CU" Then 'client is asking for connectusers list
    'check access:
    If CheckBit(j, 15) Then
        'is allowed
        PackageConnectUsers Index
    Else
        SendPacket "MS", "You are not allowed to see connected users.", Index
    End If
End If

If a$ = "VL" Then 'client is asking for current log
    'check access:
    If CheckBit(j, 8) Then
        'is allowed
        If CheckForFile(p$) Then 'is a file
            PackageCurrentLog Index, p$
        ElseIf p$ = "server" Then
            PackageCurrentLog Index, "1"
        ElseIf p$ = "local" Then
            PackageCurrentLog Index, "0"
        End If
    Else
        SendPacket "MS", "You are not allowed to see the current log.", Index
    End If
End If

If a$ = "LE" Then 'client is asking for the list of events
    'check access:
    If CheckBit(j, 7) Then
        'is allowed
        'send script list, and then events
        PackageEventScripts Index
        PackageEvents Index
    Else
        SendPacket "MS", "You are not allowed to edit Events.", Index
    End If
End If

If a$ = "AE" Then 'client is adding an event
    'check access:
    If CheckBit(j, 7) Then
        'is allowed
        UnPackageNewEvent p$
        PackageEventScripts Index
        PackageEvents Index
        AddToLogFile "ADDEVENT: " + Users(j).Name + " added an event."
    Else
        SendPacket "MS", "You are not allowed to edit Events.", Index
    End If
End If

If a$ = "DE" Then 'client is deleting an event
    'check access:
    If CheckBit(j, 7) Then
        'is allowed
            
        'find the event
        For i = 1 To NumEvents
            If Events(i).Name = p$ Then k = i: Exit For
        Next i
        
        If k > 0 Then
            RemoveEvent k
            PackageEventScripts Index
            PackageEvents Index
            SendPacket "MS", "Event Deleted!", Index
        End If
    Else
        SendPacket "MS", "You are not allowed to edit Events.", Index
    End If
End If




If a$ = "F0" Then 'client wants a directory refresh of initial dir
    'check access:
    If CheckBit(j, 27) Then
        If Val(Replace(ConnectUsers(k).Version, ".", "")) < 1125 Then
            SendPacket "MS", "You need version 1.1.25 or higher to use the File Manager!", Index
        Else
            'is allowed
            DirFullPath = Server.BothPath
            If Users(j).FTPRoot <> "" Then DirFullPath = Users(j).FTPRoot
        
            FTPDirRefresh DirFullPath, Index
        End If
    Else
        SendPacket "MS", "You are not allowed to use the File Browser to view remote directories", Index
    End If
End If
If a$ = "F1" Then 'client wants a directory refresh
    'check access:
    If CheckBit(j, 27) Then
        'is allowed
        'see if its beyond the hl dir
        If InStr(1, LCase(p$), LCase(Server.BothPath)) = 0 Then 'beyond hl dir
            If CheckBit(j, 26) = False Then
                SendPacket "MS", "You are not allowed to use the File Browser to view directories" + vbCrLf + "other than the base Half-Life directory.", Index
                Exit Sub
            End If
        End If
        
        If InStr(1, LCase(p$), LCase(Users(j).FTPRoot)) = 0 Then  'beyond hl dir
            If CheckBit(j, 26) = False Then
                SendPacket "MS", "You may not view this directory.", Index
                Exit Sub
            End If
        End If
        
        If Val(Replace(ConnectUsers(k).Version, ".", "")) < 1125 Then
            SendPacket "MS", "You need version 1.1.25 or higher to use the File Manager!", Index
        Else
            DirFullPath = p$
            FTPDirRefresh p$, Index
        End If
    Else
        SendPacket "MS", "You are not allowed to use the File Browser to view remote directories", Index
    End If
End If
If a$ = "F2" Then 'client is deleting files
    'check access:
    If CheckBit(j, 25) Then
        'is allowed
        UnPackageDirList p$
        FTPDeleteFiles Index
    Else
        SendPacket "MS", "You are not allowed to delete files on the server", Index
    End If
End If
If a$ = "F5" Then 'client is moving files locally (here)
    'check access:
    If CheckBit(j, 25) Then
        'is allowed
        UnPackageDirList p$
        FTPMoveFiles Index
    Else
        SendPacket "MS", "You are not allowed to move server files", Index
    End If
End If
If a$ = "F6" Then 'client is copying files locally (here)
    'check access:
    If CheckBit(j, 25) Then
        'is allowed
        UnPackageDirList p$
        FTPCopyFiles Index
    Else
        SendPacket "MS", "You are not allowed to copy server files", Index
    End If
End If
If a$ = "F8" Then 'client is asking for a file to be sent
    'check access:
    If CheckBit(j, 28) Or p$ = "newclient\\\" Then
        'is allowed
        If CheckForFile(p$) Or p$ = "newclient\\\" Then
            PackageFileSend Index, p$
        Else
            SendPacket "MS", "File not found!", Index
        End If
    Else
        SendPacket "MS", "You are not allowed to download files", Index
    End If
End If

If a$ = "F9" Then 'client is sending a file
    'check access:
    If CheckBit(j, 30) Then
        If Val(Replace(ConnectUsers(k).Version, ".", "")) < 1125 Then
            SendPacket "MS", "You need version 1.1.25 or higher to use the File Manager!", Index
        Else
            'is allowed
            UnPackageFilePacket p$, Index
        End If
    Else
        SendPacket "MS", "You are not allowed to upload files", Index
    End If
End If


If a$ = "F." Then 'client is asking for file send to stop
    'check access:
    If CheckBit(j, 30) Then
        'is allowed
        ConnectUsers(k).FileStop = True
    Else
        SendPacket "MS", "You are not allowed to upload files", Index
    End If
End If


If a$ = "F7" Then 'client is renaming
    'check access:
    If CheckBit(j, 29) Then
        'is allowed
        UnPackageRename p$, Index
    Else
        SendPacket "MS", "You are not allowed to rename files!", Index
    End If
End If

If a$ = "LS" Then 'Log Search
    'check access:
    If CheckBit(j, 8) Then
        
        If Val(Replace(ConnectUsers(k).Version, ".", "")) >= 1122 Then
        
            'is allowed
            UnPackageLogSearch p$, Index
        Else
            SendPacket "MS", "You need version 1.1.22 or higher to start log searches.", Index
        End If
    Else
        SendPacket "MS", "You are not allowed to search the logs.", Index
    End If
End If

If a$ = "M1" Then 'New Message
    'check access:
    If CheckBit(j, 17) Then
        'is allowed
        UnPackageNewMessage p$, Users(j).Name
        PackageMessages2 Index, Users(j).Name
    Else
        SendPacket "MS", "You are not allowed to leave messages for other users.", Index
    End If
End If

If a$ = "M6" Then 'Getting list of messages
    'check access:
    If CheckBit(j, 17) Then
        'is allowed
        PackageMessages Index, Users(j).Name
    Else
        SendPacket "MS", "You are not allowed to retrieve messages.", Index
    End If
End If

If a$ = "M." Then 'Getting list of messages 2
    'check access:
    If CheckBit(j, 17) Then
        'is allowed
        PackageMessages2 Index, Users(j).Name
    End If
End If


If a$ = "M7" Then 'Getting list of messages
    'check access:
    If CheckBit(j, 17) And CheckBit(j, 11) Then
        'is allowed
        PackageMessages Index, ""
    Else
        SendPacket "MS", "You are not allowed to other peoples messages.", Index
    End If
End If

If a$ = "M5" Then 'List of Users
    'check access:
    If CheckBit(j, 17) Then
        'is allowed
        PackageMessageUsers Index, j
        
    Else
        SendPacket "MS", "You are not allowed to leave messages for other users.", Index
    End If
End If

If a$ = "M2" Then 'delete message
    'check access:
    If CheckBit(j, 17) Then
        'is allowed
        DeleteMessage Val(p$)
    Else
        SendPacket "MS", "You are not allowed to access the message panel.", Index
    End If
End If

If a$ = "M3" Then 'delete message
    'check access:
    If CheckBit(j, 17) Then
        'is allowed
      
        
        For i = 1 To NumMessages
            If Messages(i).MsgId = Val(p$) Then
                If CheckBit2(Messages(i).Flags, 1) Then
                    Messages(i).Flags = Messages(i).Flags - 2
                End If
            End If
        Next i
        
        
        
    Else
        SendPacket "MS", "You are not allowed to access the message panel.", Index
    End If
End If

If a$ = "M4" Then 'delete message
    'check access:
    If CheckBit(j, 17) Then
        'is allowed
        
        For i = 1 To NumMessages
            If Messages(i).MsgId = Val(p$) Then
                If CheckBit2(Messages(i).Flags, 1) = False Then
                    Messages(i).Flags = Messages(i).Flags + 2
                End If
            End If
        Next i
        
    Else
        SendPacket "MS", "You are not allowed to access the message panel.", Index
    End If
End If

If a$ = "MD" Then 'client is asking for the map
    'check access:
    
    If CheckBit(j, 13) Then
        If Val(Replace(ConnectUsers(k).Version, ".", "")) >= 1162 Then
            PackageMapData Index
        Else
            SendPacket "MS", "You need version 1.1.62 or higher to see the map.", Index
        End If
    Else
        SendPacket "MS", "You are not allowed to view the map.", Index
    End If
End If

If a$ = "C1" Then 'connect packet
    'check access:
    UnPackageConnectPacket p$, k
    
    If CheckBit(j, 10) Then
        For i = 1 To 20
            If LastLines(i).Line <> "" Then SendChatPacket LastLines(i).Name, LastLines(i).Line, Index, k, LastLines(i).Team, LastLines(i).TimeSent
        Next i
    End If
    
    UpdateUsersList
End If

If a$ = "TE" Then 'client is asking for the map
    'check access:
    If CheckBit(j, 13) Then
        PackageTeleporters Index
    Else
        SendPacket "MS", "You are not allowed to view the map.", Index
    End If
End If

If a$ = "W1" Then 'client is asking for bad word lists
    'check access:
    If CheckBit(j, 31) Then
        
        If Val(Replace(ConnectUsers(k).Version, ".", "")) < 1143 Then
            SendPacket "MS", "You need version 1.1.43 or higher to edit the bad word list!", Index
        Else
            PackageSwears Index
        End If
    Else
        SendPacket "MS", "You are not allowed to view or edit the bad word list.", Index
    End If
End If

If a$ = "SW" Then 'client is giving us the bad words
    'check access:
    If CheckBit(j, 31) Then
        'is allowed
        UnPackageSwears p$
        SendPacket "MS", "Bad Word update Recieved and Installed.", Index
        AddToLogFile "UPDATE: " + Users(j).Name + " updated the Bad Word lists."
    Else
        SendPacket "MS", "You are not allowed to edit the bad word list.", Index
    End If
End If

If a$ = "BP" Then 'client is asking for the ents of a BSP file
    'check access:
    If CheckBit(j, 33) Then
        SendBSPEnts p$, Index
    Else
        SendPacket "MS", "You are not allowed to ent edit BSP files.", Index
    End If
End If

If a$ = "B1" Then 'client is returning the ents of a BSP file
    'check access:
    If CheckBit(j, 33) Then
        UnPackageBSPEnts p$, Index, j
    Else
        SendPacket "MS", "You are not allowed to ent edit BSP files.", Index
    End If
End If

If a$ = "H1" Then 'client is asking for all of the user info
    'check access:
    If CheckBit(j, 11) Then
        ConnectUsers(k).HiddenMode = True
        UpdateUsersList
    Else
        SendPacket "MS", "You are not allowed to enter hidden mode.", Index
    End If
End If

If a$ = "H2" Then 'client is asking for all of the user info
    'check access:
    If CheckBit(j, 11) Then
        ConnectUsers(k).HiddenMode = False
        UpdateUsersList
    Else
        SendPacket "MS", "You are not allowed to enter hidden mode.", Index
    End If
End If

If a$ = "O1" Then 'client is giving us ONE script
    'check access:
    If CheckBit(j, 4) Then
        'is allowed
        nm$ = UnPackageOneScripts(p$)
        SendPacket "MS", "One Script update for " + nm$ + " recieved and installed.", Index
        AddToLogFile "UPDATE: " + Users(j).Name + " updated script " + nm$ + "."
    Else
        SendPacket "MS", "You are not allowed to edit scripts.", Index
    End If
End If

If a$ = "GB" Then 'client is asking for bans
    'check access:
    If CheckBit(j, 18) Then
        GetBanList = True
        
        GetBansFromBanlist
                
        SendRCONCommand "listid"
        GetBansIndex = Index
    Else
        SendPacket "MS", "You are not allowed to view or edit the server ban list.", Index
    End If
End If

If a$ = "AC" Then 'admin chat
    UnPackageAdminChat p$, j
End If

If a$ = "SC" Then 'catchup admin chat

    For i = 1 To 20
        If LastChats(i) <> "" Then

            a$ = Chr(251)
            a$ = a$ + LastChatsTime(i) + Chr(250)
            a$ = a$ + LastChats(i) + Chr(250)
            a$ = a$ + Ts(LastChatsCol(i)) + Chr(250)
            a$ = a$ + LastChatsName(i) + Chr(250)
            a$ = a$ + Chr(251)
            
            SendPacket "AC", a$, Index
            'SendPacket "AC", "[" + LastChatsTime(I) + "] " + LastChats(I), Index
        End If
    Next i
End If

If a$ = "MP" Then 'map process
    PackageMapProcess Index
End If

If a$ = "G1" Then 'unpackage g
    UnPackageGameRequest p$, Index, Users(j).Name
End If

If a$ = "G2" Then 'game packet
    UnPackageGamePacket p$, Index, Users(j).Name
End If

If a$ = "AM" Then 'admin BMP file -- new one
    UnPackageAdminBMP p$, Index, Users(j).Name
End If

If a$ = "AR" Then 'admin BMP request
    PackageAdminBMP p$, Index
End If

If a$ = "GR" Then 'Returning the Ban List
    UnPackageBanList p$, Index
End If

' WHITEBOARD STUFF

If a$ = "NS" Then 'New Shape
    UnPackageNewShape p$, Users(j).Name
End If

If a$ = "AS" Then 'all shapes
    PackageAllShapes Index
End If

If a$ = "CB" Then 'all shapes
    ClearBoard Users(j).Name
End If

If a$ = "TC" Then 'change text
    ChangeText p$
End If

If a$ = "SM" Then 'moving
    MoveObject p$
End If

If a$ = "DS" Then 'delete shape
    DeleteObject p$
End If

'' END WHITEBOARD STUFF

If a$ = "AW" Then 'away mode
    UnPackageAwayMode p$, Index
End If

If a$ = "Z1" Then 'start server
    If CheckBit(j, 36) Then
        SupposedToBeRunning = True
        ManualStart = True
    Else
        SendPacket "MS", "You are not allowed to start or stop the HLDS server.", Index
    End If
End If

If a$ = "Z2" Then 'stop server
    If CheckBit(j, 36) Then
        SupposedToBeRunning = False
        ManualStart = True
    Else
        SendPacket "MS", "You are not allowed to start or stop the HLDS server.", Index
    End If
End If

If a$ = "BE" Then 'private beep
    BeepThisUser p$, CInt(j)
End If

If a$ = "ID" Then 'idle time
    
    ConnectUsers(k).IdleTime = CLng(p$)
    
End If


End Sub

Sub BeepThisUser(UsrName As String, j As Integer)
    
    For i = 1 To NumConnectUsers
        If LCase(ConnectUsers(i).Name) = LCase(UsrName) Then
            SendPacket "BE", Users(j).Name, ConnectUsers(i).Index
        End If
    Next i


End Sub

Sub DeleteObject(p$)
                
shpid = Val(p$)

j = 0
For i = 1 To NumShapes
    If Shapes(i).ShapeID = shpid Then j = i: Exit For
Next i

If j > 0 Then
        
    'delete
    
    For i = j To NumShapes - 1
        
        Shapes(i).ExtraData = Shapes(i + 1).ExtraData
        Shapes(i).FillColour = Shapes(i + 1).FillColour
        Shapes(i).LineColour = Shapes(i + 1).LineColour
        Shapes(i).LineWidth = Shapes(i + 1).LineWidth
        Shapes(i).ObjType = Shapes(i + 1).ObjType
        Shapes(i).Pos1X = Shapes(i + 1).Pos1X
        Shapes(i).Pos1Y = Shapes(i + 1).Pos1Y
        Shapes(i).Pos2X = Shapes(i + 1).Pos2X
        Shapes(i).Pos2Y = Shapes(i + 1).Pos2Y
        Shapes(i).ShapeID = Shapes(i + 1).ShapeID
        
    Next i
    NumShapes = NumShapes - 1
    ReDim Preserve Shapes(0 To NumShapes)
    
    For i = 1 To NumConnectUsers
        SendPacket "DS", p$, ConnectUsers(i).Index
    Next i
End If

End Sub

Sub MoveObject(p$)

'moves a shape
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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                m$ = Mid(a$, G, h - G)
                
                If j = 1 Then shpid = Val(m$)
                If j = 2 Then newx = Val(m$)
                If j = 3 Then newy = Val(m$)

            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

j = 0
For i = 1 To NumShapes
    If Shapes(i).ShapeID = shpid Then j = i: Exit For
Next i

If j > 0 Then
    Shapes(i).Pos1X = newx
    Shapes(i).Pos1Y = newy
        
    For i = 1 To NumConnectUsers
        SendPacket "SM", p$, ConnectUsers(i).Index
    Next i
End If

End Sub

Sub ChangeText(p$)

'change the contents of a textbox
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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                m$ = Mid(a$, G, h - G)
                
                If j = 1 Then shpid = Val(m$)
                If j = 2 Then newtxt$ = m$

            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0


j = 0
For i = 1 To NumShapes
    If Shapes(i).ShapeID = shpid Then j = i: Exit For
Next i

If j > 0 Then
    Shapes(i).ExtraData = newtxt$
    
    For i = 1 To NumConnectUsers
        SendPacket "TC", p$, ConnectUsers(i).Index
    Next i
End If

End Sub

Sub ClearBoard(nme$)

ReDim Shapes(0 To 0)
NumShapes = 0

For i = 1 To NumConnectUsers
    SendPacket "CB", nme$, ConnectUsers(i).Index
Next i


End Sub

Sub PackageAllShapes(Index)

a$ = ""
B$ = ""
For i = 1 To NumShapes
   
    a$ = a$ + Chr(251)
    a$ = a$ + Ts(Shapes(i).ObjType) + Chr(250)
    a$ = a$ + Ts(Shapes(i).LineColour) + Chr(250)
    a$ = a$ + Ts(Shapes(i).FillColour) + Chr(250)
    a$ = a$ + Ts(Shapes(i).LineWidth) + Chr(250)
    a$ = a$ + Ts(Shapes(i).Pos1X) + Chr(250)
    a$ = a$ + Ts(Shapes(i).Pos1Y) + Chr(250)
    a$ = a$ + Ts(Shapes(i).Pos2X) + Chr(250)
    a$ = a$ + Ts(Shapes(i).Pos2Y) + Chr(250)
    a$ = a$ + Shapes(i).ExtraData + Chr(250)
    a$ = a$ + Ts(Shapes(i).ShapeID) + Chr(250)
    a$ = a$ + Shapes(i).Creator + Chr(250)
    a$ = a$ + Chr(251)

    If Len(a$) > 1000 Then B$ = B$ & a$: a$ = ""

Next i
B$ = B$ & a$

SendPacket "AS", B$, CInt(Index)

End Sub

Sub UnPackageBanList(p$, Index)
If DebugMode Then LastCalled = "UnPackageBanList"

'  First gotta clear ALL the bans

For i = 1 To NumCurrBans

    ' tell server he's not banned anymore
    
    SendActualRcon "removeid " + CurrBans(i).UIDs
Next i

DoEvents

'now delete the files

If CheckForFile(Server.BothPath + "\banlist.cfg") Then Kill Server.BothPath + "\banlist.cfg"
If CheckForFile(Server.BothPath + "\banned.cfg") Then Kill Server.BothPath + "\banned.cfg"


'now read the data and make new files

NumCurrBans = 0


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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                m$ = Mid(a$, G, h - G)
                
                ReDim Preserve CurrBans(0 To i)
                
                If j = 1 Then CurrBans(i).BannedAt = m$
                If j = 2 Then CurrBans(i).BanTime = m$
                If j = 3 Then CurrBans(i).EntryName = m$
                If j = 4 Then CurrBans(i).IP = m$
                If j = 5 Then CurrBans(i).Map = m$
                If j = 6 Then CurrBans(i).Name = m$
                If j = 7 Then CurrBans(i).RealName = m$
                If j = 8 Then CurrBans(i).Reason = m$
                If j = 9 Then CurrBans(i).UIDs = m$
                
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

NumCurrBans = i

'ok now write to file again

h = FreeFile
Open Server.BothPath + "\banlist.cfg" For Append As h
    
    For i = 1 To NumCurrBans
    
        SendActualRcon "banid " + CurrBans(i).BanTime + " " + CurrBans(i).UIDs + " kick"
        
        If (CurrBans(i).BannedAt <> "") Then _
            Print #h, "//Date/Time/Map: " + CurrBans(i).BannedAt + " - " + CurrBans(i).Map
        
        If (CurrBans(i).Name <> "") Then _
            Print #h, "//Player Name: " + CurrBans(i).Name
        
        If (CurrBans(i).EntryName <> "") Then _
            Print #h, "//Entry Name: " + CurrBans(i).EntryName
        
        If (CurrBans(i).RealName <> "") Then _
            Print #h, "//Real Name: " + CurrBans(i).RealName
            
        If (CurrBans(i).UIDs <> "") Then _
            Print #h, "//UniqueID: " + CurrBans(i).UIDs
        
        If (CurrBans(i).Reason <> "") Then _
            Print #h, "//Reason: " + CurrBans(i).Reason
    
        If (CurrBans(i).IP <> "") Then _
            Print #h, "//IP: " + CurrBans(i).IP
        
        Print #h, "banid " + CurrBans(i).BanTime + " " + CurrBans(i).UIDs
        Print #h, ""
        
    Next i
Close h

SendPacket "MS", "Ban List recieved and updated on server!", CInt(Index)

AddToLogFile "BANLIST: " + ConnectUsers(Index).Name + " updated the ban list!"

End Sub

Sub SendChatToAll(Msg As String, chatcol, nme$, TheTime As String)


a$ = a$ + Chr(251)
a$ = a$ + TheTime + Chr(250)
a$ = a$ + Msg + Chr(250)
a$ = a$ + Ts(chatcol) + Chr(250)
a$ = a$ + nme$ + Chr(250)
a$ = a$ + Chr(251)

ExecFunctionScript "spec_adminchat", 2, nme$, Msg

For i = 1 To NumConnectUsers
    If Val(Replace(ConnectUsers(i).Version, ".", "")) < 1166 Then
        SendPacket "AC", "Admin Chat Requires Version 1.1.66 and UP!", ConnectUsers(i).Index
    Else
        SendPacket "AC", a$, ConnectUsers(i).Index
    End If
Next i

f$ = Server.BothPath & "\svrlogs\ac.log"
h = FreeFile
Open f$ For Append As h
    Print #h, "[" & Format(Now, "MM/DD/YYYY HH:MM:SS") & "] " & nme$ & ": " & Msg
Close h


For i = 2 To 20
    LastChats(i - 1) = LastChats(i)
    LastChatsCol(i - 1) = LastChatsCol(i)
    LastChatsName(i - 1) = LastChatsName(i)
    LastChatsTime(i - 1) = LastChatsTime(i)
Next i

LastChats(20) = Msg
LastChatsCol(20) = chatcol
LastChatsTime(20) = TheTime
LastChatsName(20) = nme$

End Sub

Sub SendBSPEnts(p$, Index)
If DebugMode Then LastCalled = "SendBSPEnts"

If CheckForFile(p$) = False Then
    SendPacket "MS", "File not found.", CInt(Index)
    Exit Sub
End If

SendPacket "BP", GetBSPEnts(p$), CInt(Index)

End Sub


Sub DeleteMessage(MsgId As Integer)


For i = 1 To NumMessages
    If Messages(i).MsgId = MsgId Then j = i: Exit For
Next i
           
If j = 0 Then Exit Sub

NumMessages = NumMessages - 1

For i = j To NumMessages
    Messages(i).Flags = Messages(i + 1).Flags
    Messages(i).MsgFor = Messages(i + 1).MsgFor
    Messages(i).MsgFrom = Messages(i + 1).MsgFrom
    Messages(i).MsgId = Messages(i + 1).MsgId
    Messages(i).MsgSubj = Messages(i + 1).MsgSubj
    Messages(i).MsgText = Messages(i + 1).MsgText
    Messages(i).MsgTimeSent = Messages(i + 1).MsgTimeSent
Next i

ReDim Preserve Messages(0 To NumMessages)



End Sub
Sub Chat(p$, n$)
If DebugMode Then LastCalled = "Chat"

nn$ = n$ + ": "
If n$ = "" Then nn$ = ""


ExecFunctionScript "spec_admintalk", 2, p$, n$


If DLLEnabled = False Then 'talking WITHOUT DLL functionality
    'prevent OVERTALKING by clipping at 52 characters
    If Len(nn$ + p$) > 52 Then p$ = LeftR(p$, 52 - Len(nn$))
      
    SendRCONCommand "say " + nn$ + p$, 1

Else 'talking WITH DLL functionality
    nn$ = " " + n$
    If n$ = "" Then nn$ = ""
    
    
    m$ = "<ADMIN" + nn$ + "> " + p$
    If Len(m$) > 96 Then m$ = LeftR(m$, 96)
    m$ = ReadyForDLL(m$)

    'prevent OVERTALKING by clipping at 96 characters

    If LastTalk = 0 Then LastTalk = 1
    SendRCONCommand "sa_talk" + Ts(LastTalk) + " " + m$
    LastTalk = LastTalk + 1
    If LastTalk >= 5 Then LastTalk = 1
    SendActualRcon SA_CHECK
    
End If

If nn$ <> "" Then AddToHTMLFile p$, 3, 0, 0, n$, ""
If nn$ = "" Then AddToHTMLFile p$, 4, 0, 0, "", ""

SendToWatchers "ADMIN", n$, p$, 0, 0

End Sub

Sub DoTalk(m$, Optional NLen As Integer, Optional Brod As Boolean)
    
    
    AddToHTMLFile m$, 16, 0, 0, "", ""
    
    If DLLEnabled Then
        
        If Brod = False Then If Len(m$) > 96 Then m$ = LeftR(m$, 96)
        
        mm$ = m$
        If NLen > 0 Then
            n$ = Left(m$, NLen)
            mm$ = Right(m$, Len(m$) - NLen)
        End If
        
        
        SendToWatchers "OTHER", n$, mm$, 0, 0
        
        If Brod = False Then
            m$ = ReadyForDLL(m$)
        
        
            'prevent OVERTALKING by clipping at 96 characters
        
            If LastTalk = 0 Then LastTalk = 1
            SendRCONCommand "sa_talk" + Ts(LastTalk) + " " + m$
            LastTalk = LastTalk + 1
            If LastTalk >= 5 Then LastTalk = 1
            SendActualRcon SA_CHECK
        End If
    Else
        
        If Len(m$) > 64 Then m$ = LeftR(m$, 96)
        mm$ = m$
        If NLen > 0 Then
            n$ = Left(m$, NLen)
            mm$ = Right(m$, Len(m$) - NLen)
        End If
        SendToWatchers "OTHER", n$, mm$, 0, 0
        
        If Brod = False Then SendRCONCommand "say " + m$
        
    End If
End Sub

Sub ChangeLogLevel(p$, k)
If DebugMode Then LastCalled = "ChangeLogLevel"

ConnectUsers(k).LogLevel = Val(p$)

End Sub


Sub ChangePass(p$, j, Index As Integer)
If DebugMode Then LastCalled = "ChangePass"

'change the users password

'seperate them...
e = InStr(1, p$, Chr(250))
If e < 2 Then er = 1
If e = Len(p$) Then er = 1

If er = 1 Then
    SendPacket "MS", "There was an error changing your password." + vbCrLf + "Please enter valid passwords and try again.", Index
    Exit Sub
End If

p1$ = Left(p$, e - 1)
p2$ = Right(p$, Len(p$) - e)

If LCase(p1$) <> LCase(Users(j).PassWord) Then
    SendPacket "MS", "Your old passwords do not match, please try again.", Index
    Exit Sub
End If

p2$ = LCase(p2$)

For i = 1 To NumUsers
    If LCase(Users(i).PassWord) = p2$ Then
        SendPacket "MS", "The new password is invalid. Please try a different one.", Index
        Exit Sub
    End If
Next i

Users(j).PassWord = p2$
SendPacket "MS", "Password changed! You must now reconnect to use the new password.", Index

SaveCommands

End Sub

Sub UnPackageRename(p$, Index As Integer)
If DebugMode Then LastCalled = "UnPackageRename"

'seperate them...
e = InStr(1, p$, Chr(250))
If e < 2 Then er = 1
If e = Len(p$) Then er = 1

If er = 1 Then Exit Sub

p1$ = Left(p$, e - 1)
p2$ = Right(p$, Len(p$) - e)

Dim Result As Long
Dim FileOp As SHFILEOPSTRUCT

FileOp.wFunc = FO_RENAME
FileOp.pFrom = p1$ + vbNullChar + vbNullChar
FileOp.pTo = p2$ + vbNullChar + vbNullChar
FileOp.fFlags = FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOCONFIRMMKDIR
Result = SHFileOperation(FileOp)

'done
FTPDirRefresh DirFullPath, CInt(Index)

End Sub

Sub SendUserInfo(Index As Integer)
If DebugMode Then LastCalled = "SendUserInfo"

'compiles and sends the user info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'users has:
'Users.Allowed
'Users.Flags
'Users.Name
'Users.Password

'compile it

For i = 1 To NumUsers
    a$ = a$ + Chr(251)
    a$ = a$ + Users(i).Allowed + Chr(250)
    a$ = a$ + Ts(Users(i).Flags) + Chr(250)
    a$ = a$ + Users(i).Name + Chr(250)
    a$ = a$ + Users(i).PassWord + Chr(250)
    a$ = a$ + Users(i).ICQ + Chr(250)
    a$ = a$ + Users(i).FTPRoot + Chr(250)
    a$ = a$ + Chr(251)
Next i

'all set, send it
SendPacket "EU", a$, Index

End Sub

Sub PackageScripts(Index As Integer)
If DebugMode Then LastCalled = "PackageScripts"

'compiles and sends the script info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'users has:
'Commands.Exec
'Commands.MustHave
'Commands.Name
'Commands.NumParams
'Commands.ScriptName
'Commands.NumButtons
'Commands.Buttons.ButtonName
'Commands.Buttons.ButtonText
'Commands.Buttons.OptionOff
'Commands.Buttons.OptionOn
'Commands.Buttons.Type

'compile it

B$ = ""

For i = 1 To NumCommands
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
    
    If Len(a$) > 1000 Then B$ = B$ & a$: a$ = ""
Next i

B$ = B$ & a$: a$ = ""

'all set, send it
SendPacket "SD", B$, Index

End Sub

Sub PackageMenuScripts(Index As Integer)
    If DebugMode Then LastCalled = "PackageMenuScripts"
    
    'compiles and sends the script info
    'generic format for array items:
    '(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc
    
    'compile it
    a$ = ""
    For i = 1 To NumCommands
        
        
        If Commands(i).NumButtons > 0 Then
        
            If Commands(i).Buttons(1).Type = 3 Then
                
                a$ = a$ + Chr(251)
                a$ = a$ + Commands(i).Buttons(1).ButtonText + Chr(250)
                a$ = a$ + Commands(i).Buttons(1).ButtonName + Chr(250)
                a$ = a$ + Commands(i).Name + Chr(250)
                a$ = a$ + Chr(251)
            
            End If
        End If
        
    Next i
    
    'all set, send it
    SendPacket "MM", a$, Index

End Sub


Sub PackageEvents(Index As Integer)
    If DebugMode Then LastCalled = "PackageEvents"
    
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
    
    For i = 1 To NumEvents
        a$ = a$ + Chr(251)
        a$ = a$ + Events(i).ComPara + Chr(250)
        For j = 0 To 6
            a$ = a$ + Ts(CInt(Events(i).Days(j))) + Chr(250)
        Next j
        a$ = a$ + Ts(Events(i).Every) + Chr(250)
        a$ = a$ + Ts(CDbl(Events(i).FirstCheck)) + Chr(250)
        a$ = a$ + Ts(Events(i).mde) + Chr(250)
        a$ = a$ + Events(i).ScriptName + Chr(250)
        a$ = a$ + Ts(Events(i).Times) + Chr(250)
        a$ = a$ + Ts(Events(i).WhatToDo) + Chr(250)
        a$ = a$ + Events(i).Name + Chr(250)
        a$ = a$ + Chr(251)
    Next i
    
    'all set, send it
    SendPacket "LE", a$, Index

End Sub

Sub FTPDirRefresh(Pth As String, Index As Integer)
If DebugMode Then LastCalled = "FTPDirRefresh"

'sends a directory refresh of this dir.

p$ = Pth
s$ = p$ + "\*.*"

a$ = Dir(s$, vbArchive + vbDirectory + vbHidden + vbReadOnly + vbSystem)

NumDirs = 0
Do While a$ <> ""
    
    If a$ <> "." And a$ <> ".." Then
        'add this one
        NumDirs = NumDirs + 1
        
        ReDim Preserve DirList(0 To NumDirs)
                
        DirList(NumDirs).DateTime = FileDateTime(p$ + "\" + a$)
        DirList(NumDirs).FullPath = p$ + "\" + a$
        DirList(NumDirs).Name = a$
        
        If (GetAttr(p$ + "\" + a$) And vbDirectory) = vbDirectory Then
            DirList(NumDirs).Type = 1
        Else
            DirList(NumDirs).Type = 0
            DirList(NumDirs).Size = Ts(FileLen(p$ + "\" + a$))
        End If
    End If
    a$ = Dir
Loop

'done, now package and send the data
SendPacket "F1", PackageDirList, Index

End Sub

Sub FTPDeleteFiles(Index)

'Deletes the files indicated by the DIRLIST variable

Dim Result As Long
Dim FileOp As SHFILEOPSTRUCT

For i = 1 To NumDirs
    FileOp.wFunc = FO_DELETE
    FileOp.pFrom = DirList(i).FullPath + vbNullChar + vbNullChar
    FileOp.pTo = DirList(i).FullPath + vbNullChar + vbNullChar
    FileOp.fFlags = FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOCONFIRMMKDIR
    Result = SHFileOperation(FileOp)
    DoEvents
Next i
'done

FTPDirRefresh DirFullPath, CInt(Index)

End Sub

Sub FTPCopyFiles(Index)

'Copies Files Locally

Dim Result As Long
Dim FileOp As SHFILEOPSTRUCT

If Dir(DirFullPath, vbDirectory) = "" Then Exit Sub


For i = 1 To NumDirs
    FileOp.wFunc = FO_COPY
    FileOp.pFrom = DirList(i).FullPath + vbNullChar + vbNullChar
    FileOp.pTo = DirFullPath + "\" + vbNullChar + vbNullChar
    FileOp.fFlags = FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOCONFIRMMKDIR
    Result = SHFileOperation(FileOp)
    DoEvents
Next i

'done

FTPDirRefresh DirFullPath, CInt(Index)

End Sub

Sub FTPMoveFiles(Index)

'Copies Files Locally

Dim Result As Long
Dim FileOp As SHFILEOPSTRUCT

If Dir(DirFullPath, vbDirectory) = "" Then Exit Sub


For i = 1 To NumDirs
    FileOp.wFunc = FO_MOVE
    FileOp.pFrom = DirList(i).FullPath + vbNullChar + vbNullChar
    FileOp.pTo = DirFullPath + "\" + vbNullChar + vbNullChar
    FileOp.fFlags = FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOCONFIRMMKDIR
    Result = SHFileOperation(FileOp)
    DoEvents
Next i

'done

SendPacket "F-", "", CInt(Index)
'FTPDirRefresh DirFullPath, CInt(Index)

End Sub

Sub PackageButtonScripts(Index As Integer)
If DebugMode Then LastCalled = "PackageButtonScripts"

'compiles and sends the script info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'users has:
'Commands.Exec
'Commands.MustHave
'Commands.Name
'Commands.NumParams
'Commands.ScriptName
'Commands.NumButtons
'Commands.Buttons.ButtonName
'Commands.Buttons.ButtonText
'Commands.Buttons.OptionOff
'Commands.Buttons.OptionOn
'Commands.Buttons.Type

'compile it
B$ = ""

For i = 1 To NumCommands
    a$ = a$ + Chr(251)
    a$ = a$ + "" + Chr(250)
    a$ = a$ + "" + Chr(250)
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
    If Len(a$) > 1000 Then B$ = B$ & a$: a$ = ""
Next i
B$ = B$ & a$

'all set, send it
SendPacket "BS", B$, Index

End Sub

Sub PackageMapData(Index As Integer)
If DebugMode Then LastCalled = "PackageMapData"

'compiles and sends the script info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc
'compile it

For X = 0 To 64
    a$ = a$ + Chr(251)
    For Y = 0 To 64
        a$ = a$ + Ts(MapArray(X, Y)) + Chr(250)
    Next Y
    a$ = a$ + Chr(251)
Next X

'all set, send it
SendPacket "MD", a$, Index

End Sub

Function PackageDirList()
If DebugMode Then LastCalled = "PackageDirList"


a$ = a$ + Chr(251)
a$ = a$ + DirFullPath + Chr(250)
a$ = a$ + Chr(251)

'compile it
For i = 1 To NumDirs
    
    B$ = Chr(251)
    B$ = B$ + Ts(CDbl(DirList(i).DateTime)) + Chr(250)
    B$ = B$ + DirList(i).FullPath + Chr(250)
    B$ = B$ + DirList(i).Name + Chr(250)
    B$ = B$ + DirList(i).Size + Chr(250)
    B$ = B$ + Ts(DirList(i).Type) + Chr(250)
    B$ = B$ + Chr(251)
    a$ = a$ + B$
Next i

'Return
PackageDirList = a$

End Function

Sub PackageFileSend(Index, Fle As String)
If DebugMode Then LastCalled = "PackageFileSend"

'Split this into packets of 1000 bytes

'get the file
startimer = Timer

If Fle = "newclient\\\" Then Fle = Server.BothPath + "\client.zip"

If CheckForFile(Fle) Then
    

    mn = FileLen(Fle)
    
    a$ = ""
    
    For i = 1 To NumUsers
        If ConnectUsers(i).Index = Index Then Exit For
    Next i
    
    insize = 8192
    
    ConnectUsers(i).FileStop = False
    h = FreeFile
Close h
    Fl$ = ""
    ret$ = ""
    packetnum = 0
    Open Fle For Binary As h
        Do While Not (EOF(h)) And Timer - startimer <= 3000
            ret$ = Input(insize, #h)
            packetnum = packetnum + 1
            'ret$ = Convert255(ret$)\
            
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
            
            SendPacket "F8", a$, CInt(Index)
            
            'wait a small amount of time here, if set.
            If General.MaxFileSend > 0 Then
                
                wit = (Len(a$) / 1024) / General.MaxFileSend
                
                
                
                If wit > 0 And wit < 10 Then
                    st = Timer
                    Do
                        DoEvents
                        tm1 = Timer - st
                    Loop Until tm1 > wit Or st > Timer
                End If
            
            End If
            
            
            DoEvents
                                  
            If ConnectUsers(i).FileStop = True Then Exit Do
        Loop
    Close h
End If

End Sub

Function Convert255(String1 As String) As String
    If DebugMode Then LastCalled = "Convert255"
    
    Dim String2 As String
    
    'Converts all occurances of values greater than (248) to (249)(value)
    
    String2 = String1
    strt = Timer
    
    'For I = 249 To 255
    '    String2 = Replace(String2, Chr(I), Chr(249) + Chr(I - 248))
    'Next I
    
    Dim EndString As String
    Dim EndString2 As String
    
    e = 0
    Do
        
        f = e
        e = InStr255(e + 1, String2)
        
        If e > 0 Then
        
            'add everything before this
            If e > f + 1 Then EndString2 = EndString2 + Mid(String2, f + 1, e - f - 1)
            
            'now add the special
            t$ = Mid(String2, e, 1)
            EndString2 = EndString2 + Chr(249) + Chr(Asc(t$) - 248)
        
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

Function CountOcc(String1 As String, String2 As String) As Integer
    
    Do
        e = InStr(e + 1, String1, String2)
        If e > 0 Then nm = nm + 1
    Loop Until e = 0
    
    CountOcc = nm

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

Sub UnPackageFilePacket(p$, Index)
    
    For i = 1 To NumUsers
        If ConnectUsers(i).Index = Index Then Num = i: Exit For
    Next i
    
    If Num > 0 Then
        'extracts filepacket info from the sent string
        
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
                    G = h
                    h = InStr(G + 1, a$, Chr(250))
                    G = G + 1
                    j = j + 1
                    If G > 0 And h > G Then
                        m$ = Mid(a$, G, h - G)
    '            h = 0
    '            j = 0
    '            Do
    '                G = h
    '                If j < 3 Then h = InStr(G + 1, a$, Chr(250))
    '                If j >= 3 Then h = InStrRev(a$, Chr(250))
    '                G = G + 1
    '                j = j + 1
    '                If G > 0 And h > G - 1 Then
    '                    m$ = Mid(a$, G, h - G)
    
                        If j = 1 Then packetnum = Val(m$)
                        If j = 2 Then ConnectUsers(Num).FileSavePath = m$
                        If j = 3 Then ConnectUsers(Num).FileSize = Val(m$)
                        If j = 4 Then
                            m$ = DeCode255(m$)
                            
                            
                            If packetnum = 1 Then
                                Close ConnectUsers(Num).FileNum
                                ConnectUsers(Num).FileNum = FreeFile
                                If CheckForFile(ConnectUsers(Num).FileSavePath) Then Kill ConnectUsers(Num).FileSavePath
                                Open ConnectUsers(Num).FileSavePath For Binary As ConnectUsers(Num).FileNum
                                ConnectUsers(Num).BytesTransferred = 0
                            End If
                            
                            Put #ConnectUsers(Num).FileNum, ConnectUsers(Num).BytesTransferred + 1, m$
                            ConnectUsers(Num).BytesTransferred = ConnectUsers(Num).BytesTransferred + Len(m$)
                        End If
    
                        If j = 5 Then filestatus = Val(m$)
                    End If
                Loop Until h = 0
            
            End If
        Loop Until f = 0 Or e = 0
        
        'If ConnectUsers(Num).BytesTransferred >= ConnectUsers(Num).FileSize Then FileDone Num
        If filestatus = 1 Then FileDone Num, Index
        
        'send to the guy how much is done...
        'SendPacket "FS", Ts(ConnectUsers(Num).BytesTransferred), CInt(Index)
        
    End If

End Sub

Sub FileDone(Num, Index)
    
    'called when a file transfer is complete. Write the data to the required file.
    
    Fl$ = ConnectUsers(Num).FileSavePath
    
    Close ConnectUsers(Num).FileNum
    ConnectUsers(Num).FileNum = 0
    
    ConnectUsers(Num).FileSize = 0
    
    SendPacket "F-", "", CInt(Index)
    
    SendPacket "MS", "Your file has finished uploading:" + vbCrLf + ConnectUsers(Num).FileSavePath, CInt(Index)
    
    ConnectUsers(Num).FileSavePath = ""

End Sub


Sub UnPackageDirList(p$)
    If DebugMode Then LastCalled = "UnPackageDirList"
    
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
                G = h
                h = InStr(G + 1, a$, Chr(250))
                G = G + 1
                j = j + 1
                If G > 0 And h > G - 1 Then
                    m$ = Mid(a$, G, h - G)
                    If i = 1 Then
                        'DirFullpath
                        DirFullPath = m$
                    Else
                        ii = i - 1
    
                        ReDim Preserve DirList(0 To ii)
                        
                        If j = 1 Then DirList(ii).DateTime = CDate(m$)
                        If j = 2 Then DirList(ii).FullPath = m$
                        If j = 3 Then DirList(ii).Name = m$
                        If j = 4 Then DirList(ii).Size = m$
                        If j = 5 Then DirList(ii).Type = Val(m$)
                    End If
                End If
            Loop Until h = 0
        
        End If
    Loop Until f = 0 Or e = 0
    
    NumDirs = i - 1

End Sub

Sub UnPackageNewShape(p$, nme$)
    If DebugMode Then LastCalled = "UnPackageNewShape"
    ' creates a new shape and sends to everyone
    
    Dim NewShape As typWhiteboard
    
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
                G = h
                h = InStr(G + 1, a$, Chr(250))
                G = G + 1
                j = j + 1
                If G > 0 And h > G - 1 Then
                    m$ = Mid(a$, G, h - G)
                    
                    If j = 1 Then NewShape.ObjType = Val(m$)
                    If j = 2 Then NewShape.LineColour = Val(m$)
                    If j = 3 Then NewShape.FillColour = Val(m$)
                    If j = 4 Then NewShape.LineWidth = Val(m$)
                    If j = 5 Then NewShape.Pos1X = Val(m$)
                    If j = 6 Then NewShape.Pos1Y = Val(m$)
                    If j = 7 Then NewShape.Pos2X = Val(m$)
                    If j = 8 Then NewShape.Pos2Y = Val(m$)
                    
                    ' Dont DECODE255 this... the clients will do that! Keep it in 255Code format.
                    If j = 9 Then NewShape.ExtraData = m$
                    
                    If j = 10 Then NewShape.ShapeID = Val(m$)
                    
                End If
            Loop Until h = 0
        
        End If
    Loop Until f = 0 Or e = 0
    
    'now add this shape to the whiteboard
    
    NewShape.Creator = nme$
    
    NumShapes = NumShapes + 1
    
    ReDim Preserve Shapes(0 To NumShapes)
    
    Shapes(NumShapes).ExtraData = NewShape.ExtraData
    Shapes(NumShapes).FillColour = NewShape.FillColour
    Shapes(NumShapes).LineColour = NewShape.LineColour
    Shapes(NumShapes).LineWidth = NewShape.LineWidth
    Shapes(NumShapes).ObjType = NewShape.ObjType
    Shapes(NumShapes).Pos1X = NewShape.Pos1X
    Shapes(NumShapes).Pos1Y = NewShape.Pos1Y
    Shapes(NumShapes).Pos2X = NewShape.Pos2X
    Shapes(NumShapes).Pos2Y = NewShape.Pos2Y
    Shapes(NumShapes).ShapeID = NewShape.ShapeID
    Shapes(NumShapes).Creator = NewShape.Creator
    
    'now, send the new shape thingy to all connected clients
    'no need to repackage, we just send the same thing back
    'just add who its from
    
    p$ = Left(p$, Len(p$) - 1)
    p$ = p$ + nme$ + Chr(250) + Chr(251)
    
    For i = 1 To NumConnectUsers
        SendPacket "NS", p$, ConnectUsers(i).Index
    Next i

End Sub

Sub PackageEventScripts(Index As Integer)
If DebugMode Then LastCalled = "PackageEventScripts"

'compiles and sends the script info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'users has:
'Commands.Exec
'Commands.MustHave
'Commands.Name
'Commands.NumParams
'Commands.ScriptName
'Commands.NumButtons
'Commands.Buttons.ButtonName
'Commands.Buttons.ButtonText
'Commands.Buttons.OptionOff
'Commands.Buttons.OptionOn
'Commands.Buttons.Type

'compile it

For i = 1 To NumCommands
    a$ = a$ + Chr(251)
    a$ = a$ + "" + Chr(250)
    a$ = a$ + "" + Chr(250)
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
    If Len(a$) > 1000 Then B$ = B$ & a$: a$ = ""
Next i

B$ = B$ & a$: a$ = ""

'all set, send it
SendPacket "L2", B$, Index

End Sub

Sub PackageClans(Index As Integer)
If DebugMode Then LastCalled = "PackageClans"

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
SendPacket "CM", a$, Index

End Sub


Sub PackageAdminBMPList(Index As Integer)
If DebugMode Then LastCalled = "PackageAdminBMPList"

For i = 1 To NumAdminBMP
    a$ = a$ + Chr(251)
    a$ = a$ + AdminBMP(i).AdminName + Chr(250)
    a$ = a$ + AdminBMP(i).BMPFile + Chr(250)
    a$ = a$ + Chr(251)
Next i

'all set, send it
SendPacket "AL", a$, Index

End Sub


Sub PackageCurrentLog(Index As Integer, Fle As String)
If DebugMode Then LastCalled = "PackageCurrentLog"

'sends the current log file :)

If Fle = "0" Then
    a$ = Server.BothPath + "\svrlogs\log-" + Date$ + ".log"
    
    If CheckForFile(a$) = False Then
        SendPacket "MS", "Current log file not found.", Index
        Exit Sub
    End If
    
    h = FreeFile
Close h

    c$ = Space(FileLen(a$))
    
    Open a$ For Binary As h
        Get h, , c$
    Close h
    
ElseIf Fle = "1" Then
    'the current SERVER log
    
    a$ = Server.BothPath + "\logs\" + GetLastLog
    
    If CheckForFile(a$) = False Then
        SendPacket "MS", "Current log file not found.", Index
        Exit Sub
    End If
    
    h = FreeFile
Close h

    c$ = Space(FileLen(a$))
    
    Open a$ For Binary As h
        Get h, , c$
    Close h
    
    'remove RCON passwords
    c$ = Replace(c$, Server.RCONPass, "*****")
        
    'remove special characters
    
Else
    
    a$ = Fle
    
    If CheckForFile(a$) = False Then
        SendPacket "MS", "Log file " + Fle + " not found.", Index
        Exit Sub
    End If
    
    h = FreeFile
Close h

    c$ = Space(FileLen(a$))
    
    Open a$ For Binary As h
        Get h, , c$
    Close h
    
    'remove RCON passwords
    c$ = Replace(c$, Server.RCONPass, "****")
       

End If

'remove special characters
c$ = Replace(c$, Chr(250), " ")
c$ = Replace(c$, Chr(251), " ")
c$ = Replace(c$, Chr(252), " ")
c$ = Replace(c$, Chr(253), " ")
c$ = Replace(c$, Chr(254), " ")
c$ = Replace(c$, Chr(255), " ")
c$ = Replace(c$, Chr(0), " ")
    
'all set, send it
SendPacket "VL", c$, Index

End Sub


Sub PackageSpeech(Index As Integer)
If DebugMode Then LastCalled = "PackageSpeech"

'compiles and sends the Speech info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'Speech has:

'Speech.ClientText
'Speech.NumAnswers
'speech.Answers()

'compile it

For i = 1 To NumSpeech
    
    If LeftR(Speech(i).ClientText, 5) = "ADMIN" And Len(Speech(i).ClientText) > 5 Then Speech(i).ClientText = Trim(Right(Speech(i).ClientText, Len(Speech(i).ClientText) - 5))
    
    a$ = a$ + Chr(251)
    a$ = a$ + Speech(i).ClientText + Chr(250)
    a$ = a$ + Ts(Speech(i).NumAnswers) + Chr(250)
    For j = 1 To Speech(i).NumAnswers
        a$ = a$ + Speech(i).Answers(j) + Chr(250)
    Next j
    a$ = a$ + Chr(251)
Next i

'all set, send it
SendPacket "SP", a$, Index

End Sub

Sub UnPackageSpeech(p$)
If DebugMode Then LastCalled = "UnPackageSpeech"

'extracts speech from the sent string

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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                m$ = Mid(a$, G, h - G)
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

SaveCommands


End Sub

Sub UnPackageRealPlayers(p$)
    If DebugMode Then LastCalled = "UnPackageRealPlayers"
    
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
                G = h
                h = InStr(G + 1, a$, Chr(250))
                G = G + 1
                j = j + 1
                If G > 0 And h > G - 1 Then
                    m$ = Mid(a$, G, h - G)
                    ReDim Preserve RealPlayers(0 To i)
                    
                    If j = 1 Then RealPlayers(i).Flags = Val(m$)
                    If j = 2 Then RealPlayers(i).RealName = m$
                    If j = 3 Then RealPlayers(i).UniqueID = m$
                    If j = 4 Then RealPlayers(i).LastTime = Val(m$)
                    If j = 5 Then RealPlayers(i).LastName = m$
                    If j = 6 Then
                        '10000000
                        If Len(m$) > 8 Then m$ = "0"
                        RealPlayers(i).Points = m$
                    End If
                    If j = 7 Then RealPlayers(i).TimesSeen = m$
                End If
            Loop Until h = 0
        
        End If
    Loop Until f = 0 Or e = 0
    
    NumRealPlayers = i
    DoPlayerScan
    
    SaveCommands

End Sub

Sub UnPackageSwears(p$)
If DebugMode Then LastCalled = "UnPackageSwears"

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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                m$ = Mid(a$, G, h - G)
                ReDim Preserve Swears(0 To i)
                
                If j = 1 Then Swears(i).BadWord = m$
                If j = 2 Then Swears(i).Flags = Val(m$)
                If j = 3 Then Swears(i).Replacement = m$
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

NumSwears = i
DoPlayerScan

SaveCommands

MakeSwearList

SendRCONCommand "sa_reloadcensorfile 1"


End Sub

Sub MakeSwearList()


'creates CENSOR.TXT


a$ = Server.BothPath + "\censor.lst"

If CheckForFile(a$) Then Kill a$

'sort based on length
For i = 1 To NumSwears
    dne = 0
    For j = 1 To NumSwears - 1
        
        If Len(Swears(j + 1).BadWord) > Len(Swears(j).BadWord) Then
            'swap
            Swap Swears(j + 1).BadWord, Swears(j).BadWord
            Swap Swears(j + 1).Flags, Swears(j).Flags
            Swap Swears(j + 1).Replacement, Swears(j).Replacement
            dne = 1
        End If
    Next j
    If dne = 0 Then Exit For
Next i

h = FreeFile
Close h
Open a$ For Append As h

    Print #h, "!        PROFESSIONAL VICTIM'S UTILITY DLL"
    Print #h, "!            CENSOR WORDS LIST"
    Print #h, "!         DO NOT MODIFY THIS FILE!"
    Print #h, ""
    
    For i = 1 To NumSwears
    
        If CheckBit2(Swears(i).Flags, 7) Then
            Print #h, Swears(i).BadWord + ":" + Swears(i).Replacement
        End If
    Next i
    
Close h




End Sub

Sub Swap(a As Variant, B As Variant)
Dim c As Variant

c = a
a = B
B = c

End Sub

Sub UnPackageButtonScript(p$, scriptdata As typScriptData)
If DebugMode Then LastCalled = "UnPackageButtonScript"
'extracts script param from sent string

Dim UserParms() As typParams
ReDim UserParms(1 To 200)

f = 0
i = 0
k = 0
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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                m$ = Mid(a$, G, h - G)
                
                If j = 1 Then cmd$ = m$
                If j > 1 Then
                    UserParms(j - 1).Value = m$
                    k = k + 1
                End If
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

bbc$ = ExecuteScript(cmd$, UserParms, CInt(k), nnm$, scriptdata)
'done

End Sub



Sub UnPackageClans(p$)
If DebugMode Then LastCalled = "UnPackageClans"
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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                m$ = Mid(a$, G, h - G)
                
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

SaveCommands

End Sub

Sub PackageServerInfo(Index As Integer)
If DebugMode Then LastCalled = "PackageServerInfo"

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

a$ = a$ + Ts(CInt(ServerStart.AutoRestart)) + Chr(250)
a$ = a$ + Ts(CInt(ServerStart.UseFeature)) + Chr(250)
a$ = a$ + ServerStart.CommandLine + Chr(250)
a$ = a$ + ServerStart.HLDSDir + Chr(250)
a$ = a$ + ServerStart.HLDSPath + Chr(250)

a$ = a$ + Chr(251)

'all set, send it
SendPacket "SI", a$, Index

End Sub

Sub SendUpdate()
If DebugMode Then LastCalled = "SendUpdate"

'compiles and sends an update

'compile it
For i = 1 To NumPlayers
    If Players(i).ConnectOnly = False Then e = e + 1
Next i

c$ = Ts(e) + " playing, " + Ts(NumPlayers) + " connected, " + Ts(Vars.MaxPlayers) + " max"

If LastCrashCall = 3 Then c$ = "SERVER CRASHED"

a$ = a$ + Chr(251)
a$ = a$ + Ts(Vars.MapTimeLeft) + Chr(250)
a$ = a$ + Vars.Map + Chr(250)
a$ = a$ + c$ + Chr(250)
a$ = a$ + TeamNames(1) + Chr(250)
a$ = a$ + TeamNames(2) + Chr(250)
a$ = a$ + TeamNames(3) + Chr(250)
a$ = a$ + TeamNames(4) + Chr(250)
a$ = a$ + Chr(251)

'all set, send it
For i = 1 To NumConnectUsers
    SendPacket "UP", a$, ConnectUsers(i).Index
Next i
End Sub

Sub PackageGeneral(Index As Integer)
If DebugMode Then LastCalled = "PackageGeneral"

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
SendPacket "GI", a$, Index

End Sub

Sub PackageConnectPacket(Index As Integer, j As Boolean)
If DebugMode Then LastCalled = "PackageConnectPacket"

'compiles and sends the GENERAL info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'compile it

a$ = a$ + Chr(251)
a$ = a$ + Ts(CInt(DLLEnabled)) + Chr(250)
a$ = a$ + Ts(Server.GameMode) + Chr(250)
a$ = a$ + Ts(App.Major) + "." + Ts(App.Minor) + "." + Ts(App.Revision) + Chr(250)
a$ = a$ + Ts(CInt(j)) + Chr(250)
a$ = a$ + Server.LocalIP + Chr(250)
a$ = a$ + Server.ServerPort + Chr(250)
a$ = a$ + Server.GamePath + Chr(250)
a$ = a$ + General.CustomFlag1 + Chr(250)
a$ = a$ + General.CustomFlag2 + Chr(250)
a$ = a$ + General.CustomFlag3 + Chr(250)
a$ = a$ + General.CustomFlag4 + Chr(250)
a$ = a$ + TeamNames(1) + Chr(250)
a$ = a$ + TeamNames(2) + Chr(250)
a$ = a$ + TeamNames(3) + Chr(250)
a$ = a$ + TeamNames(4) + Chr(250)
a$ = a$ + General.NewestClient + Chr(250)

a$ = a$ + Chr(251)

'all set, send it
SendPacket "C1", a$, Index

End Sub

Sub UnPackageConnectPacket(p$, jjj)
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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                
                m$ = Mid(a$, G, h - G)
                
                If j = 1 Then ConnectUsers(jjj).Version = m$
                If j = 2 Then ConnectUsers(jjj).HiddenMode = CBool(m$)

            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0
UpdateUsersList

End Sub

Sub UnPackageGeneral(p$)
    If DebugMode Then LastCalled = "UnPackageGeneral"
    
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
                G = h
                h = InStr(G + 1, a$, Chr(250))
                
                G = G + 1
                j = j + 1
                If G > 0 And h > G - 1 Then
                    
                    m$ = Mid(a$, G, h - G)
                    
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
    
    SaveCommands
    
    For i = 1 To NumConnectUsers
        PackageConnectPacket ConnectUsers(i).Index, CheckBit(ConnectUsers(i).UserNum, 11)
    Next i

End Sub

Sub UnPackageAdminBMP(p$, Index As Integer, AdminName As String)
If DebugMode Then LastCalled = "UnPackageAdminBMP"

'extracts the BMP file

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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                
                m$ = Mid(a$, G, h - G)
                
                If j = 1 Then bmp$ = m$
                
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

bmp$ = DeCode255(bmp$)

'got the file

'determine a random number for this admin
Randomize

n = Int(Rnd * 30000) + 1

'make the filename

fn$ = AdminName + Ts(n) + ".bmp"

'write the file

If CheckForFile(Server.BothPath + "\Assistant\Data\" + fn$) Then Kill Server.BothPath + "\Assistant\Data\" + fn$

h = FreeFile
Open Server.BothPath + "\Assistant\Data\" + fn$ For Binary As h
    Put #h, , bmp$
Close h

'update the list
j = 0
For i = 1 To NumAdminBMP
    If AdminBMP(i).AdminName = AdminName Then j = i: Exit For
Next i

' tell client to delete old file
If j <> 0 Then


    For i = 1 To NumConnectUsers
        SendPacket "AD", AdminBMP(j).BMPFile, ConnectUsers(i).Index
    Next i
    ' and also delete MY old copy.
    
    If CheckForFile(Server.BothPath + "\Assistant\Data\" + AdminBMP(j).BMPFile) Then Kill Server.BothPath + "\Assistant\Data\" + AdminBMP(j).BMPFile
End If

If j = 0 Then
    NumAdminBMP = NumAdminBMP + 1
    ReDim Preserve AdminBMP(0 To NumAdminBMP)
    j = NumAdminBMP
End If

AdminBMP(j).AdminName = AdminName
AdminBMP(j).BMPFile = fn$

'finally, tell everyone that there is a new file

For i = 1 To NumConnectUsers
   PackageAdminBMPList ConnectUsers(i).Index
Next i

SaveCommands

End Sub


Sub PackageAdminBMP(AdminName As String, Index As Integer)

'Send
'find the admin

For i = 1 To NumAdminBMP
    If AdminBMP(i).AdminName = AdminName Then j = i: Exit For
Next i

Dim LocFle As String

If j > 0 Then

    LocFle = Server.BothPath + "\Assistant\Data\" + AdminBMP(j).BMPFile
    
    If CheckForFile(LocFle) Then
    
        mn = FileLen(LocFle)
        a$ = ""
        h = FreeFile
        
        ret$ = Space(mn)
        Open LocFle For Binary As h
            
            ' read the file
            Get h, , ret$
        Close h
    
        ret$ = Convert255(ret$)
        
        a$ = Chr(251)
        a$ = a$ + AdminBMP(j).AdminName + Chr(250)
        a$ = a$ + AdminBMP(j).BMPFile + Chr(250)
        a$ = a$ + ret$ + Chr(250)
        a$ = a$ + Chr(251)
        
        ' send the BMP to the client
        SendPacket "AM", a$, Index
    
    End If
    
End If

End Sub

Sub UnPackageNewEvent(p$)
If DebugMode Then LastCalled = "UnPackageNewEvent"

'extracts a new event from the sent string
Dim NewEvent As typEvent

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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                
                m$ = Mid(a$, G, h - G)
                
                If j = 1 Then NewEvent.ComPara = m$
                If j >= 2 And j <= 8 Then NewEvent.Days(j - 2) = CBool(m$)
                If j = 9 Then NewEvent.Every = Val(m$)
                If j = 10 Then NewEvent.FirstCheck = CDate(m$)
                If j = 11 Then NewEvent.mde = Val(m$)
                If j = 12 Then NewEvent.ScriptName = m$
                If j = 13 Then NewEvent.Times = Val(m$)
                If j = 14 Then NewEvent.WhatToDo = Val(m$)
                If j = 15 Then NewEvent.Name = m$
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

'now lets check to make sure everything is in order...
again:
For i = 1 To NumEvents
    If Events(i).Name = NewEvent.Name Then 'uh oh...
        RemoveEvent i
        GoTo again
    End If
Next i

'ok, thats all set...
'now add new one

NumEvents = NumEvents + 1
ReDim Preserve Events(0 To NumEvents)

Events(NumEvents).ComPara = NewEvent.ComPara
For j = 0 To 6
    Events(NumEvents).Days(j) = NewEvent.Days(j)
Next j
Events(NumEvents).Every = NewEvent.Every
Events(NumEvents).FirstCheck = NewEvent.FirstCheck
Events(NumEvents).mde = NewEvent.mde
Events(NumEvents).Name = NewEvent.Name
Events(NumEvents).ScriptName = NewEvent.ScriptName
Events(NumEvents).Times = NewEvent.Times
Events(NumEvents).WhatToDo = NewEvent.WhatToDo
'all done, now SAVE!
SaveCommands

End Sub

Sub PackageWebInfo(Index As Integer)
If DebugMode Then LastCalled = "PackageWebInfo"

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
SendPacket "WI", a$, Index

End Sub

Sub UnPackageWebColors(p$)
If DebugMode Then LastCalled = "UnPackageWebColors"

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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                
                m$ = Mid(a$, G, h - G)
                
                If j = 1 Then Web.Colors(i).r = Val(m$)
                If j = 2 Then Web.Colors(i).G = Val(m$)
                If j = 3 Then Web.Colors(i).B = Val(m$)
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

SaveCommands

End Sub


Sub UnPackageWebInfo(p$)
If DebugMode Then LastCalled = "UnPackageWebInfo"

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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                
                m$ = Mid(a$, G, h - G)
                
                If j = 1 Then Web.Enabled = CBool(m$)
                If j = 2 Then Web.LogFlags = Val(m$)
                If j = 3 Then Web.LogPath = m$
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

SaveCommands

End Sub


Sub UnPackageServerInfo(p$)
If DebugMode Then LastCalled = "UnPackageServerInfo"

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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                
                m$ = Mid(a$, G, h - G)
                
                If j = 1 Then Server.HLPath = m$
                If j = 2 Then Server.GamePath = m$
                If j = 3 Then Server.ServerPort = m$
                If j = 4 Then Server.RconListenPort = m$
                If j = 5 Then Server.LocalConnectPort = m$
                If j = 6 Then Server.LocalIP = m$
            
                If j = 7 Then ServerStart.AutoRestart = CBool(m$)
                If j = 8 Then ServerStart.UseFeature = CBool(m$)
                If j = 9 Then ServerStart.CommandLine = m$
                If j = 10 Then ServerStart.HLDSDir = m$
                If j = 11 Then ServerStart.HLDSPath = m$
            
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

Server.BothPath = Server.HLPath + "\" + Server.GamePath

SaveCommands

End Sub

Sub UnPackageLogSearch(p$, Index)
If DebugMode Then LastCalled = "UnPackageLogSearch"

Dim FromDay As Date
Dim ToDay As Date
Dim Text As String
Dim SearchSubs As Integer, SearchPath As String, ExactPhrase As Boolean, AllWords As Integer, SearchSaysOnly As Integer


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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                
                m$ = Mid(a$, G, h - G)
                
                If j = 1 Then Text = m$
                If j = 2 Then chck = Val(m$)
                If j = 3 Then FromDay = CDate(m$)
                If j = 4 Then ToDay = CDate(m$)
                If j = 5 Then SearchSubs = Val(m$)
                If j = 6 Then SearchPath = m$
                If j = 7 Then ExactPhrase = CBool(m$)
                If j = 8 Then AllWords = Val(m$)
                If j = 9 Then SearchSaysOnly = Val(m$)
            
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0


LogSearch Text, CInt(chck), FromDay, ToDay, CInt(Index), SearchSubs, SearchPath, ExactPhrase, AllWords, SearchSaysOnly

End Sub


Sub PackagePlayers(Index As Integer)
If DebugMode Then LastCalled = "PackagePlayers"


'compiles and sends the player info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'players has:

'Players.Class
'Players.IP
'Players.Name
'Players.Team
'Players.UniqueID
'Players.UserID

'compile it

For i = 1 To NumPlayers
        a$ = a$ + Chr(251)
        a$ = a$ + Ts(Players(i).Class) + Chr(250)
        a$ = a$ + Players(i).IP + Chr(250)
        a$ = a$ + Players(i).Name + Chr(250)
        a$ = a$ + Ts(Players(i).Team) + Chr(250)
        a$ = a$ + Players(i).UniqueID + Chr(250)
        a$ = a$ + Ts(Players(i).UserID) + Chr(250)
        a$ = a$ + Ts(CInt(Players(i).ConnectOnly)) + Chr(250)
        rn$ = Players(i).RealName
        
        If rn$ <> "" And Players(i).TempRealMode = True Then rn$ = "++ " + rn$
        If rn$ = "" Then rn$ = "--  " + Players(i).EntryName
        
        a$ = a$ + rn$ + Chr(250)
        a$ = a$ + Ts(Players(i).Pos.X) + Chr(250)
        a$ = a$ + Ts(Players(i).Pos.Y) + Chr(250)
        a$ = a$ + Ts(Players(i).Pos.Z) + Chr(250)
        a$ = a$ + Ts(CDbl(Now - Players(i).TimeJoined)) + Chr(250)
        a$ = a$ + Players(i).EntryName + Chr(250)
        a$ = a$ + Ts(Players(i).NumKickVotes) + Chr(250)
        a$ = a$ + Ts(Players(i).Port) + Chr(250)
        a$ = a$ + Ts(Players(i).ShutUp) + Chr(250)
        a$ = a$ + Ts(Players(i).Warn) + Chr(250)
        If Players(i).Points = 0 Then Players(i).Points = GetPoints(i)
        a$ = a$ + Ts(Players(i).Points) + Chr(250)
        a$ = a$ + Ts(CDbl(Now - Players(i).LastEvent)) + Chr(250)
        
        a$ = a$ + Chr(251)
Next i

'all set, send it
SendPacket "SU", a$, Index

End Sub

Sub RemoveOldMessages()

    'Removes messages more than 2 weeks old.
    
    
    strt = 1
again:
    
    For i = strt To NumMessages
        If GetDays(Messages(i).MsgTimeSent) > 14 Then
            DeleteMessage Messages(i).MsgId
            strt = i
            GoTo again
        End If
    Next i
    
    
    

End Sub

Function GetDays(Tme As Date)

GetDays = Int(CDbl(Now - Tme))

End Function

Function GetSec(Tme As Date)

e = CDbl(Now - Tme)
f = 1 / 24 / 60 / 60

mn = Round(e / f)

GetSec = mn

End Function


Sub PackageMessageUsers(Index As Integer, j)
If DebugMode Then LastCalled = "PackageMessageUsers"

'compile it

For i = 1 To NumUsers
    If CheckBit(i, 17) Then
        a$ = a$ + Chr(251)
        a$ = a$ + Users(i).Name + Chr(250)
        a$ = a$ + Chr(251)
    End If
Next i

'all set, send it
SendPacket "M3", a$, Index

End Sub

Sub PackageWebColors(Index As Integer)
If DebugMode Then LastCalled = "PackageWebColors"

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
        a$ = a$ + Ts(Web.Colors(i).G) + Chr(250)
        a$ = a$ + Ts(Web.Colors(i).B) + Chr(250)
        a$ = a$ + Chr(251)
Next i

'all set, send it
SendPacket "WC", a$, Index

End Sub

Sub PackageRealPlayers(Index As Integer)
    If DebugMode Then LastCalled = "PackageRealPlayers"
    
    'compiles and sends the real player info
    'generic format for array items:
    '(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc
    
    'realplayers has:
    
    'RealPlayers.LastName
    'RealPlayers.RealName
    'RealPlayers.UniqueID
    
    'compile it
    B$ = ""
    For i = 1 To NumRealPlayers
        a$ = a$ + Chr(251)
        a$ = a$ + RealPlayers(i).LastName + Chr(250)
        a$ = a$ + RealPlayers(i).RealName + Chr(250)
        a$ = a$ + RealPlayers(i).UniqueID + Chr(250)
        a$ = a$ + Ts(RealPlayers(i).LastTime) + Chr(250)
        a$ = a$ + Ts(RealPlayers(i).Flags) + Chr(250)
        a$ = a$ + RealPlayers(i).Points + Chr(250)
        a$ = a$ + RealPlayers(i).TimesSeen + Chr(250)
        a$ = a$ + Chr(251)
        If Len(a$) > 1000 Then B$ = B$ & a$: a$ = ""
        
    Next i
    
    B$ = B$ & a$
    
    'all set, send it
    SendPacket "RP", B$, Index

End Sub

Sub PackageSwears(Index As Integer)
If DebugMode Then LastCalled = "PackageSwears"

'compile it

For i = 1 To NumSwears
    a$ = a$ + Chr(251)
    a$ = a$ + Swears(i).BadWord + Chr(250)
    a$ = a$ + Ts(Swears(i).Flags) + Chr(250)
    a$ = a$ + Swears(i).Replacement + Chr(250)
    a$ = a$ + Chr(251)
Next i

'all set, send it
SendPacket "SW", a$, Index

End Sub

Sub PackageKickBans(Index As Integer)
If DebugMode Then LastCalled = "PackageKickBans"

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
SendPacket "KB", a$, Index

End Sub

Sub PackageLogSearch(Index As Integer)
If DebugMode Then LastCalled = "PackageLogSearch"

'compiles and sends the kickban info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc


'compile it
B$ = ""
For i = 1 To NumLogFound
    a$ = a$ + Chr(251)
    a$ = a$ + LogFound(i).LogFile + Chr(250)
    a$ = a$ + LogFound(i).LogLine + Chr(250)
    a$ = a$ + Chr(251)
    
    If Len(a$) > 500 Then
        B$ = B$ + a$
        a$ = ""
    End If
    
Next i
B$ = B$ + a$

'all set, send it
SendPacket "LS", B$, Index

End Sub

Sub PackageConnectUsers(Index As Integer)
If DebugMode Then LastCalled = "PackageConnectUsers"

'compiles and sends the connected user lists
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'players has:

'ConnectUsers.IP
'ConnectUsers.Name

'compile it

For i = 1 To NumConnectUsers
    
    If ConnectUsers(i).HiddenMode = False Then
        a$ = a$ + Chr(251)
        a$ = a$ + ConnectUsers(i).IP + ":" + Ts(ConnectUsers(i).Port) + Chr(250)
        a$ = a$ + ConnectUsers(i).Name + Chr(250)
        a$ = a$ + ConnectUsers(i).Version + Chr(250)
        a$ = a$ + Ts(ConnectUsers(i).AwayMode) + Chr(250)
        a$ = a$ + ConnectUsers(i).AwayMsg + Chr(250)
        a$ = a$ + Ts(ConnectUsers(i).IdleTime) + Chr(250)
        a$ = a$ + Chr(251)
    End If
    
Next i

'all set, send it
SendPacket "CU", a$, Index

End Sub

Sub PackageMapProcess(Index As Integer)
If DebugMode Then LastCalled = "PackageMapProcess"

'compiles and sends the connected user lists
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'players has:

'ConnectUsers.IP
'ConnectUsers.Name

'compile it

For i = 1 To NumMapProcess
    
    a$ = a$ + Chr(251)
    a$ = a$ + Ts(CDbl(MapProcess(i).LastTimePlayed)) + Chr(250)
    a$ = a$ + MapProcess(i).MapName + Chr(250)
    a$ = a$ + Ts(MapProcess(i).TimesPlayed) + Chr(250)
    a$ = a$ + Chr(251)
    
Next i

'all set, send it
SendPacket "MP", a$, Index

End Sub

Sub PackageTeleporters(Index As Integer)
If DebugMode Then LastCalled = "PackageTeleporters"

'compiles and sends the connected user lists
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'players has:

'compile it

For i = 1 To NumTele
    a$ = a$ + Chr(251)
    a$ = a$ + Tele(i).Name + Chr(250)
    a$ = a$ + Ts(Tele(i).X) + Chr(250)
    a$ = a$ + Ts(Tele(i).Y) + Chr(250)
    a$ = a$ + Ts(Tele(i).Z) + Chr(250)
    a$ = a$ + Chr(251)
Next i

'all set, send it
SendPacket "TE", a$, Index

End Sub

Sub PackageMessages(Index As Integer, UserName As String)
If DebugMode Then LastCalled = "PackageConnectUsers"

'compiles and sends the connected user lists
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'players has:

'ConnectUsers.IP
'ConnectUsers.Name

'compile it


For i = 1 To NumMessages
    If LCase(Messages(i).MsgFor) = LCase(UserName) Or UserName = "" Then
        a$ = a$ + Chr(251)
        a$ = a$ + Ts(Messages(i).Flags) + Chr(250)
        a$ = a$ + Messages(i).MsgFor + Chr(250)
        a$ = a$ + Messages(i).MsgFrom + Chr(250)
        a$ = a$ + Ts(Messages(i).MsgId) + Chr(250)
        a$ = a$ + Messages(i).MsgSubj + Chr(250)
        a$ = a$ + Messages(i).MsgText + Chr(250)
        a$ = a$ + Ts(CDbl(Messages(i).MsgTimeSent)) + Chr(250)
        a$ = a$ + Chr(251)
    End If
Next i

'all set, send it
SendPacket "M2", a$, Index


End Sub

Sub PackageMessages2(Index As Integer, UserName As String)
If DebugMode Then LastCalled = "PackageMessages2"

'compiles and sends the connected user lists
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'players has:

'ConnectUsers.IP
'ConnectUsers.Name

'compile it



For i = 1 To NumMessages
    If LCase(Messages(i).MsgFor) = LCase(UserName) And CheckBit2(Messages(i).Flags, 1) Then
        Num = Num + 1
    End If
Next i

'all set, send it
SendPacket "M1", Ts(Num), Index



End Sub


Sub UnPackageNewMessage(p$, UserName As String)

Dim MsgData As typMessages
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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                
                m$ = Mid(a$, G, h - G)
                                
                If j = 1 Then MsgData.Flags = Val(m$)
                If j = 2 Then MsgData.MsgFor = m$
                If j = 3 Then MsgData.MsgSubj = m$
                If j = 4 Then MsgData.MsgText = m$
                
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

MsgData.MsgTimeSent = Now
MsgData.MsgFrom = UserName

If CheckBit2(MsgData.Flags, 1) = False Then MsgData.Flags = MsgData.Flags + 2



If MsgData.MsgFor = "(ALL)" Then 'send to all
    For i = 1 To NumUsers
        If CheckBit(i, 17) Then
            MsgData.MsgFor = Users(i).Name
            CreateMessage MsgData
        End If
    Next i
Else
    CreateMessage MsgData
End If



End Sub

Sub CreateMessage(MsgData As typMessages)

'adds this message


NumMessages = NumMessages + 1
i = NumMessages
ReDim Preserve Messages(0 To i)

Messages(i).Flags = MsgData.Flags
Messages(i).MsgFor = MsgData.MsgFor
Messages(i).MsgFrom = MsgData.MsgFrom
Messages(i).MsgSubj = MsgData.MsgSubj
Messages(i).MsgText = MsgData.MsgText
Messages(i).MsgTimeSent = MsgData.MsgTimeSent

'generate an id
Do
    Randomize
    j = Int(Rnd * 30000) + 1
    bd = 0
    For k = 1 To NumMessages
        If Messages(k).MsgId = j Then bd = 1: Exit For
    Next k
Loop Until bd = 0

Messages(i).MsgId = j

'all done



End Sub


Sub InterpritUsers(p$)
If DebugMode Then LastCalled = "InterpritUsers"

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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                m$ = Mid(a$, G, h - G)
                
                If j = 1 Then Users(i).Allowed = m$
                If j = 2 Then Users(i).Flags = Val(m$)
                If j = 3 Then Users(i).Name = m$
                If j = 4 Then Users(i).PassWord = m$
                If j = 5 Then Users(i).ICQ = m$
                If j = 6 Then Users(i).FTPRoot = m$
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0
NumUsers = i

'save it
SaveCommands


End Sub


Sub UnPackageScripts(p$)
If DebugMode Then LastCalled = "UnPackageScripts"

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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                
                m$ = Mid(a$, G, h - G)
                
                ReDim Preserve Commands(0 To i)
                
                If j = 1 Then Commands(i).Exec = m$
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
                If j = 12 Then Commands(i).ScriptID = Val(m$)
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

'save it
SaveCommands

For i = 1 To NumConnectUsers
    PackageMenuScripts ConnectUsers(i).Index
Next i

End Sub


Function UnPackageChangedScripts(p$) As String
If DebugMode Then LastCalled = "UnPackageChangedScripts"

'extracts scripts from the sent string

Dim CurrScript() As typScripts

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        ReDim CurrScript(0 To 1)
            
        h = 0
        j = 0
        Do
            G = h
            h = InStr(G + 1, a$, Chr(250))
            
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                
                m$ = Mid(a$, G, h - G)
                
                
                If j = 1 Then CurrScript(0).Exec = m$
                If j = 2 Then CurrScript(0).MustHave = Val(m$)
                If j = 3 Then CurrScript(0).Name = m$
                If j = 4 Then CurrScript(0).NumParams = Val(m$)
                If j = 5 Then CurrScript(0).ScriptName = m$
                If j = 6 Then CurrScript(0).AutoMakeVars = CBool(m$)
                If j = 7 Then CurrScript(0).Group = m$
                If j = 8 Then CurrScript(0).LogExec = CBool(m$)
                If j = 9 Then CurrScript(0).Unused1 = m$
                If j = 10 Then CurrScript(0).unused2 = m$
                If j = 11 Then CurrScript(0).unused3 = m$
                If j = 12 Then CurrScript(0).ScriptID = Val(m$)
                If j = 13 Then CurrScript(0).Unused5 = Val(m$)
                
                If j = 14 Then
                    CurrScript(0).NumButtons = Val(m$)
                    ReDim CurrScript(0).Buttons(0 To Val(m$))
                End If
                If j > 14 Then 'clan member list
                    
                    kk = (j - 15) Mod 5
                    k = ((j - 10) \ 5)
                    
                    If kk = 0 Then CurrScript(0).Buttons(k).ButtonName = m$
                    If kk = 1 Then CurrScript(0).Buttons(k).ButtonText = m$
                    If kk = 2 Then CurrScript(0).Buttons(k).OptionOff = m$
                    If kk = 3 Then CurrScript(0).Buttons(k).OptionOn = m$
                    If kk = 4 Then CurrScript(0).Buttons(k).Type = Val(m$)
                
                End If
                
            End If
        Loop Until h = 0
    
        'Now match this script to something
    
        j = 0
        For k = 1 To NumCommands
            If Commands(k).ScriptID = CurrScript(0).ScriptID Then j = k: Exit For
        Next k
        
        If j = 0 Then 'must be new script, create a new entry!
            NumCommands = NumCommands + 1
            ReDim Preserve Commands(0 To NumCommands)
            j = NumCommands
        End If
        
        If j > 0 Then
            'Copy script over!
            
            UnPackageChangedScripts = UnPackageChangedScripts + CurrScript(0).Name + ", "
            
            Commands(j).AutoMakeVars = CurrScript(0).AutoMakeVars
            Commands(j).Exec = CurrScript(0).Exec
            Commands(j).Group = CurrScript(0).Group
            Commands(j).LogExec = CurrScript(0).LogExec
            Commands(j).MustHave = CurrScript(0).MustHave
            Commands(j).Name = CurrScript(0).Name
            Commands(j).NumButtons = CurrScript(0).NumButtons
            Commands(j).NumParams = CurrScript(0).NumParams
            Commands(j).ScriptID = CurrScript(0).ScriptID
            Commands(j).ScriptName = CurrScript(0).ScriptName
            Commands(j).Unused1 = CurrScript(0).Unused1
            Commands(j).unused2 = CurrScript(0).unused2
            Commands(j).unused3 = CurrScript(0).unused3
            Commands(j).Unused5 = CurrScript(0).Unused5
            
            ReDim Preserve Commands(j).Buttons(0 To Commands(j).NumButtons)
            For k = 1 To Commands(j).NumButtons
                Commands(j).Buttons(k).ButtonName = CurrScript(0).Buttons(k).ButtonName
                Commands(j).Buttons(k).ButtonText = CurrScript(0).Buttons(k).ButtonText
                Commands(j).Buttons(k).OptionOff = CurrScript(0).Buttons(k).OptionOff
                Commands(j).Buttons(k).OptionOn = CurrScript(0).Buttons(k).OptionOn
                Commands(j).Buttons(k).Type = CurrScript(0).Buttons(k).Type
            Next k
        End If
    
    End If
Loop Until f = 0 Or e = 0

'save it
SaveCommands

For i = 1 To NumConnectUsers
    PackageMenuScripts ConnectUsers(i).Index
Next i

If Len(UnPackageChangedScripts) > 2 Then UnPackageChangedScripts = Left(UnPackageChangedScripts, Len(UnPackageChangedScripts) - 2)

End Function

Function UnPackageOneScripts(p$) As String
If DebugMode Then LastCalled = "UnPackageScripts"

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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                
                m$ = Mid(a$, G, h - G)
                
                If j = 1 Then nm$ = m$
                If j = 2 Then exc$ = m$
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

For i = 1 To NumCommands
    If Commands(i).Name = nm$ Then
        
        Commands(i).Exec = exc$
        Exit For
    End If
Next i

'save it
SaveCommands

UnPackageOneScripts = nm$
End Function


Sub UnPackageKickBans(p$)
    If DebugMode Then LastCalled = "UnPackageKickBans"
    
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
                G = h
                h = InStr(G + 1, a$, Chr(250))
                
                G = G + 1
                j = j + 1
                If G > 0 And h > G - 1 Then
                    
                    m$ = Mid(a$, G, h - G)
                    
                    If j = 1 Then KickBans(i).Clan = m$
                    If j = 2 Then KickBans(i).Name = m$
                    If j = 3 Then KickBans(i).Type = Val(m$)
                    If j = 4 Then KickBans(i).UID = m$
                
                End If
            Loop Until h = 0
        
        End If
    Loop Until f = 0 Or e = 0
    NumKickBans = i
    
    'save it
    SaveCommands


End Sub

Sub UnPackageAdminChat(p$, UserNum)
If DebugMode Then LastCalled = "UnPackageAdminChat"

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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                
                m$ = Mid(a$, G, h - G)
                
                If j = 1 Then chatmsg$ = m$
                If j = 2 Then chatcol = Val(m$)
            
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

SendChatToAll chatmsg$, chatcol, Users(UserNum).Name, Time$

End Sub


Sub UnPackageBSPEnts(p$, Index, j)
If DebugMode Then LastCalled = "UnPackageBSPEnts"

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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                
                m$ = Mid(a$, G, h - G)
                
                If j = 1 Then file$ = m$
                If j = 2 Then bsp$ = m$

            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

If CheckForFile(file$) = False Then
    SendPacket "MS", "File not found.", CInt(Index)
    Exit Sub
End If
ff$ = file$
SetBSPEnts file$, bsp$

SendPacket "MS", "Ent Update for " + ff$ + " recieved and set!", CInt(Index)
AddToLogFile "ENTUPD: " + Users(j).Name + " updated the ents for " + ff$

End Sub


Sub UnPackageAwayMode(p$, Index)
If DebugMode Then LastCalled = "UnPackageAwayMode"

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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                
                m$ = Mid(a$, G, h - G)
                
                If j = 1 Then mde = Val(m$)
                If j = 2 Then Msg$ = m$

            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

j = 0
For i = 1 To NumConnectUsers
    If ConnectUsers(i).Index = Index Then j = i: Exit For
Next i

If j > 0 Then
    ConnectUsers(j).AwayMode = mde
    ConnectUsers(j).AwayMsg = Msg$
End If

UpdateUsersList

End Sub
Sub UnPackageBanPlayer(p$, Index, nme$, usr)
If DebugMode Then LastCalled = "UnPackageBanPlayer"

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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            
            G = G + 1
            j = j + 1
            If G > 0 And h > G - 1 Then
                
                m$ = Mid(a$, G, h - G)
                
                If j = 1 Then UsID$ = m$
                If j = 2 Then Tme$ = m$
                If j = 3 Then rs$ = m$

            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

n = FindPlayer(UsID$)
If n > 0 Then

    If Val(Tme$) = 0 Then
        
        BanPlayerReason "", "", "Banned by " + nme$ + " - " + rs$, CInt(n)
        AddToLogFile "BAN: " + Users(j).Name + " permanently banned " + Players(n).Name + "(UNIQUE: " + Players(n).UniqueID + "). The entered reason was: """ + rs$ + """"
    Else
        SendActualRcon "banid " + Tme$ + " " + Players(n).UniqueID + " kick"
        'BanPlayer Players(n).UniqueID, Index, usr
        AddToLogFile "BAN: " + Users(j).Name + " banned " + Players(n).Name + "(UNIQUE: " + Players(n).UniqueID + ")  for " + Tme$ + " minutes."
    End If
End If
End Sub


Sub SendPacket(Cde As String, Params As String, Index As Integer)
If DebugMode Then LastCalled = "SendPacket"
On Error GoTo errocc

a$ = Chr(255) + Cde + Chr(255) + Params + Chr(255)

For i = 1 To NumConnectUsers
    If ConnectUsers(i).Index = Index Then j = i
Next i

If j = 0 Then Exit Sub

If ConnectUsers(j).EncryptedMode = True Then
    a$ = Encrypt(a$, ConnectUsers(j).PassWord)
End If

a$ = Chr(254) + Chr(254) + Chr(254) + a$ + Chr(253) + Chr(253) + Chr(253)

'tell the size
If Len(a$) > 5000 And Cde <> "SS" Then SendPacket "SS", Ts(Len(a$)), Index

If TCPCreated(Index) Then
    If Form1.TCP1(Index).State = sckConnected Then
        'send it in increments of 65000 bytes
        If Len(a$) <= 65000 Then
            Form1.TCP1(Index).SendData a$
        Else
            Do
                'cut off a segment
                If Len(a$) > 65000 Then
                    B$ = Left(a$, 65000)
                    'cut a$
                    a$ = Right(a$, Len(a$) - 65000)
                Else
                    B$ = a$
                End If
                
                Form1.TCP1(Index).SendData B$
                'DoEvents
            Loop Until Len(B$) < 65000
        End If
    End If
End If

'DoEvents

Exit Sub
errocc:
ErrorReport Err.Number, Err.Description + ", " + Err.Source

End Sub

Function CheckBit(UserNum, BitToCheck) As Boolean
If DebugMode Then LastCalled = "CheckBit"

Dim a As Variant, B As Variant
Dim c As Variant
Dim d As Variant
Dim d1 As Variant

If UserNum = 0 Then Exit Function
B = Users(UserNum).Flags
a = 2 ^ BitToCheck

c = CDec(B)
d = CDec(a)

'perform bitwise

For i = 36 To BitToCheck Step -1
    d1 = CDec(2 ^ i)
    If c >= d1 Then
        c = c - d1
        If d1 = d Then CheckBit = True: Exit Function
    End If
Next i

'If (B And a) = a Then CheckBit = True

End Function

Function CheckBit2(BitNum, BitToCheck) As Boolean
If DebugMode Then LastCalled = "CheckBit2"

Dim a As Long, B As Long
B = BitNum
a = 2 ^ BitToCheck

If (B And a) = a Then CheckBit2 = True

End Function

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

Sub AskTimeRemaining()
If DebugMode Then LastCalled = "AskTimeRemaining"
'asks for the map time remaining
LastRCON = ""
SendRCONCommand "mp_timeleft"
End Sub

Sub AskClanBattle()
If DebugMode Then LastCalled = "AskClanBattle"

'asks for the map time remaining
LastRCON = ""
SendRCONCommand "tfc_clanbattle"
SendRCONCommand "maxplayers"
End Sub


Function CalcMapTimeRemaining() As String
If DebugMode Then LastCalled = "CalcMapTimeRemaining"

'fills the map time remaining var with the right value

a = Vars.MapTimeLeft

Do
    If a >= 60 Then a = a - 60: m = m + 1
Loop Until a < 60

'a = Vars.MapTimeTotal - Vars.MapTimeElapsed - 1
'b = 60 - MapCounter

If m > 0 Then c$ = Ts(m) + " minutes "
If a > 0 Then c$ = c$ + Ts(a) + " seconds "

If c$ <> "" Then c$ = c$ + "remaining."
If c$ = "" Then c$ = "Unknown": d = 1

CalcMapTimeRemaining = c$
If Not Vars.ClanBattle And d = 0 And AnnTime Then
    If DLLEnabled = False Then
        SendRCONCommand "say Map Time: " + c$
    Else
        SendMessage "Map Time: " + c$, 1, 255, 147, 147, 1, 193, 1, 2, 6, 1, 0.01, 1, -1, 0.6
    End If
End If
AnnTime = False

If BanScriptData.UserIP <> "" Then
    SendToUser "Map Time: " + c$, BanScriptData
End If

BanScriptData.UserIP = ""


SendUpdate

End Function

Sub KickSomeOne(scriptdata As typScriptData)
        
        'kicks a player, in this criteria:
        '-not PLAYING
        '-no realname
        '-no clan tag
        '-most recent joiner
        
        Dim databox() As Integer
        ReDim databox(1 To NumPlayers)
        
        For i = 1 To NumPlayers
            If Players(i).ConnectOnly = True Then databox(i) = databox(i) + 3
            If InStr(1, Players(i).Name, "[") = 0 Then databox(i) = databox(i) + 1
            If Players(i).RealName = "" Then databox(i) = databox(i) + 2
            If i = NumPlayers Then databox(i) = databox(i) + 1
        Next i
        
        mx = 0
        For i = 1 To NumPlayers
            
            If databox(i) > mx Then
                mx = databox(i)
                u = i
            End If
        Next i
        
        If u = 0 Then
            SendToUser "Could not determine a player to be kicked.", scriptdata
            Exit Sub
        End If
        
        SendToUser "Kicking " + Players(u).Name, scriptdata
        SendRCONCommand "kick # " + Ts(Players(u).UserID)

End Sub

Sub CheckForDLL()
        If DebugMode Then LastCalled = "CheckForDLL"
        'ensures the DLL is there and running. If it isn't, disable DLL features.
        
        'first, make sure the file is actually there
        a$ = Server.BothPath + "\dlls\" + DLLFile
        B$ = Server.BothPath + "\liblist.gam"
        
        DLLEnabled = False
        If CheckForFile(a$) Then 'the file is there
        
            'now check to ensure its configured in LIBLIST.GAM
            
            If CheckForFile(B$) Then
            'gamedll "dlls\HPB_bot.dll"
            
                c$ = GetVarFromFile(B$, "gamedll")
                       
                If InStr(1, LCase(c$), LCase(DLLFile)) Then 'yes its set, so set the var to TRUE
                    DLLEnabled = True
                End If
            End If
        End If
        
        'DLLEnabled = True

End Sub

Function CheckWindowThere(WindowName As String) As Boolean
If DebugMode Then LastCalled = "CheckWindowThere"

'checks if the requested window is present, and returns TRUE if it is

'get the current window

hw = GetWindow(Form1.hwnd, GW_HWNDFIRST)

Do While hw <> 0
    hw = GetWindow(hw, GW_HWNDNEXT)
    a$ = WindowText(hw)
    If LCase(a$) = LCase(WindowName) Then CheckWindowThere = True: Exit Function
Loop

End Function

Function CloseWindow() As String
If DebugMode Then LastCalled = "CheckWindowThere"

'checks if the requested window is present, and returns TRUE if it is

'get the current window

hw = GetWindow(Form1.hwnd, GW_HWNDFIRST)

ar$ = ""

Do While hw <> 0
    hw = GetWindow(hw, GW_HWNDNEXT)
    a$ = WindowText(hw)
    ar$ = ar$ + a$ + vbCrLf
    If LeftR(a$, 4) = "HLDS" Then
        's = SendMessageWindow(hw, WM_CLOSE, 0, 0)
        m = GetWindowThreadProcessId(hw, kk)
        'm = GetCurrentProcess
        Debug.Print kk
        
        'm =
        hw2 = OpenProcess(PROCESS_TERMINATE, False, kk)
        
        k = TerminateProcess(hw2, 0)
        kk = CloseHandle(hw2)
        Debug.Print hw, k, kk
        
        
        Exit Function
    End If
Loop

CloseWindow = ar$

End Function

Function WindowText(hwnd) As String
If DebugMode Then LastCalled = "WindowText"

'gets the title of a window

l = GetWindowTextLength(hwnd)
l = l + 1
a$ = Space(l)

n = GetWindowText(hwnd, a$, l)
a$ = Left(a$, Len(a$) - 1)
a$ = Trim(a$)

WindowText = a$

End Function

Sub EventHandler()
If DebugMode Then LastCalled = "EventHandler"

'This sub gets called every 10 seconds.
'It's duty is to run through the events, see if any of them need to be called,
'or updated, or whatever.

'Lets start by going through the events.
Dim NowDate As Date
Dim RemoveFlag As Boolean
NowDate = Now

For i = 1 To NumEvents

    If Events(i).FirstCheck <= NowDate Then
        'This events NEXTCHECK is NOW, or has already passed.
        
        'Work with it.

        If Events(i).mde = 0 Then 'this event only gets called ONCE.
            'so start it!
            StartEvent i
            'set flag to remove it
            Events(i).Name = "***REMOVE_ME***"
            RemoveFlag = True
            
        Else
            'Type 1 or 2 .. runs only on certain days, so check if today is one of those days
            '0-mon, 1-tue, 2-wed, 3-thu, 4-fri, 5-sat, 6-sun
            dy = Weekday(NowDate) - 2
            If dy = -1 Then dy = 6 'sunday
                       
            If Events(i).Days(dy) = True Then 'yes, run it today
                StartEvent i
                CalcNextRun i
                If Events(i).Name = "***REMOVE_ME***" Then RemoveFlag = True
            Else 'do not run it today... so calculate the next time to run it
                CalcNextRun i
                If Events(i).Name = "***REMOVE_ME***" Then RemoveFlag = True
            End If
        End If
    End If
Next i

If RemoveFlag = True Then
again:
    For i = 1 To NumEvents
        If Events(i).Name = "***REMOVE_ME***" Then 'remove flagged events
            RemoveEvent i
            GoTo again
        End If
    Next i
End If

End Sub

Sub StartEvent(Num)
If DebugMode Then LastCalled = "StartEvent"

If LeftR(Events(Num).Name, 5) = "nolog" Then
Else
    AddToLogFile "EVENT: The event " + Chr(34) + Events(Num).Name + Chr(34) + " was started."
End If

If Events(Num).WhatToDo = 0 Then
    Dim NewScriptData As typScriptData
    NewScriptData.TimeStarted = Timer
    NewScriptData.StartedName = "EventStart: " + Events(Num).Name + ", " + Events(Num).ScriptName
    NewScriptData.UserName = "<SERVER>"

    ExecuteScriptParams Trim(Events(Num).ScriptName + " " + Events(Num).ComPara), NewScriptData
Else
    SendRCONCommand Trim(Events(Num).ComPara)
End If

End Sub

Sub CalcNextRun(Num)
If DebugMode Then LastCalled = "CalcNextRun"

Dim CalcDate As Date
Dim NowDate As Date
Dim DayDate As Date
NowDate = Now
DayDate = "01"
CalcDate = NowDate

Const MD_WEEK = 0
Const MD_DAY = 1
Const MD_HOUR = 2
Const MD_MIN = 3
Const MD_SEC = 4

'takes all available data on this event, and calculates when its to be run again
Dim DHour As Integer, DMin As Integer, DSec As Integer

'start with mode 2 -- once a day on certain days
If Events(Num).mde = 2 Then

    'at this time...
    DHour = Hour(Events(Num).FirstCheck)
    DMin = Minute(Events(Num).FirstCheck)
    DSec = Second(Events(Num).FirstCheck)
    
    'find the NEXT fitting weekday, up to 400 days away
    Do
        dy = Weekday(CalcDate) - 2
        If dy = -1 Then dy = 6 'sunday
                   
        If Events(Num).Days(dy) = True Then 'yes, this is a run day
            'set next run
            t$ = Ts(DHour) + ":" + Ts(DMin) + ":" + Ts(DSec)
            d$ = Ts(Day(CalcDate)) + "/" + Ts(Month(CalcDate)) + "/" + Ts(Year(CalcDate))
            Events(Num).FirstCheck = d$
            Events(Num).FirstCheck = Events(Num).FirstCheck + t$
            Exit Do
        End If
        'add a day
        CalcDate = CalcDate + DayDate
    Loop Until a > 400

    If a > 400 Then ' no day found
        Events(Num).Name = "***REMOVE_ME***"
    End If
    
    'all done! next day is set!!
   
Else 'this is of type Do this event x times every y zzzz's (like 5 times a day)

    'ok, we have the first date, lets find out how much we have to add to it
    
    addamt = Events(Num).Times
    
    
    CalcDate = Now
    
    If Events(Num).Every = MD_WEEK Then 'add x weeks to the date
        'do this by finding the date of the run, and then counting until the same day passes x times
        fdy = Weekday(Events(Num).FirstCheck) 'last check day
        a = 0
        Do
            CalcDate = CalcDate + DayDate 'add one day
            If Weekday(CalcDate) = fdy Then 'is this the day we want?
                a = a + 1 'add one
            End If
        Loop Until a >= addamt And CalcDate > NowDate
        'done
        Events(Num).FirstCheck = CalcDate
    
    ElseIf Events(Num).Every = MD_DAY Then
        'add x many days
        'CalcDate = Events(Num).FirstCheck
        a = 0
        Do
            CalcDate = CalcDate + DayDate 'add one day
            a = a + 1 'add one
        Loop Until a >= addamt And CalcDate > NowDate
        'done
        Events(Num).FirstCheck = CalcDate
    
    ElseIf Events(Num).Every = MD_HOUR Then
        'add x many hours
        'CalcDate = Events(Num).FirstCheck
        DayDate = "01:00:00"
        a = 0
        Do
            CalcDate = CalcDate + DayDate 'add one hour
            a = a + 1 'add one
        Loop Until a >= addamt And CalcDate > NowDate
        'done
        Events(Num).FirstCheck = CalcDate
    ElseIf Events(Num).Every = MD_MIN Then
        'add x many minutes
        'CalcDate = Events(Num).FirstCheck
        DayDate = "00:01:00"
        a = 0
        Do
            CalcDate = CalcDate + DayDate 'add one minute
            a = a + 1 'add one
        Loop Until a >= addamt And CalcDate > NowDate
        'done
        Events(Num).FirstCheck = CalcDate
    ElseIf Events(Num).Every = MD_SEC Then
        'add x many seconds
        'CalcDate = Events(Num).FirstCheck
        DayDate = "00:00:01"
        a = 0
        Do
            CalcDate = CalcDate + DayDate 'add one second
            a = a + 1 'add one
        Loop Until a >= addamt And CalcDate > NowDate
        'done
        Events(Num).FirstCheck = CalcDate
    End If
    
    'all done!!

End If

End Sub

'LOG SEARCHING FEATURE:

Sub LogSearch(SearchString As String, Check1 As Integer, FromDay As Date, ToDay As Date, _
Index As Integer, SearchSubs As Integer, SearchPath As String, ExactPhrase As Boolean, _
AllWords As Integer, SearchSaysOnly As Integer)

If Check1 = 0 Then FromDay = 0: ToDay = 0

'Start searching
LogSearchString = SearchString
NumLogFound = 0
ReDim LogFound(0 To 1)


Pth$ = Server.BothPath + "\logs\"
If Trim(SearchPath) <> "" Then Pth$ = Trim(SearchPath)
If RightR(Pth$, 1) <> "\" Then Pth$ = Pth$ + "\"

'make the search strings

Dim SearchArgs() As String
If Not ExactPhrase Then
    SearchArgs = Split(SearchString, " ")
Else
    ReDim SearchArgs(0 To 0)
    SearchArgs(0) = SearchString
End If
LogDoDir Pth$, Index, FromDay, ToDay, SearchArgs, SearchSubs, ExactPhrase, AllWords, SearchSaysOnly

'Done, now SEND this log found record
PackageLogSearch Index

End Sub

Private Sub LogDoDir(DirPath As String, Index As Integer, FromDay As Date, ToDay As Date, SearchArgs() As String, SearchSubs As Integer, ExactPhrase As Boolean, AllWords As Integer, SearchSaysOnly As Integer)
If DebugMode Then LastCalled = "LogDoDir"

'Scan for Sub-Directories
Dim MyPath As String
Dim MyName As String
pathyes = 0
MyPath = DirPath   ' Set the path.
MyName = Dir(MyPath, vbDirectory)   ' Retrieve the first entry.

If SearchSubs <> 0 Then
    Do While MyName <> ""   ' Start the loop.
       If MyName <> "." And MyName <> ".." Then
          If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
             
             LogDoDir DirPath + MyName + "\", Index, FromDay, ToDay, SearchArgs, SearchSubs, ExactPhrase, AllWords, SearchSaysOnly
              
             MyNewName = MyName
             MyName = Dir(MyPath, vbDirectory)
             Do
                MyName = Dir
             Loop Until MyNewName = MyName
             
             If NumLogFound >= 1000 Then Exit Sub
             
          End If   ' it represents a directory.
       End If
       MyName = Dir   ' Get next entry.Loop
    Loop
End If

LogItDir DirPath, Index, FromDay, ToDay, SearchArgs, SearchSubs, ExactPhrase, AllWords, SearchSaysOnly

End Sub

Private Sub LogItDir(DirPath As String, Index As Integer, FromDay As Date, ToDay As Date, SearchArgs() As String, SearchSubs As Integer, ExactPhrase As Boolean, AllWords As Integer, SearchSaysOnly As Integer)
If DebugMode Then LastCalled = "LogItDir"

Dim Xdate As Date

a$ = Dir(DirPath + "*.log")

SendPacket "L1", "Found: " + Ts(NumLogFound) + " matches" + vbCrLf + "Now Searching: " + DirPath, Index
DoEvents

If a$ = "" Then Exit Sub

Do
    
    If ToDay > 0 And FromDay > 0 Then
        Xdate = FileDateTime(DirPath + a$)
               
        If Xdate >= FromDay And Xdate <= ToDay Then
            LogDoFile a$, DirPath, Index, FromDay, ToDay, SearchArgs, SearchSubs, ExactPhrase, AllWords, SearchSaysOnly
        End If
    Else
        LogDoFile a$, DirPath, Index, FromDay, ToDay, SearchArgs, SearchSubs, ExactPhrase, AllWords, SearchSaysOnly
    End If
    
    a$ = Dir
    If a$ = "" Then Exit Sub
    If NumLogFound >= 1000 Then Exit Sub
    DoEvents
Loop

End Sub

Private Sub LogDoFile(FileName As String, DirPath As String, Index As Integer, FromDay As Date, ToDay As Date, SearchArgs() As String, SearchSubs As Integer, ExactPhrase As Boolean, AllWords As Integer, SearchSaysOnly As Integer)
If DebugMode Then LastCalled = "LogDoFile"

a$ = FileName

'Search this file for requested text
'

'Fle$ = ""
'h = freefile
Close h
'fln = FileLen(DirPath + a$)
'
'Open DirPath + a$ For Binary As h
'    Do While Not (EOF(h))
'        Fle$ = Fle$ + Input(fln, #h)
'    Loop
'
'    Fle$ = Replace(Fle$, vbCrLf, Chr(10))
'
'    For I = 0 To UBound(SearchArgs)
'
'        Do
'
'            e = InStr(e + 1, LCase(Fle$), LCase(SearchArgs(I)))
'
'
'            If e > 0 And NumLogFound < 1000 Then
'                'get the line
'
'                f = InStrRev(Fle$, Chr(10), e)
'                G = InStr(e, Fle$, Chr(10))
'                If f = 0 Then f = 1
'                If G = 0 Then G = Len(Fle$)
'                'extract
'
'                c$ = Mid(Fle$, f + 1, G - f - 1)
'
'                If AllWords = 1 Then
'                    For j = 0 To UBound(SearchArgs)
'                        If InStr(1, LCase(c$), LCase(SearchArgs(j))) = 0 Then
'                            GoTo nxtline
'                        End If
'                    Next j
'                End If
'                If SearchSaysOnly = 1 Then
'                    If InStr(1, LCase(c$), "say " + Chr(34)) = 0 And InStr(1, LCase(c$), "say_team " + Chr(34)) = 0 Then
'                        GoTo nxtline
'                    End If
'                End If
'
'                'add this log line
'                NumLogFound = NumLogFound + 1
'                ReDim Preserve LogFound(0 To NumLogFound)
'
'                c$ = Replace(c$, Chr(250), " ")
'                c$ = Replace(c$, Chr(251), " ")
'                c$ = Replace(c$, Chr(252), " ")
'                c$ = Replace(c$, Chr(253), " ")
'                c$ = Replace(c$, Chr(254), " ")
'                c$ = Replace(c$, Chr(255), " ")
'                c$ = Replace(c$, Chr(0), " ")
'                c$ = Replace(c$, Server.RCONPass, "****")
'
'                LogFound(NumLogFound).LogLine = c$
'                LogFound(NumLogFound).LogFile = DirPath + a$
'            End If
'
'        Loop Until e = 0
'    Next I
'Close h


h = FreeFile
Close h


filesiz = FileLen(DirPath + a$)
If filesiz > 0 Then
    
    ' Make the string the same length as the file
    Dim theFileData As String
    
    theFileData = Space(filesiz)
    Open DirPath + a$ For Binary As h
        Get #h, , theFileData
    Close h

    theFileData = Replace(theFileData, vbCrLf, vbCr)
    theFileData = Replace(theFileData, vbLf, vbCr)
    
    ' Search for the arguements in the file.
    
    n = 0
    
    For i = 0 To UBound(SearchArgs)
    
        arg12 = LCase(SearchArgs(i))
        
        anyfound = 0
        
        e = 0
        Do
            f = e
            e = InStr(e + 1, theFileData, arg12)
            foundmatch = 0
            
            If e > 0 Then
                
                ' Got a match. Extract the line.
                
                G = InStrRev(theFileData, vbCr, e)
                
                h = InStr(e, theFileData, vbCr)
                If h = 0 Then h = Len(theFileData) + 1
                
                If G < e And h > e Then
                    lne$ = Mid(theFileData, G + 1, h - G - 1)
                    foundmatch = 1
                    'Got the line.
                    
                
                    If SearchSaysOnly = 1 Then
                        If InStr(1, lne$, "say " + Chr(34)) = 0 And InStr(1, lne$, "say_team " + Chr(34)) = 0 Then
                            foundmatch = 0
                        End If
                    End If
                
            
                    If AllWords = 1 And foundmatch = 1 And UBound(SearchArgs) > 0 Then
                        ' Make sure this line has all the words.
                        n = 0
                        foundmatch = 0
                        For k = 0 To UBound(SearchArgs)
                            arg13 = LCase(SearchArgs(k))
                            If InStr(1, lne$, arg13) > 0 Then n = n + 1
                        Next
                        If n = UBound(SearchArgs) + 1 Then foundmatch = 1
                    End If
                    
                    If foundmatch = 1 And NumLogFound < 1000 Then
                        'add this log line
                        NumLogFound = NumLogFound + 1
                        ReDim Preserve LogFound(0 To NumLogFound)
            
                        c$ = lne$
                        c$ = Replace(c$, Chr(250), " ")
                        c$ = Replace(c$, Chr(251), " ")
                        c$ = Replace(c$, Chr(252), " ")
                        c$ = Replace(c$, Chr(253), " ")
                        c$ = Replace(c$, Chr(254), " ")
                        c$ = Replace(c$, Chr(255), " ")
                        c$ = Replace(c$, Chr(0), " ")
                        c$ = Replace(c$, Server.RCONPass, "**PASS**")
            
                        LogFound(NumLogFound).LogLine = c$
                        LogFound(NumLogFound).LogFile = DirPath + a$
                    End If
                End If
            End If
        Loop Until e = 0
        

    Next
    

End If

'
'Open DirPath + a$ For Input As h
'    Do While Not (EOF(h))
'        Line Input #h, B$
'
'        'check this line
'
'        n = 0
'        For i = 0 To UBound(SearchArgs)
'            e = InStr(1, LCase(B$), ))
'            If e Then n = n + 1
'        Next i
'
'        If AllWords = 1 Then
'            If n < UBound(SearchArgs) + 1 Then GoTo nxtline
'        End If
'        If SearchSaysOnly = 1 Then
'            If InStr(1, LCase(B$), "say " + Chr(34)) = 0 And InStr(1, LCase(B$), "say_team " + Chr(34)) = 0 Then
'                GoTo nxtline
'            End If
'        End If
'
'        If n > 0 And NumLogFound < 1000 Then
'            'add this log line
'            NumLogFound = NumLogFound + 1
'            ReDim Preserve LogFound(0 To NumLogFound)
'
'            c$ = B$
'            c$ = Replace(c$, Chr(250), " ")
'            c$ = Replace(c$, Chr(251), " ")
'            c$ = Replace(c$, Chr(252), " ")
'            c$ = Replace(c$, Chr(253), " ")
'            c$ = Replace(c$, Chr(254), " ")
'            c$ = Replace(c$, Chr(255), " ")
'            c$ = Replace(c$, Chr(0), " ")
'            c$ = Replace(c$, Server.RCONPass, "****")
'
'            LogFound(NumLogFound).LogLine = c$
'            LogFound(NumLogFound).LogFile = DirPath + a$
'        End If
'nxtline:
'    Loop
'Close h

End Sub

Function GetBSPEnts(BSPPath As String) As String
If DebugMode Then LastCalled = "GetBSPEnts"

Dim Size As Long
Dim CurrBank As String

CurrBank = Space(BankSize)
Size = InitEnts(BSPPath)
If Size > 0 Then
    CurrBank = Left(CurrBank, Size)
    ddd = EntData(CurrBank)
    
    GetBSPEnts = CurrBank
Else
    CurrBank = ""
    GetBSPEnts = "An Error has Occured."
End If

CurrBank = vbNullString

DllCanUnloadNow
End Function

Sub SetBSPEnts(BSPPath As String, Ents As String)
If DebugMode Then LastCalled = "SetBSPEnts"

Ents = Replace(Ents, Chr(0), "")
Size = ImportEnts(BSPPath, Ents)
DllCanUnloadNow


End Sub

Sub CheckIfServerCrashed()

'This will ask the server for something, and wait 10 seconds for a reply. If none arrives, it will try again, and then alert admins.

'LastCrashCall Modes:
'0 = normal
'1 = Sent first USERS request, waiting for reply
'2 = Sent SECOND users request, waiting for reply
'3 = Server is not responding


If LastCrashCall = 1 Then 'Checking for first time
        
    'The server was asked for USERS, but did not reply within 10 seconds.
    'Ask again.
    LastCrashCall = 2
    CrashTimer = 40
    SendRCONCommand "users"

ElseIf LastCrashCall = 2 Then ' Checked for a second time, STILL no reply. The server has crashed.

    LastCrashCall = 3
    SendUpdate
    SendRCONCommand "say Server Assistant cannot achieve a link with the server."
    AlertAdmins "The Server is not Responding!" + vbCrLf + "It may have crashed. User Intervention is required to restart it."
End If

End Sub

Sub AlertAdmins(Msg As String)

'loop thru any admins with the ICQ flag set. Send the ICQ, waiting between.

For i = 1 To NumUsers
    If CheckBit(i, 35) Then
        If Users(i).ICQ <> "" Then

            'send the message
            SendICQMessage Users(i).ICQ, "Message from SERVER " + Server.HostName + vbCrLf + Msg
            
            SendingICQ = True
            
            'wait up to 30 seconds.
            st = Timer
            Do
                tmgo = Timer - st
                DoEvents
            Loop Until tmgo > 30 Or SendingICQ = False Or Timer < st
            
            'done.
            
            'now, wait an additional 1 sec.
            
            st = Timer
            Do
                tmgo = Timer - st
                DoEvents
            Loop Until tmgo > 1 Or Timer < st
            
            'now, continue sending ICQs.
        End If
    End If
Next i

End Sub

Sub GetBansFromBanlist()
If DebugMode Then LastCalled = "GetBansFromBanlist"


'Now, read BANLIST.CFG, the file created by Server Assistant.

ReDim CurrBans(0 To 0)
NumCurrBans = 0

B$ = Server.BothPath + "\banlist.cfg"

If CheckForFile(B$) Then

    a$ = ""
    h = FreeFile
Close h
    Open B$ For Binary As h
        Do While Not (EOF(h))
            a$ = a$ + Input(65000, #h)
        Loop
    Close h

End If

a$ = Replace(a$, vbCrLf, Chr(10))



'blank line seperates bans

currnum = NumCurrBans + 1
lastline = 0

Un$ = ""
en$ = ""
rn$ = ""
rs$ = ""
tm$ = ""
Mp$ = ""
dt$ = ""
nm$ = ""

f = 0
Do
    e = f
    f = InStr(e + 1, a$, Chr(10))
    
    If f <> 0 Then
        d$ = Mid(a$, e + 1, f - e - 1)
        d$ = Trim(d$)
       
        If d$ = "" Then
            currnum = currnum + 1
            'set indicating there was a blank line
            lastline = 1
        End If
        
        h = InStr(1, d$, " ")
        j = InStr(h + 1, d$, " ")
        k = InStr(j + 1, d$, " ")

'    //Date/Time/Map: 12-12-2000 10:55:06 - CROSSFIRE
'    //Player Name: BiIIDoor
'    //UniqueID: 371537
'    //Reason: Banned by Avatar-X.
'    //IP: 63.21.148.253
'    //Real Name: **KEA**Juden[BSMFH]-[A]-{AoE}[+
'    //Entry Name: **KEA**Juden[BSMFH]-[A]-{AoE}[+
'
        Lg$ = "//Date/Time/Map:"
        If LeftR(d$, Len(Lg$)) = Lg$ Then
            'read date, time, and map
            
            If h > 0 And j > 0 And k > 0 And j > h And k - 1 > j Then
                'date
                dt$ = Mid(d$, h + 1, j - h - 1)
                'time
                tm$ = Mid(d$, j + 1, k - j - 1)
                'map
                Mp$ = Right(d$, Len(d$) - k - 2)
            End If
            lastline = 0
        End If
       
        ' un$ = "111870" Then MsgBox "here"
        
       
        Lg$ = "//Player Name:"
        If LeftR(d$, Len(Lg$)) = Lg$ Then
            If j > 0 Then
                'name
                nm$ = Right(d$, Len(d$) - j)
            End If
            lastline = 0
        End If
       
        Lg$ = "//UniqueID:"
        If LeftR(d$, Len(Lg$)) = Lg$ Then
            If h > 0 Then
                'unique
                Un$ = Right(d$, Len(d$) - h)
            End If
            lastline = 0
        End If
       
        Lg$ = "//Reason:"
        If LeftR(d$, Len(Lg$)) = Lg$ Then
            If h > 0 Then
                'reason
                rs$ = Right(d$, Len(d$) - h)
            End If
            lastline = 0
        End If
        
        Lg$ = "//IP:"
        If LeftR(d$, Len(Lg$)) = Lg$ Then
            If h > 0 Then
                'reason
                IP$ = Right(d$, Len(d$) - h)
            End If
            lastline = 0
        End If
       
        Lg$ = "//Real Name:"
        If LeftR(d$, Len(Lg$)) = Lg$ Then
            If j > 0 Then
                'real name
                rn$ = Right(d$, Len(d$) - j)
            End If
            lastline = 0
        End If
       
        Lg$ = "//Entry Name:"
        If LeftR(d$, Len(Lg$)) = Lg$ Then
            If j > 0 Then
                'entry name
                en$ = Right(d$, Len(d$) - j)
            End If
            lastline = 0
        End If
       
        If lastline = 1 And Un$ <> "" Then
       
            which = 0
            For i = 1 To NumCurrBans
                If InStr(1, CurrBans(i).UIDs, Un$) Then which = i: Exit For
            Next i
            
            If which = 0 Then
                NumCurrBans = NumCurrBans + 1
                ReDim Preserve CurrBans(0 To NumCurrBans)
                which = NumCurrBans
            End If
            
            If Un$ <> "" Then CurrBans(which).UIDs = Un$
            CurrBans(which).BanTime = "0"
            If en$ <> "" Then CurrBans(which).EntryName = en$
            If IP$ <> "" Then CurrBans(which).IP = IP$
            If Mp$ <> "" Then CurrBans(which).Map = Mp$
            If nm$ <> "" Then CurrBans(which).Name = nm$
            If rn$ <> "" Then CurrBans(which).RealName = rn$
            If rs$ <> "" Then CurrBans(which).Reason = rs$
            If dt$ <> "" Then CurrBans(which).BannedAt = dt$ + " " + tm$
            
            'clear
            Un$ = ""
            en$ = ""
            rn$ = ""
            rs$ = ""
            tm$ = ""
            Mp$ = ""
            dt$ = ""
            nm$ = ""
            IP$ = ""
            
        End If
    End If
nxt3:
Loop Until f = 0



'Now, lets do the BANNED.CFG file.

B$ = Server.BothPath + "\banned.cfg"

If CheckForFile(B$) Then

    a$ = ""
    h = FreeFile
Close h
    Open B$ For Binary As h
        Do While Not (EOF(h))
            a$ = a$ + Input(65000, #h)
        Loop
    Close h

End If

'fortmat:
'banid <time> <uid>
'banid 0.0 1056867
a$ = Replace(a$, vbCrLf, Chr(10))

f = 0
Do
    e = f
    f = InStr(e + 1, a$, Chr(10))
    
    If f <> 0 Then
        d$ = Mid(a$, e + 1, f - e - 1)
        d$ = Trim(d$)
        If LeftR(d$, 2) = "//" Then GoTo nxt2
       
        h = InStr(1, d$, " ")
        j = InStr(h + 1, d$, " ")
       
        If h > 0 And j > h + 1 Then
            'uniqueid
            tm$ = Mid(d$, h + 1, j - h - 1)
            tm$ = Trim(tm$)
            
            Un$ = Trim(RightR(d$, Len(d$) - j))
            
            'add to list
            
            nodo = 0
            For i = 1 To NumCurrBans
                If InStr(1, CurrBans(i).UIDs, Un$) Then nodo = 1
            Next i
            
            If nodo = 0 Then
                NumCurrBans = NumCurrBans + 1
                ReDim Preserve CurrBans(0 To NumCurrBans)
                CurrBans(NumCurrBans).UIDs = Un$
                           
                If tm$ = "permanent" Then tm$ = "0"
                CurrBans(NumCurrBans).BanTime = tm$
            End If
        End If
    End If
nxt2:
Loop Until f = 0


End Sub



Sub GetBans(a$, Index)
If DebugMode Then LastCalled = "GetBans"

'loads all the bans from various files
GetBanList = False

'FIRST, lets load all of the data from the server's LISTID command

'User filter list:
'17155499:  permanent
'5343608:  permanent

f = 0
Do
    e = f
    f = InStr(e + 1, a$, Chr(10))
    
    If f <> 0 Then
        d$ = Mid(a$, e + 1, f - e - 1)
        d$ = Trim(d$)
        If Left(d$, 5) = "User " Then GoTo nxt
       
        h = InStr(1, d$, ":")
       
        If h > 0 Then
            'uniqueid
            Un$ = LeftR(d$, h - 1)
            Un$ = Trim(Un$)
            
            'add to list
            
            NumCurrBans = NumCurrBans + 1
            ReDim Preserve CurrBans(0 To NumCurrBans)
            CurrBans(NumCurrBans).UIDs = Un$
            
            tm$ = Trim(RightR(d$, Len(d$) - h))
            
            If tm$ = "permanent" Then tm$ = "0"
            CurrBans(NumCurrBans).BanTime = tm$
            
        End If
    End If
nxt:
Loop Until f = 0


'DONE server LISTID command.

'Package bans and send to player.



a$ = ""

For i = 1 To NumCurrBans
    a$ = a$ + Chr(251)
    a$ = a$ + CurrBans(i).BannedAt + Chr(250)
    a$ = a$ + CurrBans(i).BanTime + Chr(250)
    a$ = a$ + CurrBans(i).EntryName + Chr(250)
    a$ = a$ + CurrBans(i).IP + Chr(250)
    a$ = a$ + CurrBans(i).Map + Chr(250)
    a$ = a$ + CurrBans(i).Name + Chr(250)
    a$ = a$ + CurrBans(i).RealName + Chr(250)
    a$ = a$ + CurrBans(i).Reason + Chr(250)
    a$ = a$ + CurrBans(i).UIDs + Chr(250)
    a$ = a$ + Chr(251)
Next i

SendPacket "GB", a$, CInt(Index)

End Sub

Sub TestProc(scriptdata As typScriptData)

SendToDebug "Starting Process Search...", scriptdata

'GetAllProc32 ScriptData
a$ = CloseWindow

SendToDebug "Got " + vbCrLf + a$, scriptdata

'
'For ic = 0 To vgProc32(0).count - 1
'
'Fle$ = LCase$(vgProc32(ic).exeFile)
'm = vgProc32(ic).procID
'
'a$ = a$ + Fle$ + vbCrLf
'
'e = InStrRev(Fle$, "\")
'If e > 0 Then
'    Fle$ = Right(Fle$, Len(Fle$) - e)
'End If
'
'e = InStr(1, Fle$, Chr(0))
'
'If e > 1 Then
'    Fle$ = Left(Fle$, e - 1)
'End If
'
'If Fle$ = "hlds.exe" Then
'
'    Debug.Print Fle$
'    Debug.Print m
'
'    Dim c As Long
'
'    'B = GetExitCodeProcess(m, c)
'
'
'    hw = OpenProcess(PROCESS_TERMINATE, False, m)
'
'    k = TerminateProcess(hw, 0)
'    kk = CloseHandle(hw)
'    'Debug.Print "c is", c
'    Debug.Print hw, k, kk
'End If
'
'Next ic

SendToDebug "Done!", scriptdata

End Sub

Public Sub GetAllProc32(scriptdata As typScriptData)

  Dim iIdx     As Integer
  Dim bRet     As Boolean
  Dim hSnap    As Long
  Dim p32stru  As PROCESSENTRY32

  hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
 
  SendToDebug "HSnap: " + Ts(hSnap), scriptdata

  p32stru.dwSize = Len(p32stru)
  
  bRet = Process32First(hSnap, p32stru)
  
  SendToDebug "BRet: " + Ts(bRet), scriptdata
  
  vgTotalThrd = 0
  iIdx = 0
  
  vgProc32(0).Index = 0
  vgProc32(0).procID = p32stru.th32ProcessID
  vgProc32(0).threadCount = p32stru.cntThreads
  vgProc32(0).countUsage = p32stru.cntUsage
  vgProc32(0).defaultHeapID = p32stru.th32DefaultHeapID
  vgProc32(0).moduleID = p32stru.th32ModuleID
  vgProc32(0).parentProcessID = p32stru.th32ParentProcessID
  vgProc32(0).pcPriClassBase = p32stru.pcPriClassBase
  vgProc32(0).Flags = p32stru.dwFlags
  vgProc32(0).exeFile = Trim(p32stru.szExeFile)
  vgTotalThrd = vgTotalThrd + vgProc32(0).threadCount
  
  Call GetThreadsByProc32(vgProc32(0).procID, 0)
  Call GetModulesByProc32(vgProc32(0).procID, 0)
  
  While True
  
     bRet = Process32Next(hSnap, p32stru)
     If bRet = False Then
        If GetLastError() = 18 Then GoTo END_GAP
        Exit Sub
     End If
     
     iIdx = iIdx + 1
     vgProc32(iIdx).Index = iIdx
     vgProc32(iIdx).procID = p32stru.th32ProcessID
     vgProc32(iIdx).threadCount = p32stru.cntThreads
     vgProc32(iIdx).countUsage = p32stru.cntUsage
     vgProc32(iIdx).defaultHeapID = p32stru.th32DefaultHeapID
     vgProc32(iIdx).moduleID = p32stru.th32ModuleID
     vgProc32(iIdx).parentProcessID = p32stru.th32ParentProcessID
     vgProc32(iIdx).pcPriClassBase = p32stru.pcPriClassBase
     vgProc32(iIdx).Flags = p32stru.dwFlags
     vgProc32(iIdx).exeFile = Trim(p32stru.szExeFile)
     vgTotalThrd = vgTotalThrd + vgProc32(iIdx).threadCount
  
     Call GetThreadsByProc32(vgProc32(iIdx).procID, iIdx)
     Call GetModulesByProc32(vgProc32(iIdx).procID, iIdx)
     Call IncProcInd(1)
     
     DoEvents
     
  Wend
  
END_GAP:

  vgProc32(0).count = iIdx + 1

  bRet = CloseHandle(hSnap)

End Sub

Public Sub GetThreadsByProc32(procID As Long, anIdx As Integer)

  Dim iIdx     As Integer
  Dim bRet     As Boolean
  Dim hSnap    As Long
  Dim t32stru  As THREAD32ENTRY
  
  hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, procID)

  t32stru.dwSize = Len(t32stru)
  
  
  bRet = Thread32First(hSnap, t32stru)
  
  iIdx = -1
  
  If t32stru.th32OwnerProcessID = procID Then
     iIdx = 0
     vgProc32(anIdx).thrdInfo(0).Index = 0
     vgProc32(anIdx).thrdInfo(0).thrdId = t32stru.th32ThreadID
     vgProc32(anIdx).thrdInfo(0).BasePriority = t32stru.tpBasePri
     vgProc32(anIdx).thrdInfo(0).countUsage = t32stru.cntUsage
  End If
  
  While True
  
     bRet = Thread32Next(hSnap, t32stru)
     If bRet = False Then
        If GetLastError() = 18 Then GoTo END_GTBP32
        Exit Sub
     End If
     
     If t32stru.th32OwnerProcessID = procID Then
        iIdx = iIdx + 1
        vgProc32(anIdx).thrdInfo(iIdx).Index = iIdx
        vgProc32(anIdx).thrdInfo(iIdx).thrdId = t32stru.th32ThreadID
        vgProc32(anIdx).thrdInfo(iIdx).BasePriority = t32stru.tpBasePri
        vgProc32(anIdx).thrdInfo(iIdx).countUsage = t32stru.cntUsage
        Call IncProcInd(1)
     End If
     
     DoEvents
     
  Wend
  
END_GTBP32:

  bRet = CloseHandle(hSnap)

End Sub

Public Sub GetModulesByProc32(procID As Long, anIdx As Integer)

  Dim iIdx     As Integer
  Dim bRet     As Boolean
  Dim hSnap    As Long
  Dim m32stru  As MODULEENTRY32
  
  hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, procID)

  m32stru.dwSize = Len(m32stru)
  
  bRet = Module32First(hSnap, m32stru)
  
  iIdx = -1
  
  If m32stru.th32ProcessID = procID Then
     iIdx = 0
     vgProc32(anIdx).moduleInfo(0).Index = 0
     vgProc32(anIdx).moduleInfo(0).moduleFullPath = m32stru.szExePath
  End If
  
  While True
  
     bRet = Module32Next(hSnap, m32stru)
     If bRet = False Then
        If GetLastError() = 18 Then GoTo END_GMBP32
        Exit Sub
     End If
     
     If m32stru.th32ProcessID = procID Then
        iIdx = iIdx + 1
        vgProc32(anIdx).moduleInfo(iIdx).Index = 0
        vgProc32(anIdx).moduleInfo(iIdx).moduleFullPath = m32stru.szExePath
        vgProc32(anIdx).countModules = iIdx + 1
        Call IncProcInd(1)
     End If
     
     DoEvents
     
  Wend
  
END_GMBP32:

  bRet = CloseHandle(hSnap)

End Sub

Public Sub IncProcInd(Val As Integer)

  vgProcInd = vgProcInd + Val
  If vgProcInd > 100 Then vgProcInd = 100
  'frmTaskMonitor.prgInd.Value = vgProcInd

End Sub

Sub GetTeamNames()

TeamNames(1) = "Blue"
TeamNames(2) = "Red"
TeamNames(3) = "Yellow"
TeamNames(4) = "Green"

f$ = Server.BothPath + "\Assistant"
If Dir(f$, vbDirectory) = "" Then MkDir f$

f$ = Server.BothPath + "\Assistant\Data"
If Dir(f$, vbDirectory) = "" Then MkDir f$
f$ = f$ + "\teamname.dat"

If CheckForFile(f$) Then
    h = FreeFile
    If DebugMode Then LastCalled = "TeamNames - AfterFreeFile"
    Close h
    
    If DebugMode Then LastCalled = "TeamNames - Start of Loop"
    
    Open f$ For Input As h
        Do While Not EOF(h)
            Input #h, Mp$, tm1$, tm2$, tm3$, tm4$
            
            If LCase(Mp$) = LCase(Vars.Map) Then
            
                TeamNames(1) = tm1$
                TeamNames(2) = tm2$
                TeamNames(3) = tm3$
                TeamNames(4) = tm4$
                
                If TeamNames(1) = "" Then TeamNames(1) = "Blue"
                If TeamNames(2) = "" Then TeamNames(2) = "Red"
                If TeamNames(3) = "" Then TeamNames(3) = "Yellow"
                If TeamNames(4) = "" Then TeamNames(4) = "Green"
                
                fnd = 1
                Exit Do
            End If
        Loop
    Close h
    If DebugMode Then LastCalled = "TeamNames - End of Loop"
End If

If fnd = 1 Then Exit Sub
AddToLogFile "TEAMNAMES.DAT: Entry for map " + Vars.Map + " not found! Starting search..."

Dim Ents() As Byte


'Ents = GetBSPEnts(Server.BothPath + "\maps\" + Vars.Map + ".bsp")


Mp$ = Server.BothPath + "\maps\" + Vars.Map + ".bsp"

If CheckForFile(Mp$) = False Then Exit Sub
GetFile Mp$, Ents


'"team1_name" "foolworthy"





TeamNames(1) = GetKeyValue(Ents, "team1_name")
TeamNames(2) = GetKeyValue(Ents, "team2_name")
TeamNames(3) = GetKeyValue(Ents, "team3_name")
TeamNames(4) = GetKeyValue(Ents, "team4_name")

If TeamNames(1) = "" Then TeamNames(1) = "Blue"
If TeamNames(2) = "" Then TeamNames(2) = "Red"
If TeamNames(3) = "" Then TeamNames(3) = "Yellow"
If TeamNames(4) = "" Then TeamNames(4) = "Green"


'now, save.

h = FreeFile
Close h
Open f$ For Append As h
    Print #h, Chr(34) + LCase(Vars.Map) + Chr(34) + ", " + Chr(34) + TeamNames(1) + Chr(34) + ", " + Chr(34) + TeamNames(2) + Chr(34) + ", " + Chr(34) + TeamNames(3) + Chr(34) + ", " + Chr(34) + TeamNames(4) + Chr(34)
Close h

AddToLogFile "TEAMNAMES.DAT: Finished search and saved team-names."



End Sub

Function SearchInByteArray(ByRef Rett() As Byte, Srch As String) As Long

Dim SrchArr() As Byte
ReDim SrchArr(1 To Len(Srch))

For i = 1 To Len(Srch)
    SrchArr(i) = Asc(Mid(Srch, i, 1))
Next i

For i = 1 To UBound(Rett) - 1
    If Rett(i) = SrchArr(1) And Rett(i + 1) = SrchArr(2) Then
        'check if rest matches
        mtch = 1
        For j = 2 To Len(Srch)
            If Rett(i + j - 1) <> SrchArr(j) Then
                mtch = 0
            End If
        Next j
        If mtch = 1 Then k = i: Exit For
    End If
Next i

SearchInByteArray = k

End Function

Function GetKeyValue(Rett() As Byte, KeyName As String) As String

e = SearchInByteArray(Rett, KeyName)

If e > 0 Then
    e = e + Len(KeyName)
    f$ = ""
    For i = e + 3 To UBound(Rett)
    
        If Rett(i) = 34 Then Exit For
        f$ = f$ + Chr(Rett(i))
    
    Next i

End If
'e = InStr(1, Srch, Chr(34) + KeyName + Chr(34))
'f = InStr(e + 1, Srch, Chr(34))
'f = InStr(f + 1, Srch, Chr(34))
'If e > 0 And f > e Then
'    e = f
'    f = InStr(e + 1, Srch, Chr(34))
'
'    If e > 0 And f > e Then
'        nm$ = Mid(Srch, e + 1, f - e - 1)
'    End If
'End If
GetKeyValue = f$

End Function

Function GetFile(a$, Rett() As Byte)
h = FreeFile
Close h
ret$ = ""

If Not CheckForFile(a$) Then Exit Function

f = FileLen(a$)

'Dim ret2 As String * f

'Dim Rett() As Byte
ReDim Rett(0 To f)


Open a$ For Binary As h
    Do Until EOF(h)
        
        Get #h, , Rett
        
        'ret$ = ret$ + Input(1000, #h)
       '
       ' If Len(ret$) > 32000 Then
       '     ret2 = ret2 + ret$
       '     ret$ = ""
       ' End If
        
    Loop
Close h
'ret2 = ret2 + ret$

End Function


Sub UnPackageGameRequest(p$, Index, who$)
    If DebugMode Then LastCalled = "UnPackageGameRequest"
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
                G = h
                h = InStr(G + 1, a$, Chr(250))
                G = G + 1
                j = j + 1
                If G > 0 And h > G - 1 Then
                    m$ = Mid(a$, G, h - G)
                    
                    If j = 1 Then GameName$ = m$
                    If j = 2 Then ToWho$ = m$
                    If j = 3 Then GameID = Val(m$)
                    
                End If
            Loop Until h = 0
        
        End If
    Loop Until f = 0 Or e = 0
    
    If ToWho$ <> "" Then 'only to one guy
        
        For i = 1 To NumConnectUsers
            If LCase(ConnectUsers(i).Name) = LCase(ToWho$) And ConnectUsers(i).HiddenMode = False Then
                j = i
                SendGameRequestPacket GameName$, who$, CInt(Index), ConnectUsers(j).Index, GameID
            End If
        Next i
        
        If j = 0 Then
            SendPacket "MS", "Connected Admin Not Found!", CInt(Index)
        End If
        
    Else
        For i = 1 To NumConnectUsers
           If Index <> ConnectUsers(i).Index Then SendGameRequestPacket GameName$, who$, CInt(Index), ConnectUsers(i).Index, GameID
        Next i
    End If

End Sub


Sub SendGameRequestPacket(GameName As String, WhoFrom As String, Index As Integer, IndexTo As Integer, GameID)

'compiles and sends the clan info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'compile it

a$ = a$ + Chr(251)
a$ = a$ + GameName + Chr(250)
a$ = a$ + WhoFrom + Chr(250)
a$ = a$ + Ts(Index) + Chr(250)
a$ = a$ + Ts(GameID) + Chr(250)
a$ = a$ + Chr(251)

'all set, send it
SendPacket "G1", a$, IndexTo

End Sub


Sub UnPackageGamePacket(p$, Index, who$)
    If DebugMode Then LastCalled = "UnPackageGamePacket"
    
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
                G = h
                h = InStr(G + 1, a$, Chr(250))
                G = G + 1
                j = j + 1
                If G > 0 And h > G - 1 Then
                    m$ = Mid(a$, G, h - G)
                    
                    If j = 1 Then IndexTo = Val(m$)
                    If j = 2 Then GameID = Val(m$)
                    If j = 3 Then HisGameID = Val(m$)
                    If j = 4 Then DataToPass$ = m$
                
                End If
            Loop Until h = 0
        
        End If
    Loop Until f = 0 Or e = 0
    
    SendGamePacket DataToPass$, who$, CInt(GameID), CInt(Index), CInt(IndexTo), HisGameID


End Sub

Sub SendGamePacket(PacketData As String, who$, GameID As Integer, IndexFrom As Integer, IndexTo As Integer, HisGameID)

a$ = a$ + Chr(251)
a$ = a$ + Ts(IndexFrom) + Chr(250)
a$ = a$ + who$ + Chr(250)
a$ = a$ + Ts(GameID) + Chr(250)
a$ = a$ + Ts(HisGameID) + Chr(250)
a$ = a$ + PacketData + Chr(250)
a$ = a$ + Chr(251)

SendPacket "G2", a$, IndexTo

End Sub



