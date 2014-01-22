Attribute VB_Name = "Module2"
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
'      FILE: server-db.bas
'      PURPOSE: This file has to do with when I tried making SA use an
'      Access DB, but it turned out to be a disaster. So I scrapped it.
'
'
' ===========================================================================
' ---------------------------------------------------------------------------

' ****************************************************************
' THIS MODULE CONTAINS ALL CODE IN RELATION TO THE ACCESS DATABASE
' ****************************************************************

'Define Databases
'Public Db As Database
'Public Rs As Recordset
'private RecordNames(1 To 20, 0 To 20) As String
'
'Sub OpenData()
'If DebugMode Then LastCalled = "OpenData"
'
'p$ = App.Path + "\assistant.mdb"
'
'If CheckForFile(p$) Then
'    Set Db = OpenDatabase(p$)
'Else
'    If DebugMode Then LastCalled = "OpenData - Create Database"
'    Set Db = CreateDatabase(p$, dbLangGeneral)
'End If
'
'DefineRecordNames
'
'End Sub
'
'Function OpenData2() As Boolean
'If DebugMode Then LastCalled = "OpenData2"
'
'p$ = App.Path + "\assistant.mdb"
'
'If CheckForFile(p$) Then
'    Set Db = OpenDatabase(p$)
'
'    OpenData2 = True
'
'    DefineRecordNames
'Else
'    AddToLogFile "SERVER: Database file " + p$ + " not found."
'End If
'
'End Function
'
'Sub DefineRecordNames()
'
'
'RecordNames(1, 0) = "Commands"
'RecordNames(2, 0) = "Buttons"
'RecordNames(3, 0) = "Users"
'RecordNames(4, 0) = "KickBans"
'RecordNames(5, 0) = "Clans"
'RecordNames(6, 0) = "Members"
'RecordNames(7, 0) = "KillList"
'RecordNames(8, 0) = "RealPlayers"
'RecordNames(9, 0) = "Web"
'RecordNames(10, 0) = "WebColors"
'RecordNames(11, 0) = "General"
'RecordNames(12, 0) = "Events"
'RecordNames(13, 0) = "Messages"
'RecordNames(14, 0) = "Speech"
'RecordNames(15, 0) = "SpeechAnswers"
'
'a = 1
'RecordNames(a, 1) = "Name"
'RecordNames(a, 2) = "Exec"
'RecordNames(a, 3) = "NumParams"
'RecordNames(a, 4) = "MustHave"
'RecordNames(a, 5) = "ScriptName"
'RecordNames(a, 6) = "NumButtons"
'RecordNames(a, 7) = "ID"
'a = a + 1
'
'RecordNames(a, 1) = "ButtonName"
'RecordNames(a, 2) = "ButtonText"
'RecordNames(a, 3) = "OptionOn"
'RecordNames(a, 4) = "OptionOff"
'RecordNames(a, 5) = "Type"
'RecordNames(a, 6) = "BelongsTo"
'a = a + 1
'
'RecordNames(a, 1) = "Name"
'RecordNames(a, 2) = "Password"
'RecordNames(a, 3) = "Allowed"
'RecordNames(a, 4) = "Flags"
'a = a + 1
'
'RecordNames(a, 1) = "Name"
'RecordNames(a, 2) = "Clan"
'RecordNames(a, 3) = "UniqueID"
'RecordNames(a, 4) = "Type"
'a = a + 1
'
''5
'RecordNames(a, 1) = "Clan"
'RecordNames(a, 2) = "JoinPass"
'RecordNames(a, 3) = "NumMembers"
'RecordNames(a, 4) = "Flags"
'RecordNames(a, 5) = "ID"
'a = a + 1
'
''6
'RecordNames(a, 1) = "Name"
'RecordNames(a, 2) = "UniqueID"
'RecordNames(a, 3) = "LastIP"
'RecordNames(a, 4) = "BelongsTo"
'a = a + 1
'
''7
'RecordNames(a, 1) = "Name"
'RecordNames(a, 2) = "Ent"
'RecordNames(a, 3) = "Award"
'a = a + 1
'
''8
'RecordNames(a, 1) = "RealName"
'RecordNames(a, 2) = "UniqueID"
'RecordNames(a, 3) = "LastName"
'RecordNames(a, 4) = "LastTime"
'RecordNames(a, 5) = "Flags"
'a = a + 1
'
''9
'RecordNames(a, 1) = "Enabled"
'RecordNames(a, 2) = "LogPath"
'RecordNames(a, 3) = "LogFlags"
'a = a + 1
'
''10
'RecordNames(a, 1) = "ColorNum"
'RecordNames(a, 2) = "R"
'RecordNames(a, 3) = "G"
'RecordNames(a, 4) = "B"
'a = a + 1
'
''11
'RecordNames(a, 1) = "NoAutoVotes"
'RecordNames(a, 2) = "NoKickVotes"
'RecordNames(a, 3) = "MaxMsg"
'RecordNames(a, 4) = "MaxKickVotes"
'RecordNames(a, 5) = "MaxTime"
'RecordNames(a, 6) = "MaxKicks"
'RecordNames(a, 7) = "BanTime"
'RecordNames(a, 8) = "VotePercent"
'RecordNames(a, 9) = "LoggingDisabled"
'RecordNames(a, 10) = "LastMapsDisabled"
'RecordNames(a, 11) = "Flags"
'RecordNames(a, 12) = "MaxSpeech"
'RecordNames(a, 13) = "MaxSpeechTime"
'RecordNames(a, 14) = "MapVoteStartTime"
'RecordNames(a, 15) = "MapVoteStartTimeMode"
'RecordNames(a, 16) = "MapVoteMode"
'a = a + 1
'
''12
'RecordNames(a, 1) = "Mode"
'RecordNames(a, 2) = "Times"
'RecordNames(a, 3) = "Every"
'RecordNames(a, 4) = "Days0"
'RecordNames(a, 5) = "Days1"
'RecordNames(a, 6) = "Days2"
'RecordNames(a, 7) = "Days3"
'RecordNames(a, 8) = "Days4"
'RecordNames(a, 9) = "Days5"
'RecordNames(a, 10) = "Days6"
'RecordNames(a, 11) = "FirstCheck"
'RecordNames(a, 12) = "WhatToDo"
'RecordNames(a, 13) = "ScriptName"
'RecordNames(a, 14) = "ComPara"
'RecordNames(a, 15) = "Name"
'a = a + 1
'
''13
'RecordNames(a, 1) = "MsgText"
'RecordNames(a, 2) = "MsgFor"
'RecordNames(a, 3) = "MsgFrom"
'RecordNames(a, 4) = "MsgTimeSent"
'RecordNames(a, 5) = "MsgSubj"
'RecordNames(a, 6) = "MsgID"
'RecordNames(a, 7) = "Flags"
'a = a + 1
'
''14- speech
'RecordNames(a, 1) = "ClientText"
'RecordNames(a, 2) = "NumAnswers"
'RecordNames(a, 3) = "ID"
'a = a + 1
'
''15 - speech answers
'RecordNames(a, 1) = "Answer"
'RecordNames(a, 2) = "BelongsTo"
'
''
''RecordNames(a, 1) = ""
''RecordNames(a, 2) = ""
''RecordNames(a, 3) = ""
''RecordNames(a, 4) = ""
''RecordNames(a, 5) = ""
''RecordNames(a, 6) = ""
''RecordNames(a, 7) = ""
''a = a + 1
''
'
'
'End Sub
'
'Sub MakeDBTest()
'
'OpenData
'
''Set Db = New DBEngine
'
'Dim Tb As TableDef
'Set Tb = Db.CreateTableDef("Abc")
'
''Db.CreateTableDef "Abc2",
'
'
''Tb '
'
'Tb.Fields.Append Tb.CreateField("Field1", dbMemo)
'Tb.Fields.Append Tb.CreateField("Field2", dbMemo)
'Tb.Fields.Append Tb.CreateField("Field3", dbMemo)
'
'Db.TableDefs.Append Tb
'
'Set Rs = Db.OpenRecordset("Abc")
'Rs.AddNew
'Rs.Close
'
'Db.Close
'
'
'End Sub
'
'Function IsTableThere(Num) As Boolean
'If DebugMode Then LastCalled = "IsTableThere"
'
'Dim tdfLoop As TableDef
'
'For Each tdfLoop In Db.TableDefs
'    If tdfLoop.Name = RecordNames(Num, 0) Then IsTableThere = True: Exit For
'Next tdfLoop
'
'End Function
'
'Function IsFieldThere(Num, SubNum) As Boolean
'If DebugMode Then LastCalled = "IsFieldThere"
'
'Dim tdfLoop As Field
'
'For Each tdfLoop In Rs.Fields
'    If tdfLoop.Name = RecordNames(Num, SubNum) Then IsFieldThere = True: Exit For
'Next tdfLoop
'
'End Function
'
'Sub CreateTable(Num)
'If DebugMode Then LastCalled = "CreateTable"
'
'Dim Tb As TableDef
'Dim Ff As Field
'Set Tb = Db.CreateTableDef(RecordNames(Num, 0))
'
'For I = 1 To 20
'
'    a$ = RecordNames(Num, I)
'    If a <> "" Then
'
'        Set Ff = Tb.CreateField
'
'        Ff.Name = a$
'        Ff.Type = dbMemo
'        Ff.AllowZeroLength = True
'
'        Tb.Fields.Append Ff
'    Else
'        Exit For
'    End If
'Next I
'
''done, append it
'Db.TableDefs.Append Tb
'
'End Sub
'
'Sub AddField(Num, SubNum)
'
'Dim Ff As Field
'Dim Tb As TableDef
'
'Set Tb = Db.TableDefs(Rs.Name)
'Set Ff = Tb.CreateField
'
'Ff.Name = RecordNames(Num, SubNum)
'Ff.Type = dbMemo
'Ff.AllowZeroLength = True
'
'Rs.Fields.Append Ff
'
'End Sub
'
'Sub SaveDB()
'If DebugMode Then LastCalled = "SaveDB"
'
'Dim MissingField
'
''Try to open the file
'OpenData
'If DebugMode Then LastCalled = "SaveDB"
'
'AddToLogFile "SERVER: Starting Save of DB..."
'
''Start by making the Commands
'
'For I = 1 To 15
'
'    If DebugMode Then LastCalled = "SaveDB - i = " + Ts(I)
'
'    If IsTableThere(I) = False Then
'        'Create it
'        CreateTable I
'    End If
'
'    'Now Open it
'    Set Rs = Db.OpenRecordset(RecordNames(I, 0))
'
'    'Check to ensure all fields exist
'
'    MissingField = False
'    For j = 1 To 20
'        If RecordNames(I, j) = "" Then Exit For
'        If IsFieldThere(I, j) = False Then MissingField = True
'    Next j
'
'    If MissingField = True Then
'
'        'Delete this table
'        Rs.Close
'        Db.TableDefs.Delete RecordNames(I, 0)
'
'        'Re-create it
'
'        CreateTable I
'        'Now Open it
'        Set Rs = Db.OpenRecordset(RecordNames(I, 0))
'    End If
'
'    'Great! Now Delete all records in this table
'
'    If Not Rs.EOF Then Rs.MoveFirst
'
'    Do Until Rs.EOF
'        Rs.Delete
'        Rs.MoveNext
'    Loop
'
'    'Now add the records in
'
'    If I = 1 Then
'        For k = 1 To NumCommands
'            Rs.AddNew
'            With Rs
'
'                !Name = Commands(k).Name
'                !Exec = Commands(k).Exec
'                !NumParams = Ts(Commands(k).NumParams)
'                !MustHave = Ts(Commands(k).MustHave)
'                !ScriptName = Commands(k).ScriptName
'                !NumButtons = Ts(Commands(k).NumButtons)
'                !ID = Ts(k)
'                .Update
'            End With
'        Next k
'
'    ElseIf I = 2 Then
'
'        For k = 1 To NumCommands
'            For kk = 1 To Commands(k).NumButtons
'                Rs.AddNew
'                With Rs
'
'                    !ButtonName = Commands(k).Buttons(kk).ButtonName
'                    !ButtonText = Commands(k).Buttons(kk).ButtonText
'                    !OptionOn = Commands(k).Buttons(kk).OptionOn
'                    !OptionOff = Commands(k).Buttons(kk).OptionOff
'                    !Type = Ts(Commands(k).Buttons(kk).Type)
'                    !BelongsTo = Ts(k)
'                    .Update
'
'                End With
'            Next kk
'        Next k
'
'    ElseIf I = 3 Then
'        For k = 1 To NumUsers
'            Rs.AddNew
'            With Rs
'
'                !Name = Users(k).Name
'                !Password = Users(k).Password
'                !Allowed = Users(k).Allowed
'                !Flags = Ts(Users(k).Flags)
'
'                .Update
'            End With
'        Next k
'
'    ElseIf I = 4 Then
'        For k = 1 To NumKickBans
'            Rs.AddNew
'            With Rs
'
'                !Name = KickBans(k).Name
'                !Clan = KickBans(k).Clan
'                !UniqueID = KickBans(k).uID
'                !Type = Ts(KickBans(k).Type)
'
'                .Update
'            End With
'        Next k
'
'    ElseIf I = 5 Then
'        For k = 1 To NumClans
'            Rs.AddNew
'            With Rs
'
'                !Clan = Clans(k).Clan
'                !JoinPass = Clans(k).JoinPass
'                !NumMembers = Ts(Clans(k).NumMembers)
'                !Flags = Ts(Clans(k).Flags)
'                !ID = Ts(k)
'
'                .Update
'            End With
'        Next k
'
'
'    ElseIf I = 6 Then
'        For k = 1 To NumClans
'            For kk = 1 To Clans(k).NumMembers
'                Rs.AddNew
'                With Rs
'
'                    !Name = Clans(k).Members(kk).Name
'                    !LastIP = Clans(k).Members(kk).LastIP
'                    !UniqueID = Clans(k).Members(kk).UIN
'                    !BelongsTo = Ts(k)
'                    .Update
'
'                End With
'            Next kk
'        Next k
'
'    ElseIf I = 7 Then
'        For k = 1 To NumKills
'            Rs.AddNew
'            With Rs
'
'                !Name = KillList(k).Name
'                !Ent = KillList(k).Ent
'                !Award = KillList(k).Award
'
'                .Update
'            End With
'        Next k
'    ElseIf I = 8 Then
'        For k = 1 To NumRealPlayers
'            Rs.AddNew
'            With Rs
'
'                !RealName = RealPlayers(k).RealName
'                !UniqueID = RealPlayers(k).UniqueID
'                !LastName = RealPlayers(k).LastName
'                !LastTime = Ts(CDbl(RealPlayers(k).LastTime))
'                !Flags = Ts(RealPlayers(k).Flags)
'                .Update
'            End With
'        Next k
'    ElseIf I = 9 Then
'
'        Rs.AddNew
'        With Rs
'
'            !Enabled = CInt(Web.Enabled)
'            !LogPath = Web.LogPath
'            !LogFlags = Ts(Web.LogFlags)
'            .Update
'        End With
'
'    ElseIf I = 10 Then
'
'        For k = 1 To 21
'            Rs.AddNew
'            With Rs
'
'                !ColorNum = Ts(k)
'                !R = Ts(Web.Colors(k).R)
'                !G = Ts(Web.Colors(k).G)
'                !B = Ts(Web.Colors(k).B)
'
'                .Update
'            End With
'        Next k
'
'    ElseIf I = 11 Then
'
'        Rs.AddNew
'        With Rs
'
'            !NoAutoVotes = Ts(CInt(General.NoAutoVotes))
'            !NoKickVotes = Ts(CInt(General.NoKickVotes))
'            !MaxMsg = Ts(General.MaxMsg)
'            !MaxKickVotes = Ts(General.MaxKickVotes)
'            !MaxTime = Ts(General.MaxTime)
'            !MaxKicks = Ts(General.MaxKicks)
'            !BanTime = Ts(General.BanTime)
'            !VotePercent = Ts(General.VotePercent)
'            !LoggingDisabled = Ts(CInt(General.LoggingDisabled))
'            !LastMapsDisabled = Ts(CInt(General.LastMapsDisabled))
'            !Flags = Ts(General.Flags)
'            !MaxSpeech = Ts(General.MaxSpeech)
'            !MaxSpeechTime = Ts(General.MaxSpeechTime)
'            !MapVoteStartTime = Ts(General.MapVoteStartTime)
'            !MapVoteStartTimeMode = Ts(General.MapVoteStartTimeMode)
'            !MapVoteMode = General.MapVoteMode
'
'            .Update
'        End With
'
'    ElseIf I = 12 Then
'
'        For k = 1 To NumEvents
'            Rs.AddNew
'            With Rs
'
'                !Mode = Ts(Events(k).mde)
'                !Times = Ts(Events(k).Times)
'                !Every = Ts(Events(k).Every)
'                !Days0 = Ts(CInt(Events(k).Days(0)))
'                !Days1 = Ts(CInt(Events(k).Days(1)))
'                !Days2 = Ts(CInt(Events(k).Days(2)))
'                !Days3 = Ts(CInt(Events(k).Days(3)))
'                !Days4 = Ts(CInt(Events(k).Days(4)))
'                !Days5 = Ts(CInt(Events(k).Days(5)))
'                !Days6 = Ts(CInt(Events(k).Days(6)))
'                !FirstCheck = Ts(CDbl(Events(k).FirstCheck))
'                !WhatToDo = Ts(Events(k).WhatToDo)
'                !ScriptName = Events(k).ScriptName
'                !ComPara = Events(k).ComPara
'                !Name = Events(k).Name
'
'                .Update
'            End With
'        Next k
'
'    ElseIf I = 13 Then
'
'        For k = 1 To NumMessages
'            Rs.AddNew
'            With Rs
'
'                !MsgText = Messages(k).MsgText
'                !MsgFor = Messages(k).MsgFor
'                !MsgFrom = Messages(k).MsgFrom
'                !MsgTimeSent = Ts(CDbl(Messages(k).MsgTimeSent))
'                !MsgSubj = Messages(k).MsgSubj
'                !MsgId = Ts(Messages(k).MsgId)
'                !Flags = Ts(Messages(k).Flags)
'
'                .Update
'            End With
'        Next k
'
'    ElseIf I = 14 Then
'
'        For k = 1 To NumSpeech
'            Rs.AddNew
'            With Rs
'
'                !ClientText = Speech(k).ClientText
'                !NumAnswers = Speech(k).NumAnswers
'                !ID = Ts(k)
'
'                .Update
'            End With
'        Next k
'
'    ElseIf I = 15 Then
'        For k = 1 To NumSpeech
'            For kk = 1 To Speech(k).NumAnswers
'                Rs.AddNew
'                With Rs
'
'                    !Answer = Speech(k).Answers(kk)
'                    !BelongsTo = Ts(k)
'                    .Update
'
'                End With
'            Next kk
'        Next k
'
'
'    End If
'
'    Rs.Close
'
'Next I
'
'Db.Close
'
'AddToLogFile "SERVER: DB was SAVED!"
'
'End Sub
'
'Sub LoadDB()
'If DebugMode Then LastCalled = "LoadDB"
'
'AddToLogFile "SERVER: Starting Load of DB..."
'
''Loads the DB
'
''Store Which fields are available for data pulling, so we dont steal data from blank fields
''Dim AvailFields(1 To 20, 0 To 20) As String
'
'On Error GoTo errocc
'
'If OpenData2 = False Then
'    Exit Sub
'End If
'
''Start by making the Commands
'
'For I = 1 To 15
'
'    If IsTableThere(I) = True Then
'
'        'Now Open it
'        Set Rs = Db.OpenRecordset(RecordNames(I, 0))
'
'        'Check to ensure all fields exist
'
'        MissingField = False
'        For j = 1 To 20
'            If RecordNames(I, j) = "" Then Exit For
'            If IsFieldThere(I, j) = False Then MissingField = True
'        Next j
'
'        If MissingField = True Then
'            AddToLogFile "DBLOAD: Missing field(s) in table " + RecordNames(I, 0) + "! Could be due to a newer version! Continuing load."
'        End If
'
'        'Great! Here we are.
'
'        With Rs
'            If Not .EOF Then
'                .MoveFirst
'                k = 0
'                Do Until .EOF
'                    k = k + 1
'
'                    If I = 1 Then
'
'                        'Commands
'
'                        kk = Val(GetValue(!ID))
'                        Commands(kk).Name = GetValue(!Name)
'                        Commands(kk).Exec = GetValue(!Exec)
'                        Commands(kk).NumParams = Val(GetValue(!NumParams))
'                        Commands(kk).MustHave = Val(GetValue(!MustHave))
'                        Commands(kk).ScriptName = GetValue(!ScriptName)
'                        Commands(kk).NumButtons = 0 'Val(GetValue(!NumButtons))
'                        NumCommands = k
'
'                    ElseIf I = 2 Then
'
'                        'Buttons
'                        kk = Val(GetValue(!BelongsTo))
'                        Commands(kk).NumButtons = Commands(kk).NumButtons + 1
'                        kkj = Commands(kk).NumButtons
'                        ReDim Preserve Commands(kk).Buttons(0 To kkj)
'
'                        Commands(kk).Buttons(kkj).ButtonName = GetValue(!ButtonName)
'                        Commands(kk).Buttons(kkj).ButtonText = GetValue(!ButtonText)
'                        Commands(kk).Buttons(kkj).OptionOn = GetValue(!OptionOn)
'                        Commands(kk).Buttons(kkj).OptionOff = GetValue(!OptionOff)
'                        Commands(kk).Buttons(kkj).Type = Val(GetValue(!Type))
'
'                    ElseIf I = 3 Then
'
'                        'Users
'
'                        Users(k).Name = GetValue(!Name)
'                        Users(k).Allowed = GetValue(!Allowed)
'                        Users(k).Password = GetValue(!Password)
'                        Users(k).Flags = Val(GetValue(!Flags))
'                        NumUsers = k
'
'                    ElseIf I = 4 Then
'
'                        KickBans(k).Clan = GetValue(!Clan)
'                        KickBans(k).Name = GetValue(!Name)
'                        KickBans(k).uID = GetValue(!UniqueID)
'                        KickBans(k).Type = Val(GetValue(!Type))
'                        NumKickBans = k
'
'                    ElseIf I = 5 Then
'
'                        kk = Val(GetValue(!ID))
'                        Clans(kk).Clan = GetValue(!Clan)
'                        Clans(kk).JoinPass = GetValue(!JoinPass)
'                        Clans(kk).Flags = Val(GetValue(!Flags))
'                        Clans(kk).NumMembers = 0
'                        NumClans = k
'                    ElseIf I = 6 Then
'
'                        kk = Val(GetValue(!BelongsTo))
'                        Clans(kk).NumMembers = Clans(kk).NumMembers + 1
'                        kkj = Clans(kk).NumMembers
'
'                        Clans(kk).Members(kkj).Name = GetValue(!Name)
'                        Clans(kk).Members(kkj).UIN = GetValue(!UniqueID)
'                        Clans(kk).Members(kkj).LastIP = GetValue(!LastIP)
'
'                    ElseIf I = 7 Then
'                        ReDim Preserve KillList(0 To k)
'                        KillList(k).Name = GetValue(!Name)
'                        KillList(k).Ent = GetValue(!Ent)
'                        KillList(k).Award = GetValue(!Award)
'                        NumKills = k
'
'                    ElseIf I = 8 Then
'
'                        ReDim Preserve RealPlayers(0 To k)
'                        RealPlayers(k).RealName = GetValue(!RealName)
'                        RealPlayers(k).UniqueID = GetValue(!UniqueID)
'                        RealPlayers(k).LastName = GetValue(!LastName)
'                        RealPlayers(k).LastTime = CDate(GetValue(!LastTime))
'                        RealPlayers(k).Flags = Val(GetValue(!Flags))
'                        NumRealPlayers = k
'
'                    ElseIf I = 9 Then
'
'                        Web.Enabled = CBool(GetValue(!Enabled))
'                        Web.LogPath = GetValue(!Enabled)
'                        Web.LogFlags = Val(GetValue(!Enabled))
'
'                    ElseIf I = 10 Then
'
'                        kk = Val(GetValue(!ColorNum))
'                        Web.Colors(kk).R = Val(GetValue(!R))
'                        Web.Colors(kk).G = Val(GetValue(!R))
'                        Web.Colors(kk).B = Val(GetValue(!R))
'
'                    ElseIf I = 11 Then
'
'                        General.NoAutoVotes = CBool(GetValue(!NoAutoVotes))
'                        General.NoKickVotes = CBool(GetValue(!NoKickVotes))
'                        General.MaxMsg = Val(GetValue(!MaxMsg))
'                        General.MaxKickVotes = Val(GetValue(!MaxKickVotes))
'                        General.MaxTime = Val(GetValue(!MaxTime))
'                        General.MaxKicks = Val(GetValue(!MaxKicks))
'                        General.BanTime = Val(GetValue(!BanTime))
'                        General.VotePercent = Val(GetValue(!VotePercent))
'                        General.LoggingDisabled = CBool(GetValue(!LoggingDisabled))
'                        General.LastMapsDisabled = CBool(GetValue(!LastMapsDisabled))
'                        General.Flags = Val(GetValue(!Flags))
'                        General.MaxSpeech = Val(GetValue(!MaxSpeech))
'                        General.MaxSpeechTime = Val(GetValue(!MaxSpeechTime))
'                        General.MapVoteStartTime = Val(GetValue(!MapVoteStartTime))
'                        General.MapVoteStartTimeMode = Val(GetValue(!MapVoteStartTimeMode))
'                        General.MapVoteMode = GetValue(!MapVoteMode)
'
'                    ElseIf I = 12 Then
'
'                        ReDim Preserve Events(0 To k)
'                        Events(k).mde = Val(GetValue(!Mode))
'                        Events(k).Times = Val(GetValue(!Times))
'                        Events(k).Every = Val(GetValue(!Every))
'                        Events(k).Days(0) = Val(GetValue(!Days0))
'                        Events(k).Days(1) = Val(GetValue(!Days1))
'                        Events(k).Days(2) = Val(GetValue(!Days2))
'                        Events(k).Days(3) = Val(GetValue(!Days3))
'                        Events(k).Days(4) = Val(GetValue(!Days4))
'                        Events(k).Days(5) = Val(GetValue(!Days5))
'                        Events(k).Days(6) = Val(GetValue(!Days6))
'                        Events(k).FirstCheck = CDate(GetValue(!FirstCheck))
'                        Events(k).WhatToDo = Val(GetValue(!WhatToDo))
'                        Events(k).ScriptName = GetValue(!ScriptName)
'                        Events(k).ComPara = GetValue(!ComPara)
'                        Events(k).Name = GetValue(!Name)
'                        NumEvents = k
'
'                    ElseIf I = 13 Then
'
'                        ReDim Preserve Messages(0 To k)
'                        Messages(k).MsgText = GetValue(!MsgText)
'                        Messages(k).MsgFor = GetValue(!MsgFor)
'                        Messages(k).MsgFrom = GetValue(!MsgFrom)
'                        Messages(k).MsgTimeSent = CDate(GetValue(!MsgTimeSent))
'                        Messages(k).MsgSubj = GetValue(!MsgSubj)
'                        Messages(k).MsgId = Val(GetValue(!MsgId))
'                        Messages(k).Flags = GetValue(!Flags)
'                        NumMessages = k
'
'                    ElseIf I = 14 Then
'
'                        kk = Val(GetValue(!ID))
'                        ReDim Preserve Speech(0 To k)
'                        Speech(kk).ClientText = GetValue(!ClientText)
'                        Speech(kk).NumAnswers = 0
'                        NumSpeech = k
'
'                    ElseIf I = 15 Then
'
'
'                        kk = Val(GetValue(!BelongsTo))
'                        Speech(kk).NumAnswers = Speech(kk).NumAnswers + 1
'                        kkj = Speech(kk).NumAnswers
'                        ReDim Preserve Speech(kk).Answers(0 To kkj)
'
'                        Speech(kk).Answers(kkj) = GetValue(!Answer)
'
'                    End If
'
'                    .MoveNext
'                Loop
'            End If
'        End With
'
'        Rs.Close
'    End If
'Next I
'
'
'Db.Close
'
'AddToLogFile "SERVER: DB was LOADED and Renamed!"
'
'p$ = App.Path + "\assistant.mdb"
'
'If CheckForFile(p$) Then
'    If CheckForFile(App.Path + "\assistold.mdb") Then Kill App.Path + "\assistold.mdb"
'    Name p$ As App.Path + "\assistold.mdb"
'End If
'
'Exit Sub
'
'errocc:
'
'ErrorReport Err.Number, Err.Description + ", " + Err.Source
'
'Resume Next
'
'
'End Sub
'
'Function GetValue(Val As Variant) As String
'
'On Error Resume Next
'If Len(Val) > 0 Then GetValue = Val
'
'End Function

