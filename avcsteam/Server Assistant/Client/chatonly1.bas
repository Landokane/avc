Attribute VB_Name = "Module1"
Public Type Playersss
    Name As String      'name
    Class As Integer    'currentclass (-1 = civ, 0 = random, 1-9 = scout-engy) [only applies to TFC servers]
    Team As Integer     'current team
    Status As Boolean
    JoinTime As Date
End Type

Public Type typRGB
    R As Byte
    G As Byte
    B As Byte
End Type

Public RichColors(1 To 10) As typRGB

Public DllEnabled As Boolean
Public GameMode As Integer

'time/update
Public SecondsLeft As Integer
Public MapName As String
Public PlayersOn As String
Public Players(1 To 400) As Playersss
Public NumPlayers As Integer

Function MessBox(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String) As Long

Dim MessageBox As New frmMessageBox

MessageBox.Prompt = Prompt
MessageBox.Buttons = Buttons
MessageBox.Title = Title

MessageBox.Display
MessageBox.ReturnValue = -1
Do
    DoEvents
Loop Until MessageBox.ReturnValue <> -1

MessBox = MessageBox.ReturnValue
Unload MessageBox

End Function

Sub Main()
DllEnabled = True

RichColors(1).R = 0
RichColors(1).G = 0
RichColors(1).B = 0

RichColors(2).R = 0
RichColors(2).G = 0
RichColors(2).B = 255

RichColors(3).R = 255
RichColors(3).G = 0
RichColors(3).B = 0

RichColors(4).R = 150
RichColors(4).G = 150
RichColors(4).B = 0

RichColors(5).R = 0
RichColors(5).G = 150
RichColors(5).B = 0

DataFile = App.Path + "\client.dat"
DataFile2 = App.Path + "\recentip.dat"
DataFile3 = App.Path + "\lastconn.dat"

ReDim Commands(1 To 200)
LoadCommands

MDIForm1.Caption = "Server Assistant Client - Copyright 2000 CyberWyre"
MDIForm1.Show

EditFileTemp = App.Path + "\temp1.txt"

MDIForm1.StatusBar1.Panels(1).Text = Ts(App.Major) + "." + Ts(App.Minor) + "." + Ts(App.Revision)

Form1.Show
frmConnect.Show



End Sub

Sub Swap(a As Variant, B As Variant)
Dim c As Variant

c = a
a = B
B = c

End Sub

Function Ts(a) As String
    Ts = Trim(Str(a))
End Function

Function CheckForFile(a$) As Boolean
    B$ = Dir(a$)
    If B$ = "" Then CheckForFile = False
    If B$ <> "" Then CheckForFile = True
End Function

Sub UpdatePlayerList()
'Form6.Show
bbc = -1

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
   
    cc = RGB(RichColors(Players(I).Team + 1).R, RichColors(Players(I).Team + 1).G, RichColors(Players(I).Team + 1).B)
   
    If .Item(j).SubItems(4) <> t$ Then .Item(j).SubItems(4) = t$
    If .Item(j).ListSubItems(4).ForeColor <> cc Then .Item(j).ListSubItems(4).ForeColor = cc
    
    
    If Players(I).Class = 1 Then R$ = "Scout"
    If Players(I).Class = 2 Then R$ = "Sniper"
    If Players(I).Class = 3 Then R$ = "Soldier"
    If Players(I).Class = 4 Then R$ = "Demoman"
    If Players(I).Class = 5 Then R$ = "Medic"
    If Players(I).Class = 6 Then R$ = "HWGuy"
    If Players(I).Class = 7 Then R$ = "Pyro"
    If Players(I).Class = 8 Then R$ = "Spy"
    If Players(I).Class = 9 Then R$ = "Engineer"
    If Players(I).Class = 0 Then R$ = "N/A": img = 1
    If Players(I).Class = -1 Then R$ = "Civilian"
    If Players(I).Class = -2 Then R$ = "Undecided": img = 1
    
    If .Item(j).SubItems(5) <> R$ Then .Item(j).SubItems(5) = R$
    If .Item(j).SubItems(6) <> Players(I).IP Then .Item(j).SubItems(6) = Players(I).IP
    
    If Players(I).Status = True Then R$ = "Connected": img = 2
    If Players(I).Status = False Then R$ = "Playing"
    
    If .Item(j).SubItems(7) <> R$ Then .Item(j).SubItems(7) = R$
    
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
    .Item(j).Tag = "1"
    
Next I

'Form6.ListView1.SortKey = k
'Form6.ListView1.Sorted = True
Form6.ListView1.Sorted = True

'r$ = Vars.Map

If Len(R$) >= 2 Then R$ = UCase(Left(R$, 1)) + LCase(Right(R$, Len(R$) - 1))

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
    Form1.TCP1.Close
End If

If a$ = "HI" Then 'welcome!
    AddEvent "**** Logged in."
    MessBox p$, , "Welcome!"
    PackageConnectPacket
End If

If a$ = "MS" Then 'message
    If MDIForm1.mnuWindowsIn(4).Checked = False Then
        MessBox p$, , "Server Message"
    Else
        AddMsg "----------" + vbCrLf + "Server Message:" + vbCrLf + p$ + vbCrLf + "----------"
    End If
End If

If a$ = "TY" Then 'Add to CONSOLE
    AddMsg p$
End If

End Sub

Sub AddMsg(Txt As String)

'add text to console



Txt = ReplaceString(Txt, vbCrLf, Chr(10))
Txt = ReplaceString(Txt, Chr(10), vbCrLf)


Form1.Text1 = Form1.Text1 + Txt + vbCrLf
If Len(Form1.Text1) > 5000 Then Form1.Text1 = Right(Form1.Text1, 4500)
Form1.Text1.SelStart = Len(Form1.Text1)


End Sub


Public Sub SendPacket(Cde As String, Params As String)

If SendingFile = True Then Exit Sub

a$ = Chr(254) + Chr(254) + Chr(254) + Chr(255) + Cde + Chr(255) + LoginName + Chr(255) + LoginPass + Chr(255) + Params + Chr(255) + Chr(253) + Chr(253) + Chr(253)
If Form1.TCP1.State = sckConnected Then
    'send it in increments of 65000 bytes
    If Len(a$) <= 65000 Then
        Form1.TCP1.SendData a$
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
            
            Form1.TCP1.SendData B$
            DoEvents
        Loop Until Len(B$) < 65000
    End If
End If





End Sub

Public Sub AttemptConnect(IP As String, Port As String, UserName As String, Password As String)

LoginName = UserName
LoginPass = Password
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
            G = h
            h = InStr(G + 1, a$, Chr(250))
            G = G + 1
            j = j + 1
            If G > 0 And h > G Then
                m$ = Mid(a$, G, h - G)
                
                If j = 1 Then SecondsLeft = Val(m$)
                If j = 2 Then MapName = m$
                If j = 3 Then PlayersOn = m$
            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0

UpdateLabel

End Sub

Sub UpdateLabel()

G$ = "Map Time Remaining: "

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

G$ = G$ + c$ + vbCrLf

'map

G$ = G$ + "Current Map: " + MapName + vbCrLf
G$ = G$ + "Users: " + PlayersOn

Form1.lblUpdate = G$

End Sub

Function CheckBit2(BitNum, BitToCheck) As Boolean

Dim a As Long, B As Long
B = BitNum
a = 2 ^ BitToCheck

If (B And a) = a Then CheckBit2 = True

End Function

