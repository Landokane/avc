VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form6 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Server Status"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   8550
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Command6 
      Caption         =   "Text"
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   3480
      Width           =   555
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5940
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4080
      Width           =   2595
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   0
      TabIndex        =   7
      Top             =   4080
      Width           =   5235
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Map"
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   3480
      Width           =   555
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1260
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "udp6.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "udp6.frx":059C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "udp6.frx":0B38
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "udp6.frx":10D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "udp6.frx":1670
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "udp6.frx":1C0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "udp6.frx":21A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "udp6.frx":2744
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Player Functions"
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   3480
      Width           =   1635
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3480
      Width           =   1155
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ban"
      Height          =   255
      Left            =   7380
      TabIndex        =   2
      Top             =   3480
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Kick"
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   6059
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   16777215
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3200
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Real Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "UserID"
         Object.Width           =   1295
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "UniqueID"
         Object.Width           =   1826
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Team"
         Object.Width           =   1401
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Class"
         Object.Width           =   2379
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "IP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Time"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " Select a player for stats"
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   3780
      Width           =   8535
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form6"
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
' ===========================================================================
' ---------------------------------------------------------------------------

Dim ListPerc As Double
Dim Text1Perc As Double
Dim Text2Perc As Double
Dim TextMode As Boolean

Dim Dragger As Integer
Public OldWindowProc As Long

Public Sub Functions(Index As Integer)

On Error Resume Next

'If Index <> 2 Then MsgBox Index

If Index = 31 Then
    a$ = ListView1.SelectedItem
    If a$ = "" Then Exit Sub
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(i).Text = a$ Then j = i: Exit For
    Next i
    
    b$ = ListView1.ListItems.Item(j).SubItems(3) 'uniqueid
    
    'get real name
    nn$ = InBox("Enter the REAL name for player " + a$, "Add RealPlayer", a$)
    If nn$ = "" Then Exit Sub
    
    AddRealPlayer nn$, b$

ElseIf Index = 1 Then
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected = True Then
            
            a$ = ListView1.ListItems(i).Text
            b$ = ListView1.ListItems(i).SubItems(3) 'uniqueid
            
            'get real name
            AddRealPlayer a$, b$
        End If
    Next i
    
ElseIf Index = 36 Then
    'kick
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected = True Then
        
            b$ = Trim(ListView1.ListItems.Item(i).SubItems(2)) 'userid
            'send kick command
            SendPacket "CA", b$
            
        End If
    Next i
    
ElseIf Index = 4 Then
    'kill
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected = True Then
        
            b$ = Trim(ListView1.ListItems.Item(i).SubItems(2)) 'userid
            SendPacket "RC", "kill " + b$
            
        End If
    Next i
    
ElseIf Index = 5 Then
    'kill
    
    a$ = ListView1.SelectedItem
    
    If a$ = "" Then Exit Sub
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(i).Text = a$ Then j = i: Exit For
    Next i
    b$ = Trim(ListView1.ListItems.Item(j).SubItems(2)) 'userid
    
    nn1$ = InBox("Enter new name:", "Change Player Name", a$)
    If nn1$ = "" Then Exit Sub
    
    SendPacket "RC", "changename " + b$ + " " + nn1$

ElseIf Index = 34 Then
    'kill
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected = True Then
    
            b$ = Trim(ListView1.ListItems.Item(i).SubItems(2)) 'userid
        
            SendPacket "RC", "setreal # " + b$
        End If
    Next
    
ElseIf Index = 38 Then
    'kill
    
    a$ = ListView1.SelectedItem
    
    If a$ = "" Then Exit Sub
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(i).Text = a$ Then j = i: Exit For
    Next i
    b$ = Trim(ListView1.ListItems.Item(j).SubItems(2)) 'userid
    
    SendPacket "RC", "annid # " + b$

ElseIf Index = 10 Then
    'kill
    nn1$ = InBox("Enter minutes to ban for:", "Ban Player " + a$ + " for X Minutes", "30")
    If nn1$ = "" Then Exit Sub
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected = True Then
            
            b$ = ListView1.ListItems.Item(i).SubItems(3) 'unique
        
        
            SendPacket "RC", "banid " + nn1$ + " " + b$ + " kick"
        End If
    Next
    
ElseIf Index = 33 Then
    'kill
    
    a$ = ListView1.SelectedItem
    
    If a$ = "" Then Exit Sub
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(i).Text = a$ Then j = i: Exit For
    Next i
    b$ = ListView1.ListItems.Item(j).SubItems(3) 'unique
    
    FindReal = b$
    SendPacket "RP", ""

ElseIf Index = 11 Then
    'kill
    nn1$ = InBox("Enter Private Message:", "Send Private Message", "")
    If nn1$ = "" Then Exit Sub
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected = True Then
    
    
            b$ = Trim(ListView1.ListItems.Item(i).SubItems(2)) 'userid
            
        
            SendPacket "RC", "talkto " + b$ + " " + nn1$
        End If
    Next i
    
ElseIf Index = 60 Then
    'kill
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected = True Then
    
    
            b$ = Trim(ListView1.ListItems.Item(i).SubItems(2)) 'userid
            SendPacket "RC", "devoice # " + b$
        End If
    Next i
        
    
ElseIf Index = 61 Then
    'kill
    
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected = True Then
             b$ = Trim(ListView1.ListItems.Item(i).SubItems(2)) 'userid
            
             SendPacket "RC", "revoice # " + b$
        End If
    Next i
ElseIf Index = 13 Then
    'kill
    
    a$ = ListView1.SelectedItem
   
    If a$ = "" Then Exit Sub
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(i).Text = a$ Then j = i: Exit For
    Next i
    b$ = Trim(ListView1.ListItems.Item(j).SubItems(2)) 'userid
    
    frmMap.RotPlayerNum = Val(b$)
    frmMap.Update2
    
ElseIf Index = 51 Then
    'kill
    
    a$ = ListView1.SelectedItem
    If a$ = "" Then Exit Sub
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(i).Text = a$ Then j = i: Exit For
    Next i
    b$ = Trim(ListView1.ListItems.Item(j).SubItems(2)) 'userid
    
    nn1$ = InBox("Add how many points to " + a$ + " ?", "Add/Subtract Points", "10")
    If nn1$ = "" Then Exit Sub
    
    SendPacket "RC", "addpoints " + b$ + " " + nn1$

ElseIf Index = 52 Then
    'kill
    
    a$ = ListView1.SelectedItem
    If a$ = "" Then Exit Sub
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(i).Text = a$ Then j = i: Exit For
    Next i
    b$ = Trim(ListView1.ListItems.Item(j).SubItems(2)) 'userid
    
    SendPacket "RC", "annpoints " + b$

ElseIf Index = 32 Then
    'kill
    
    a$ = ListView1.SelectedItem
    If a$ = "" Then Exit Sub
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(i).Text = a$ Then j = i: Exit For
    Next i
    b$ = Trim(ListView1.ListItems.Item(j).SubItems(2)) 'userid
    
    SendPacket "RC", "addentry " + a$

End If





End Sub


Public Sub FunctionsClass(Index As Integer)

On Error Resume Next

'change class

a$ = ListView1.SelectedItem
If a$ = "" Then Exit Sub

For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems.Item(i).Text = a$ Then j = i: Exit For
Next i
b$ = Trim(ListView1.ListItems.Item(j).SubItems(2)) 'userid


cl$ = Ts(Index + 1)
If Index = 9 Then cl$ = "11"

SendPacket "RC", "changeclass " + b$ + " " + cl$

End Sub

Private Sub Command1_Click()

'kick
On Error Resume Next
   
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected = True Then
        
            b$ = ListView1.ListItems.Item(i).SubItems(2) 'userid
        
            'send kick command
            SendPacket "SK", b$
        End If
    Next
End Sub

Private Sub Command2_Click()

Me.PopupMenu MDIForm1.mnuBanTime

End Sub

Private Sub Command3_Click()

SendPacket "SU", ""

End Sub

Private Sub Command4_Click()

MDIForm1.PopupMenu MDIForm1.mnuFunctions

End Sub

Private Sub Command5_Click()

frmMap.Show

End Sub

Private Sub Command6_Click()

TextMode = Not TextMode

Form_Resize



End Sub

Private Sub Form_Load()
AddForm True, 133, 166, 0, 0, Me

ListView1.BackColor = RGB(RichColors(9).r, RichColors(9).g, RichColors(9).b)
ListPerc = 0.3
Text1Perc = 0.5
Text2Perc = 0.2
Me.Width = 10500
ShowPlayers = True
UpdatePlayerList
UpdateLogDetail

Command5.Enabled = DllEnabled

FillCombo

On Error Resume Next
nm$ = Me.Name
winash = Val(GetSetting("Server Assistant Client", "Window", nm$ + "winash", -1))
If winash <> -1 Then Me.Show
winmd = Val(GetSetting("Server Assistant Client", "Window", nm$ + "winmd", 3))
If winmd <> 3 Then Me.WindowState = winmd
If Me.WindowState = 0 Then
    winh = GetSetting("Server Assistant Client", "Window", nm$ + "winh", -1)
    wint = GetSetting("Server Assistant Client", "Window", nm$ + "wint", -1)
    winl = GetSetting("Server Assistant Client", "Window", nm$ + "winl", -1)
    winw = GetSetting("Server Assistant Client", "Window", nm$ + "winw", -1)
    
    If winh <> -1 Then Me.Height = winh
    If wint <> -1 Then Me.Top = wint
    If winl <> -1 Then Me.Left = winl
    If winw <> -1 Then Me.Width = winw
End If



n = ListView1.ColumnHeaders.Count

For i = 1 To n
    wp = GetSetting("Server Assistant Client", "PlayerList", ListView1.ColumnHeaders(i).Text + "pos", -1)
    If wp <> -1 Then ListView1.ColumnHeaders(i).Position = wp
    wp = GetSetting("Server Assistant Client", "PlayerList", ListView1.ColumnHeaders(i).Text + "wid", -1)
    If wp <> -1 Then ListView1.ColumnHeaders(i).Width = wp
Next i


End Sub

Public Sub FillCombo()

Combo1.Clear
For i = 1 To NumMenuScripts
    
    nm$ = MenuScripts(i).MenuName
    If Len(MenuScripts(i).AskForQuestion) > 0 Then
        Combo1.AddItem nm$
        Combo1.ItemData(Combo1.NewIndex) = i
    End If

Next i

If Combo1.ListCount > 0 Then Combo1.ListIndex = 0

End Sub


Sub Update()

If Me.WindowState = 1 Then Exit Sub

If Me.Width < 3705 Then Me.Width = 3705
Label1.Width = Me.Width - 120

h = Me.Height
h = h - Command1.Height - ListView1.Top - 420 - Label1.Height - 60

If TextMode Then
    h = h - 60 - Text1.Height
End If

ListView1.Height = h '* ListPerc
'Text1.Top = ListView1.Height + 45

'Text1.Height = h * Text1Perc
'Text2.Top = Text1.Height + 45 + Text1.Top
'Text2.Height = h * Text2Perc
Command1.Top = h + 60
Command2.Top = h + 60
Command3.Top = h + 60
Command4.Top = h + 60
Command5.Top = h + 60
Command6.Top = h + 60

If TextMode = False Then hj = 500


Label1.Top = h + Command1.Height + 120

Text1.Top = Label1.Top + 60 + Label1.Height + hj
Combo1.Top = Text1.Top


Command6.Left = Me.Width - Command1.Width - Command2.Width - Command4.Width - Command5.Width - Command6.Width - 360
Command5.Left = Me.Width - Command1.Width - Command2.Width - Command4.Width - Command5.Width - 300
Command4.Left = Me.Width - Command1.Width - Command2.Width - Command4.Width - 240
Command1.Left = Me.Width - Command1.Width - Command2.Width - 180
Command2.Left = Me.Width - Command2.Width - 120
Command3.Left = 60

Text1.Width = Me.Width - Combo1.Width - 240
Combo1.Left = Text1.Left + Text1.Width + 60

End Sub

Private Sub Form_Resize()

If Me.Width < 2000 Then Me.Width = 2000
If Me.Height < 2500 Then Me.Height = 2500

Update

ListView1.Width = Me.Width - 120

'Text1.Width = Me.Width - 120
'Text2.Width = Me.Width - 120

'Label1.Top = ListView1.Height - (Label1.Height / 2)
'Label2.Top = ListView1.Height - (Label1.Height / 2)





End Sub

Private Sub Form_Unload(Cancel As Integer)
ShowPlayers = False
UpdateLogDetail


On Error Resume Next
nm$ = Me.Name
SaveSetting "Server Assistant Client", "Window", nm$ + "winmd", Me.WindowState
SaveSetting "Server Assistant Client", "Window", nm$ + "winh", Me.Height
SaveSetting "Server Assistant Client", "Window", nm$ + "wint", Me.Top
SaveSetting "Server Assistant Client", "Window", nm$ + "winl", Me.Left
SaveSetting "Server Assistant Client", "Window", nm$ + "winw", Me.Width

n = ListView1.ColumnHeaders.Count

For i = 1 To n
    SaveSetting "Server Assistant Client", "PlayerList", ListView1.ColumnHeaders(i).Text + "pos", Ts(ListView1.ColumnHeaders(i).Position)
    SaveSetting "Server Assistant Client", "PlayerList", ListView1.ColumnHeaders(i).Text + "wid", Ts(ListView1.ColumnHeaders(i).Width)
Next i


End Sub

Private Sub ShowPlayer(b)


    a$ = "Entry Name: " + Players(b).EntryName
    a$ = a$ + "     Points: " + Ts(Players(b).Points)
    a$ = a$ + "     Port: " + Ts(Players(b).Port)
    If Players(b).ShutUp Then a$ = a$ + "     IS DEVOICED!"
    If Players(b).Warn Then a$ = a$ + "     IS WARNED!"
    a$ = a$ + "     Kick Votes Started: " + Ts(Players(b).NumKickVotes)
    
    
    sec$ = Ts(Second(Players(b).LastEvent))
    mn$ = Ts(Minute(Players(b).LastEvent))
    hr$ = Ts(Hour(Players(b).LastEvent))
    
    If Len(hr$) = 1 Then hr$ = "0" + hr$
    If Len(sec$) = 1 Then sec$ = "0" + sec$
    If Len(mn$) = 1 Then mn$ = "0" + mn$
    hr$ = hr$ + ":" + mn$ + ":" + sec$

    a$ = a$ + "      Idle for: " + hr$
    
    Form6.Label1 = a$

End Sub

Private Sub ListView1_Click()

'show info
On Error Resume Next

For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems.Item(i).Selected = True Then j = i: Exit For
Next i


If j > 0 Then b = Val(Trim(ListView1.ListItems.Item(j).Tag)) 'userid


If b > 0 Then
    ShowPlayer b
End If

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

'istview1.ColumnHeaders(ListView1.SortKey + 1).for

ListView1.Sorted = True
k = ListView1.SortKey

If k = (ColumnHeader.Index - 1) Then
    If ListView1.SortOrder = lvwDescending Then
        ListView1.SortOrder = lvwAscending
    Else
        ListView1.SortOrder = lvwDescending
    End If
End If

ListView1.SortKey = (ColumnHeader.Index - 1)






End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)


'show info
On Error Resume Next

For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems.Item(i).Selected = True Then j = i: Exit For
Next i
b = Val(Trim(ListView1.ListItems.Item(j).Tag)) 'userid


If b > 0 Then
    ShowPlayer b
End If


End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then MDIForm1.PopupMenu MDIForm1.mnuFunctions

End Sub

Private Sub mnuBanTimeIn_Click(Index As Integer)


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If Trim(Text1) <> "" Then
    
        d = Combo1.ListIndex
        If d = -1 Then Exit Sub
        e = Combo1.ItemData(d)
        
    
        'userid:
        b$ = "0"
        
        For i = 1 To Form6.ListView1.ListItems.Count
            If Form6.ListView1.ListItems(i).Selected = True Then
                b$ = Form6.ListView1.ListItems(i).SubItems(2)
        
                a$ = ""
                a$ = a$ + Chr(251)
                a$ = a$ + MenuScripts(e).ScriptName + Chr(250)
                a$ = a$ + b$ + Chr(250)
                a$ = a$ + Trim(Text1) + Chr(250)
                a$ = a$ + Chr(251)
                
                SendPacket "SS", a$
            End If
        Next i
    
        KeyAscii = 0
        Text1 = ""
    
    End If
    
End If


End Sub

