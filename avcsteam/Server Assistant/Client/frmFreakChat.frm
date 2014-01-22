VERSION 5.00
Begin VB.Form frmFreakChat 
   Caption         =   "Private Chat"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   6720
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1755
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   6495
      Begin VB.CommandButton Command2 
         Caption         =   "&Beep"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox lblDisturb 
         Caption         =   "Do not beep me."
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Message"
         Height          =   375
         Left            =   1740
         TabIndex        =   10
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox lblMessage 
         Caption         =   "Do not message me."
         Height          =   255
         Left            =   1980
         TabIndex        =   9
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Emote"
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   960
         Width           =   1515
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Away"
         Height          =   375
         Left            =   4920
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox lblNickName 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Text            =   "Me"
         ToolTipText     =   "This nickname is changed and displayed to you only."
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "= Idle Status ="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   180
         Width           =   6375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Your NickName:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblAway 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Person is away."
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   660
         Visible         =   0   'False
         Width           =   2895
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   720
      Top             =   6960
   End
   Begin VB.TextBox lblDisplay 
      BackColor       =   &H80000008&
      ForeColor       =   &H80000005&
      Height          =   3555
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   420
      Width           =   6495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox lblSay 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   5475
   End
   Begin VB.Label lblWho 
      BackStyle       =   0  'Transparent
      Caption         =   "Person"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   60
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Private chat with:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmFreakChat"
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

Public GameDataNum As Integer
Public OppIndex As Integer
Public Opponent As String
Public IsOpponent As Boolean 'when false, means I'm the host
Public GameID As Integer
Public RemoteGameID As Integer
Dim MyScore As Integer


'game specific:
Dim IStarted As Boolean
Dim IAmAway As Boolean
Dim PackTime As String
Dim PackDate As String
Dim AwayCount As Integer

Public Sub StartGame()

If IsOpponent Then
    'tell the other guy i am here
    SendIt "HI", "Ready to chat!"
    SendIt "OC", ""
Else

End If

End Sub

Public Sub GameInterprit(Txt As String, IndexFrom As Integer, WhoFrom As String)

lblWho = WhoFrom

'gets the stuff
'(244)(244)(244)(245)[CODE](245)[PARAMS](245)(243)(243)(243)

e = InStr(1, Txt, Chr(245))
f = InStr(e + 1, Txt, Chr(245))
Debug.Print TheWord
If e > 0 And f > e And f > 0 Then
    'code
    a$ = Mid(Txt, e + 1, f - e - 1)
    e = f
    f = InStr(e + 1, Txt, Chr(245))
    
    If e > 0 And f > e And f > 0 Then
        'params
        p$ = Mid(Txt, e + 1, f - e - 1)
        'decode the encoded shtuff
    End If
End If

'GAME SPECIFIC PACKETS

If a$ = "HI" Then
    IStarted = True
End If

If a$ = "AB" Then
    AddText "~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    AddText "<> Private Chat designed by Freaky, made for Avatar-X's Server Assistant Client"
    AddText "<> Private Chat is Copyright 2001 JSHC Productions"
    AddText "<> Server Assistant Client is Copyright 2001 Cyberwyre"
    AddText "~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
End If

If a$ = "AW" Then
    lblAway = "[ -- " & WhoFrom & " is Away -- ]"
    lblAway.Visible = True
    AddText ">>> " & WhoFrom & " is currently away."
End If

If a$ = "AX" Then
    lblAway.Visible = False
    AddText ">>> " & WhoFrom & " is back in the chat."
End If

If a$ = "BU" Then
    If lblDisturb.Value = 0 Then
        Beep
        AddText "*** " & WhoFrom & " is try to get your attention."
    ElseIf lblDisturb.Value = 1 Then
        SendIt "BF", ""
    End If
End If

If a$ = "BF" Then
    AddText "*** " & WhoFrom & " does not wish to be beeped! No beep was sent."
End If

If a$ = "DT" Then
    AddText "< It is now " & Time & " on " & Date & " >"
End If

If a$ = "EM" Then
    AddText "<> " & WhoFrom & " " & p$
End If

If a$ = "MI" Then
    If lblMessage.Value = 0 Then
        MessBox (p$)
    ElseIf lblMessage.Value = 1 Then
        SendIt "MF", ""
    End If
End If

If a$ = "MF" Then
    AddText "*** " & WhoFrom & " does not wish to get messages! No message was sent."
End If

If a$ = "MA" Then
    AddText "<<< " & p$ & " >>>"
End If

If a$ = "MB" Then
    AddText ">>> IMPORTANT: " & p$ & " <<<"
End If

If a$ = "OC" Then
    AddText "*** " & WhoFrom & " has entered the chat."
    SendIt "SC", ""
End If

If a$ = "SC" Then
    AddText "*** " & WhoFrom & " has entered the chat."
End If

If a$ = "TT" Then
    AddText WhoFrom & ": " & p$
    lblStatus = "New messages have been received!"
End If

If a$ = "PS" Then
    SendIt "PT", ""
    AddText "---> " & WhoFrom & " is verifying his connection to you."
End If

If a$ = "PT" Then
    PackTime = Time
    PackDate = Date
End If


'standard messages
If a$ = "QU" Then
    AddText "*** " & WhoFrom & " has left the chat."
End If

If a$ = "N!" Then
    MessBox "User does not wish to chat!"
    Unload Me
End If

End Sub

Private Sub SendIt(Cde As String, Params As String)

SendGamePacket Cde, Params, OppIndex, GameID, RemoteGameID

End Sub

Sub VerParams()
    If lblSay.Text = "about" Then
        SendIt "AB", ""
        AddText "~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        AddText "<> Private Chat designed by Freaky, made for Avatar-X's Server Assistant Client"
        AddText "<> Private Chat is Copyright 2001 JSHC Productions"
        AddText "<> Server Assistant Client is Copyright 2001 Cyberwyre"
        AddText "~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        lblSay.Text = ""
    ElseIf lblSay.Text = "showtimedate" Then
        SendIt "DT", ""
        AddText "< It is now " & Time & " on " & Date & " >"
        lblSay.Text = ""
    ElseIf lblSay.Text = "logtofile" Then
        pcl$ = App.Path & "\pchatlog.log"
        h = FreeFile
        Open pcl$ For Append As h
            Print #h, Date & " : " & Time
            Print #h, "Private Chat with: " & lblWho
            Print #h, lblDisplay
            Print #h,
            Print #h,
        Close #h
        AddText "*** Chat has successfully been logged to file pchatlog.log."
        lblSay.Text = ""
    ElseIf lblSay.Text = "pingtest" Then
        SendIt "PS", ""
        AddText "---> The last packet received by " & lblWho & " was on " & PackDate & " at " & PackTime & "."
        lblSay.Text = ""
    ElseIf lblSay.Text = "helpme" Then
        AddText "~~~  Private Chat Parameters ~~~"
        AddText "about : Displays Private Chat info."
        AddText "showtimedate : Displays the current time and date to its own user."
        AddText "logtofile : Logs the entire chat into pchatlog.log."
        AddText "pingtest : Checks if connection with opposing user is active. (Check twice)"
        AddText "helpme : Displays help."
        AddText "~~~~~~~~~~~~~~~~~~~~~~~~"
    Else
        If Not lblSay.Text = "" Then
            SendIt "TT", lblSay.Text
            AddText lblNickName & ": " & lblSay.Text
            lblSay.Text = ""
            
        Else
            lblStatus = "You can't send an empty message!"
        End If
    End If

End Sub

Sub CheckAway()
If IAmAway = True Then
    SendIt "AX", ""
    AddText "*** " & lblWho & " has been informed that you are back."
    Me.Caption = "Private Chat - " & lblWho
    IAmAway = False
ElseIf IAmAway = False Then
    SendIt "AW", ""
    AddText "*** " & lblWho & " has been informed that you are away."
    Me.Caption = "Private Chat - " & lblWho & " - Away Mode"
    IAmAway = True
End If

End Sub

Private Sub Command1_Click()
VerParams

End Sub

Private Sub Command2_Click()
SendIt "BU", ""
AddText "*** " & lblWho & " has been beeped."

End Sub

Private Sub Command3_Click()
smb$ = InBox("Enter your message for " & lblWho, "Send Message Box")
SendIt "MI", smb$

End Sub

Private Sub Command4_Click()
eme$ = InBox("Enter your emotion. Start off with a verb (e.g says hello.).", "Send Emotion")
SendIt "EM", eme$
AddText "<> Emote sent: [Your name] " & eme$

End Sub

Private Sub Command5_Click()
CheckAway

End Sub

Private Sub Form_Load()

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
End Sub

Private Sub Form_Resize()

w = Me.Width
h = Me.Height - 520

If Me.WindowState = 1 Then Exit Sub
If Me.WindowState <> 2 Then
    If w < 3000 Then Me.Width = 3000
    If h < 4000 Then Me.Height = 4000
End If

w = Me.Width
h = Me.Height - 520


lblDisplay.Width = w - (lblDisplay.Left * 2) - 160

lblSay.Width = lblDisplay.Width - Command1.Width
Command1.Left = lblSay.Width + 120

Frame1.Top = h - Frame1.Height - 120

lblSay.Top = Frame1.Top - lblSay.Height - 60
Command1.Top = lblSay.Top


lblDisplay.Height = lblSay.Top - lblDisplay.Top - 60



End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next
nm$ = Me.Name
SaveSetting "Server Assistant Client", "Window", nm$ + "winmd", Me.WindowState
SaveSetting "Server Assistant Client", "Window", nm$ + "winh", Me.Height
SaveSetting "Server Assistant Client", "Window", nm$ + "wint", Me.Top
SaveSetting "Server Assistant Client", "Window", nm$ + "winl", Me.Left
SaveSetting "Server Assistant Client", "Window", nm$ + "winw", Me.Width

SendIt "QU", ""
RemoveGameData GameDataNum

End Sub

Private Sub lblSay_Change()
lblStatus = "Conversation in progress..."
AwayCount = 0

End Sub

Private Sub lblSay_KeyPress(KeyAscii As Integer)
If IAmAway = True Then
    CheckAway
End If

If KeyAscii = 13 Then
    KeyAscii = 0
    VerParams
End If

End Sub

Sub AddText(Msg As String)

lblDisplay = lblDisplay + vbCrLf + Msg
lblDisplay.SelStart = Len(lblDisplay)

End Sub

Private Sub Timer2_Timer()
AwayCount = AwayCount + 1
If AwayCount > 20000 Then AwayCount = 20000
If AwayCount >= 90 And IAmAway = False Then
    AddText "*** Activating auto-away mode..."
    IAmAway = False
    CheckAway
End If

End Sub
