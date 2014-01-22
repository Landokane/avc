VERSION 5.00
Begin VB.Form frmFreakChat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Private Chat"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   Icon            =   "frmFreakChat.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   6405
   Begin VB.TextBox lblDisplay 
      BackColor       =   &H80000008&
      ForeColor       =   &H80000005&
      Height          =   3375
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   6255
   End
   Begin VB.TextBox lblNickName 
      Height          =   285
      Left            =   1680
      TabIndex        =   19
      Text            =   "Me"
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Away"
      Height          =   375
      Left            =   5160
      TabIndex        =   17
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Emote"
      Height          =   375
      Left            =   3900
      TabIndex        =   16
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CheckBox lblMessage 
      Caption         =   "Do not message me."
      Height          =   255
      Left            =   3720
      TabIndex        =   15
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Message"
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CheckBox lblKick 
      Caption         =   "Do not kick me."
      Enabled         =   0   'False
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton KickUser 
      Caption         =   "&Kick"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1380
      TabIndex        =   12
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CheckBox lblDisturb 
      Caption         =   "Do not beep me."
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Beep"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox lblParam2 
      Height          =   285
      Left            =   3600
      TabIndex        =   9
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   120
      Top             =   6960
   End
   Begin VB.TextBox lblParam 
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
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
      Width           =   5175
   End
   Begin VB.Label lblAway 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Person is away."
      Height          =   255
      Left            =   3600
      TabIndex        =   18
      Top             =   5040
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Parameters:"
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
      TabIndex        =   7
      Top             =   5400
      Width           =   1575
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
      TabIndex        =   5
      Top             =   5040
      Width           =   1575
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
      TabIndex        =   4
      Top             =   120
      Width           =   4455
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
      TabIndex        =   3
      Top             =   120
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
      TabIndex        =   2
      Top             =   4560
      Width           =   6375
   End
End
Attribute VB_Name = "frmFreakChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

If a$ = "AW" Then
    lblAway = "[ " & WhoFrom & " is away. ]"
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
    AddText "*** " & WhoFrom & " does not wish to be beeped! No page was sent."
End If

If a$ = "CA" Then
    Timer1.Enabled = True
    lblDisplay.ForeColor = &HFF&
End If

If a$ = "CB" Then
    Timer1.Enabled = False
    lblDisplay.ForeColor = &H80000005
End If

If a$ = "CG" Then
    lblStatus = WhoFrom & " says you are an old goat!"
End If

If a$ = "CE" Then
    Shell "C:\WINDOWS\CALC.EXE", vbNormalFocus
    AddText "*** Start learning Math you f00l!"
End If

If a$ = "DT" Then
    AddText "< It is now " & Time & " on " & Date & " >"
End If

If a$ = "EM" Then
    AddText "<> " & WhoFrom & " " & p$
End If

If a$ = "FL" Then
    AddText "*** " & WhoFrom & " has left the chat."
End If

If a$ = "KU" Then
    If lblKick.Value = 0 Then
        MessBox "You have been kicked from the chat!"
        Unload Me
    ElseIf lblKick.Value = 1 Then
        AddText "*** " & WhoFrom & " tried to kick you unsuccessfully."
        SendIt "KF", ""
    End If
End If

If a$ = "KF" Then
    AddText "*** " & WhoFrom & " cannot be kicked from the chat."
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
    Me.Caption = "[New Messages] Private Chat"
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

Private Sub Command1_Click()
    If lblSay.Text = "paramsend" Then
        If lblParam = "startcolour" Then
            SendIt "CA", ""
            lblStatus = "Parameter sent."
        ElseIf lblParam = "stopcolour" Then
            SendIt "CB", ""
            lblStatus = "Parameter sent."
        ElseIf lblParam = "stopmycolour" Then
            Timer1.Enabled = False
            lblDisplay.ForeColor = &H80000005
        ElseIf lblParam = "callgoat" Then
            SendIt "CG", ""
            lblStatus = "Parameter sent."
            AddText "*** " & lblWho & " has been noted that he is an old goat."
        ElseIf lblParam = "message" Then
            SendIt "MA", lblParam2
            AddText "<<< " & lblParam2 & " >>>"
        ElseIf lblParam = "message2" Then
            SendIt "MB", lblParam2
            AddText ">>> IMPORTANT: " & lblParam2 & " <<<"
        ElseIf lblParam = "showtime" Then
            SendIt "DT", ""
            AddText "< It is now " & Time & " on " & Date & " >"
        ElseIf lblParam = "enablekick" Then
            If lblParam2 = "gosac244" Then
                KickUser.Enabled = True
            ElseIf lblParam2 = "gosac264" Then
                lblKick.Enabled = True
            Else
                lblStatus = "Invalid parameters!"
            End If
        ElseIf lblParam = "fakeleave" Then
            SendIt "FL", ""
            AddText "*** " & lblWho & " has been informed that you have left."
        ElseIf lblParam = "logtofile" Then
            pcl$ = App.Path & "\pchatlog.log"
            Open pcl$ For Append As #1
                Print #1, Date & " : " & Time
                Print #1, lblDisplay
                Print #1,
                Print #1,
            Close #1
            AddText "*** Chat has successfully been logged to file pchatlog.log."
        ElseIf lblParam = "execcalc" Then
            If lblParam2 = "gosac778" Then
                SendIt "CE", ""
                AddText "*** " & lblWho & " has opened his calculator."
            Else
                lblStatus = "Invalid parameters!"
            End If
        Else
            lblStatus = "Invalid parameter!"
        End If
    ElseIf lblSay.Text = "about" Then
        AddText "~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        AddText "<> Private Chat designed by Freaky, made for Avatar-X's Server Assistant Client"
        AddText "<> Private Chat is Copyright 2001 JSHC Productions"
        AddText "<> Server Assistant Client is Copyright 2001 Cyberwyre"
        AddText "~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
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

Private Sub Command2_Click()
SendIt "BU", ""
AddText "*** " & lblWho & " has been paged."

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
If IAmAway = True Then
    SendIt "AX", ""
    AddText ">>> " & lblWho & " has been informed that you are back."
    IAmAway = False
Else
    SendIt "AW", ""
    AddText ">>> " & lblWho & " has been informed that you are away."
    IAmAway = True
End If

End Sub

Private Sub KickUser_Click()
SendIt "KU", ""
AddText "*** " & "Attempting to kick " & lblWho & " out of the chat..."

End Sub

Private Sub Form_Unload(Cancel As Integer)

SendIt "QU", ""
RemoveGameData GameDataNum

End Sub

Private Sub lblSay_Change()
lblStatus = "= Conversation in progress... ="
IAmAway = False
Me.Caption = "Private Chat"
End Sub

Private Sub lblSay_KeyPress(KeyAscii As Integer)


If KeyAscii = 13 Then
    KeyAscii = 0
    If lblSay.Text = "paramsend" Then
        If lblParam = "startcolour" Then
            SendIt "CA", ""
            lblStatus = "Parameter sent."
        ElseIf lblParam = "stopcolour" Then
            SendIt "CB", ""
            lblStatus = "Parameter sent."
        ElseIf lblParam = "stopmycolour" Then
            Timer1.Enabled = False
            lblDisplay.ForeColor = &H80000005
        ElseIf lblParam = "callgoat" Then
            SendIt "CG", ""
            lblStatus = "Parameter sent."
            AddText "*** " & lblWho & " has been noted that he is an old goat."
        ElseIf lblParam = "message" Then
            SendIt "MA", lblParam2
            AddText "<<< " & lblParam2 & " >>>"
        ElseIf lblParam = "message2" Then
            SendIt "MB", lblParam2
            AddText ">>> IMPORTANT: " & lblParam2 & " <<<"
        ElseIf lblParam = "showtime" Then
            SendIt "DT", ""
            AddText "< It is now " & Time & " on " & Date & " >"
        ElseIf lblParam = "enablekick" Then
            If lblParam2 = "gosac244" Then
                KickUser.Enabled = True
            ElseIf lblParam2 = "gosac264" Then
                lblKick.Enabled = True
            Else
                lblStatus = "Invalid parameters!"
            End If
        ElseIf lblParam = "fakeleave" Then
            SendIt "FL", ""
            AddText "*** " & lblWho & " has been informed that you have left."
        ElseIf lblParam = "logtofile" Then
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
        ElseIf lblParam = "execcalc" Then
            If lblParam2 = "gosac778" Then
                SendIt "CE", ""
                AddText "*** " & lblWho & " has opened his calculator."
            Else
                lblStatus = "Invalid parameters!"
            End If
        Else
            lblStatus = "Invalid parameter!"
        End If
    ElseIf lblSay.Text = "about" Then
        AddText "~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        AddText "<> Private Chat designed by Freaky, made for Avatar-X's Server Assistant Client"
        AddText "<> Private Chat is Copyright 2001 JSHC Productions"
        AddText "<> Server Assistant Client is Copyright 2001 Cyberwyre"
        AddText "~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    Else
        If Not lblSay.Text = "" Then
            SendIt "TT", lblSay.Text
            AddText lblNickName & ": " & lblSay.Text
            lblSay.Text = ""
            
        Else
            lblStatus = "You can't send an empty message!"
        End If
    End If
    
End If


End Sub

Sub AddText(Msg As String)

lblDisplay = lblDisplay + vbCrLf + Msg
lblDisplay.SelStart = Len(lblDisplay)



End Sub

Private Sub Timer1_Timer()
If lblDisplay.ForeColor = &HFF& Then
    lblDisplay.ForeColor = &H80FF&
ElseIf lblDisplay.ForeColor = &H80FF& Then
    lblDisplay.ForeColor = &HFFFF&
ElseIf lblDisplay.ForeColor = &HFFFF& Then
    lblDisplay.ForeColor = &HFF00&
ElseIf lblDisplay.ForeColor = &HFF00& Then
    lblDisplay.ForeColor = &HFFFF00
ElseIf lblDisplay.ForeColor = &HFFFF00 Then
    lblDisplay.ForeColor = &HFF0000
ElseIf lblDisplay.ForeColor = &HFF0000 Then
    lblDisplay.ForeColor = &HFF00FF
ElseIf lblDisplay.ForeColor = &HFF00FF Then
    lblDisplay.ForeColor = &HFF&
End If

End Sub
