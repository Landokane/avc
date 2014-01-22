VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Setup"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   Icon            =   "udp2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text9 
      Height          =   315
      Left            =   1560
      TabIndex        =   23
      Text            =   "127.0.0.1"
      Top             =   1800
      Width           =   1995
   End
   Begin VB.TextBox Text8 
      Height          =   315
      Left            =   1560
      TabIndex        =   21
      Text            =   "avatar"
      Top             =   3960
      Width           =   1995
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reset Defaults"
      Height          =   435
      Left            =   4080
      TabIndex        =   20
      Top             =   3900
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   555
      Left            =   2940
      TabIndex        =   19
      Top             =   4380
      Width           =   2715
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   555
      Left            =   60
      TabIndex        =   18
      Top             =   4380
      Width           =   2775
   End
   Begin VB.TextBox Text7 
      Height          =   315
      Left            =   1560
      TabIndex        =   16
      Text            =   "25005"
      Top             =   2880
      Width           =   1995
   End
   Begin VB.TextBox Text6 
      Height          =   315
      Left            =   1560
      TabIndex        =   14
      Text            =   "26001"
      Top             =   3600
      Width           =   1995
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   1560
      TabIndex        =   12
      Text            =   "26000"
      Top             =   3240
      Width           =   1995
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   1560
      TabIndex        =   10
      Text            =   "25000"
      Top             =   2520
      Width           =   1995
   End
   Begin VB.OptionButton Option2 
      Caption         =   "I am connecting to the server remotely"
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   360
      Width           =   5535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "I am running Server Assistant on a server machine"
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Value           =   -1  'True
      Width           =   5535
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Text            =   "127.0.0.1"
      Top             =   1440
      Width           =   1995
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Text            =   "27015"
      Top             =   2160
      Width           =   1995
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5100
      Top             =   2700
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   4035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Text            =   "c:\sierra\half-life"
      Top             =   720
      Width           =   4035
   End
   Begin VB.Label Label10 
      Caption         =   "Log IP"
      Height          =   195
      Left            =   60
      TabIndex        =   24
      Top             =   1860
      Width           =   1035
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "File Password"
      Height          =   195
      Left            =   60
      TabIndex        =   22
      Top             =   4020
      Width           =   975
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Log Feedback Port"
      Height          =   195
      Left            =   60
      TabIndex        =   17
      Top             =   2940
      Width           =   1365
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Send File Port"
      Height          =   195
      Left            =   60
      TabIndex        =   15
      Top             =   3660
      Width           =   990
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Recieve File port"
      Height          =   195
      Left            =   60
      TabIndex        =   13
      Top             =   3300
      Width           =   1200
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Log Recieve Port"
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   2580
      Width           =   1245
   End
   Begin VB.Label Label4 
      Caption         =   "Server IP"
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   1500
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Game Server Port"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   2220
      Width           =   1260
   End
   Begin VB.Label Label2 
      Caption         =   "Game Directory"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   1140
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Half-Life Path"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   780
      Width           =   1035
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TimerVal As Integer
Dim TextPos As Integer


Private Sub RefreshCombo()
Combo1.Clear
a$ = Text1
a$ = Trim(a$)
If Right(a$, 1) = "\" Then a$ = Left(a$, Len(a$) - 1)
Text1 = a$
DoEvents
TimerVal = -1
b$ = Dir(a$, vbDirectory)

If b$ = "" Then Exit Sub

a$ = a$ + "\"
i = 0
b$ = Dir(a$, vbDirectory)
Do While b$ <> ""
   If b$ <> "." And b$ <> ".." Then
      
      If (GetAttr(a$ & b$) And vbDirectory) = vbDirectory Then
         b$ = LCase(b$)
         If b$ <> "auth" And b$ <> "gldrv" And b$ <> "logos" Then
             Combo1.AddItem b$
             If b$ = "tfc" Then i = Combo1.NewIndex
         End If
      End If
    End If
   b$ = Dir
Loop

If Combo1.ListCount > 0 Then Combo1.ListIndex = i
End Sub


Private Sub Command1_Click()
Server.HLPath = Text1
Server.GamePath = Combo1.List(Combo1.ListIndex)
Server.Port = Ts(Val(Text2))
Server.IP = Text3
Server.LocalLogPort = Ts(Val(Text4))
Server.LocalFilePort = Ts(Val(Text5))
Server.RemoteFilePort = Ts(Val(Text6))
Server.LogFeedbackPort = Ts(Val(Text7))
Server.Password = Text8
Server.LogIP = Text9
If Option1.Value = True Then Server.WhatIs = 0
If Option2.Value = True Then Server.WhatIs = 1
SaveCommands
If Reload = True Then Main

Unload Form2

End Sub

Private Sub Command3_Click()
Defaults
Text2 = Server.Port
Text3 = Server.IP
Text4 = Server.LocalLogPort
Text5 = Server.LocalFilePort
Text6 = Server.RemoteFilePort
Text7 = Server.LogFeedbackPort
Text8 = Server.Password
Text9 = Server.LogIP
If Server.WhatIs = 0 Then Option1.Value = True
If Server.WhatIs = 1 Then Option2.Value = True
End Sub

Private Sub Form_Load()

Text1 = Server.HLPath
RefreshCombo
a$ = Server.GamePath
For i = 0 To Combo1.ListCount - 1
    If UCase(Combo1.List(i)) = UCase(a$) Then j = i
Next i
Combo1.ListIndex = j
Text2 = Server.Port
Text3 = Server.IP
Text4 = Server.LocalLogPort
Text5 = Server.LocalFilePort
Text6 = Server.RemoteFilePort
Text7 = Server.LogFeedbackPort
Text8 = Server.Password
Text9 = Server.LogIP
If Server.WhatIs = 0 Then Option1.Value = True
If Server.WhatIs = 1 Then Option2.Value = True
End Sub

Private Sub Text1_Change()
TimerVal = 3
TextPos = Text1.SelStart
End Sub

Private Sub Timer1_Timer()
If TimerVal > 0 Then TimerVal = TimerVal - 1
If TimerVal = 0 Then
    tt = TextPos
    RefreshCombo
    TimerVal = -1
    TextPos = tt
    If TextPos > Len(Text1) Then
        Text1.SelStart = Len(Text1)
    Else
        Text1.SelStart = TextPos
    End If
End If
End Sub
