VERSION 5.00
Begin VB.Form frmAdminChat 
   Caption         =   "Admin Chat"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmAdminChat.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.PictureBox Picture1 
      Height          =   315
      Left            =   4320
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   2400
      Width           =   315
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Text            =   "Chat Here!"
      Top             =   2400
      Width           =   4275
   End
   Begin VB.TextBox Text1 
      Height          =   2355
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4635
   End
End
Attribute VB_Name = "frmAdminChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ShowChat = True

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
h = Me.Height

If Me.WindowState = 1 Then Exit Sub

If h < 2000 Then h = 2000: Me.Height = 2000


Text1.Width = w - Text1.Left - 120
Text2.Width = Text1.Width

Text1.Height = h - Text1.Top - 60 - Text2.Height - 400
Text2.Top = Text1.Top + Text1.Height + 60



End Sub

Private Sub Form_Unload(Cancel As Integer)
ShowChat = False


On Error Resume Next
nm$ = Me.Name
SaveSetting "Server Assistant Client", "Window", nm$ + "winmd", Me.WindowState
SaveSetting "Server Assistant Client", "Window", nm$ + "winh", Me.Height
SaveSetting "Server Assistant Client", "Window", nm$ + "wint", Me.Top
SaveSetting "Server Assistant Client", "Window", nm$ + "winl", Me.Left
SaveSetting "Server Assistant Client", "Window", nm$ + "winw", Me.Width


End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text2_GotFocus()
If Text2 = "Chat Here!" Then Text2 = ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   
    If Text2 = "a" Then
ad:
    Randomize
        b = Int(Rnd * NumSpeech) + 1
 Randomize
        c = Int(Rnd * Speech(b).NumAnswers) + 1
        If Speech(b).NumAnswers = 0 Then GoTo ad
        d$ = Speech(b).Answers(c)
        
        d$ = ReplaceString(d$, "%n", "Freaky")
        d$ = ReplaceString(d$, "%a", "")
        d$ = ReplaceString(d$, "say ", "")
        Text2 = d$
        
    End If

    If Trim(Text2) <> "" Then SendPacket "AC", Text2
    Text2 = ""
    KeyAscii = 0
End If

End Sub
