VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
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
   Begin VB.Frame Frame1 
      Caption         =   "Chat Color"
      Height          =   1815
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CommandButton Command1 
         Caption         =   "Set"
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   1500
         Width           =   615
      End
      Begin VB.PictureBox Picture2 
         Height          =   1095
         Left            =   2220
         ScaleHeight     =   1035
         ScaleWidth      =   975
         TabIndex        =   11
         Top             =   360
         Width           =   1035
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   255
         Left            =   300
         TabIndex        =   5
         Top             =   720
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   450
         _Version        =   393216
         Max             =   255
         TickFrequency   =   16
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   300
         TabIndex        =   7
         Top             =   300
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   450
         _Version        =   393216
         Max             =   255
         TickFrequency   =   16
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   255
         Left            =   300
         TabIndex        =   9
         Top             =   1140
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   450
         _Version        =   393216
         Max             =   255
         TickFrequency   =   16
      End
      Begin VB.Label Label3 
         Caption         =   "B"
         Height          =   195
         Left            =   60
         TabIndex        =   10
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "R"
         Height          =   195
         Left            =   60
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "G"
         Height          =   195
         Left            =   60
         TabIndex        =   6
         Top             =   780
         Width           =   255
      End
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   3413
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAdminChat.frx":08CA
   End
   Begin VB.PictureBox Picture1 
      Height          =   315
      Left            =   4320
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   2400
      Width           =   315
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Text            =   "Chat Here!"
      Top             =   2400
      Width           =   4275
   End
   Begin RichTextLib.RichTextBox RT2 
      Height          =   735
      Left            =   5040
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      TextRTF         =   $"frmAdminChat.frx":0984
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAdminChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChatCol As Long

Sub ShowCol(Optional nosetslide As Boolean)

If nosetslide Then
    ChatCol = RGB(Slider1, Slider2, Slider3)
End If

Picture1.BackColor = ChatCol
Picture2.BackColor = ChatCol

c = ChatCol

r = c Mod 256
g = (c \ 256) Mod 256
b = c \ 256 \ 256

If Not nosetslide Then Slider1.Value = r
If Not nosetslide Then Slider2.Value = g
If Not nosetslide Then Slider3.Value = b

End Sub

Private Sub Command1_Click()

If Frame1.Visible = True Then
    Frame1.Visible = False
    Text1.Visible = True
Else
    Frame1.Visible = True
    Text1.Visible = False
End If

End Sub

Private Sub Command2_Click()

'Text1.OLEObjects.Add , , App.Path & "\avpic.bmp"

End Sub

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


ChatCol = GetSetting("Server Assistant Client", "Window", "adminchatcol", RGB(200, 200, 200))
ShowCol
End Sub

Private Sub Form_Resize()

w = Me.Width
h = Me.Height

If Me.WindowState = 1 Then Exit Sub

If h < 2000 Then h = 2000: Me.Height = 2000


Text1.Width = w - Text1.Left - 120
Text2.Width = Text1.Width - Picture1.Width - 60
Picture1.Left = Text2.Width + Text2.Left + 60
Picture1.Top = Text2.Top


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
SaveSetting "Server Assistant Client", "Window", "adminchatcol", ChatCol


End Sub

Private Sub Picture1_Click()

If Frame1.Visible = True Then
    Frame1.Visible = False
    Text1.Visible = True
Else
    Frame1.Visible = True
    Text1.Visible = False
End If


End Sub

Private Sub Slider1_Change()
ShowCol True
End Sub

Private Sub Slider1_Scroll()

ShowCol True

End Sub

Private Sub Slider2_Change()
ShowCol True
End Sub

Private Sub Slider2_Scroll()
ShowCol True
End Sub

Private Sub Slider3_Change()
ShowCol True
End Sub

Private Sub Slider3_Scroll()
ShowCol True
End Sub

Private Sub Text2_GotFocus()
If Text2 = "Chat Here!" Then Text2 = ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   If Trim(Text2) <> "" Then
        
        'send chat
        
        a$ = Chr(251)
        a$ = a$ + Text2 + Chr(250)
        a$ = a$ + Ts(ChatCol) + Chr(250)
        a$ = a$ + Chr(251)
    
        SendPacket "AC", a$
    End If
    Text2 = ""
    KeyAscii = 0
End If

End Sub
