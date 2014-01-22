VERSION 5.00
Begin VB.Form frmNewMessage 
   Caption         =   "Compose Message"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   Icon            =   "frmNewMessage.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3930
   ScaleWidth      =   6330
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   60
      TabIndex        =   6
      Top             =   3600
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   315
      Left            =   5280
      TabIndex        =   4
      Top             =   3600
      Width           =   1035
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   420
      Width           =   5475
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   840
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   780
      Width           =   6255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Subject"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   480
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "To:"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "frmNewMessage"
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

Public ReplyText As String
Public ReplyTo As String
Public ReplySubj As String

Private Sub Command1_Click()

Dim MsgData As typMessages

MsgData.MsgFor = Combo1.List(Combo1.ListIndex)
MsgData.MsgSubj = Text2
MsgData.MsgText = Text1

If MsgData.MsgText = "" Then MessBox "Cant send a blank message!": Exit Sub
If MsgData.MsgFor = "" Then MessBox "Must select a recipiant!": Exit Sub
If MsgData.MsgSubj = "" Then MessBox "Please enter a subject!": Exit Sub

SendNewMessage MsgData
Unload Me


End Sub

Private Sub Command2_Click()
If MessBox("Are you sure?", vbYesNo, "Cancel") = vbYes Then Unload Me

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

If Me.WindowState = 1 Then Exit Sub

If Me.WindowState <> 2 Then
    If Me.Width < 3000 Then Me.Width = 3000
    If Me.Height < 3000 Then Me.Height = 3000
End If

w = Me.Width
h = Me.Height
Text1.Width = w - Text1.Left - 120
Text2.Width = w - Text2.Left - 120

Text1.Height = h - Text1.Top - Command2.Height - 480

Command2.Top = Text1.Top + Text1.Height + 60
Command1.Top = Command2.Top
Command1.Left = w - Command1.Width - 120



End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
nm$ = Me.Name
SaveSetting "Server Assistant Client", "Window", nm$ + "winmd", Me.WindowState
SaveSetting "Server Assistant Client", "Window", nm$ + "winh", Me.Height
SaveSetting "Server Assistant Client", "Window", nm$ + "wint", Me.Top
SaveSetting "Server Assistant Client", "Window", nm$ + "winl", Me.Left
SaveSetting "Server Assistant Client", "Window", nm$ + "winw", Me.Width

End Sub
