VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Chat Window"
   ClientHeight    =   5055
   ClientLeft      =   9690
   ClientTop       =   2865
   ClientWidth     =   10245
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   10245
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   4500
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   315
      Left            =   5880
      TabIndex        =   3
      Top             =   0
      Width           =   795
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Rcon Command Box"
      Top             =   0
      Width           =   5835
   End
   Begin VB.CommandButton Command8 
      Caption         =   "C"
      Height          =   315
      Left            =   6720
      TabIndex        =   1
      ToolTipText     =   "Clear"
      Top             =   0
      Width           =   315
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   7223
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0000
   End
End
Attribute VB_Name = "frmMain"
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


Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Command1_Click
    End If
    Dim CB As Long
    Dim FindString As String

    If KeyAscii < 32 Or KeyAscii > 127 Then Exit Sub

    If Combo1.SelLength = 0 Then
        FindString = Combo1.Text & Chr$(KeyAscii)
    Else
        FindString = Left$(Combo1.Text, Combo1.SelStart) & Chr$(KeyAscii)
    End If

    CB = SendMessage(Combo1.hwnd, CB_FINDSTRING, -1, ByVal FindString)

    If CB <> CB_ERR Then
        Combo1.ListIndex = CB
        Combo1.SelStart = Len(FindString)
        Combo1.SelLength = Len(Combo1.Text) - Combo1.SelStart
        KeyAscii = 0
    End If
    
End Sub


Private Sub Command1_Click()
'    hed$ = Chr(255) + Chr(255) + Chr(255) + Chr(255) + "rcon" + Chr(32)
'    UDP1.SendData hed$ + " " + Text2 + " " + Combo1.Text + Chr(255) + Chr(255) + Chr(255) + Chr(255)
        
    SendPacket "RC", Combo1.Text
        
    Dim CB As Long
    
    CB = SendMessage(Combo1.hwnd, CB_FINDSTRING, -1, ByVal Combo1.Text)

    If CB = CB_ERR Then
    
        e = InStr(1, Combo1.Text, " ")
        If e > 1 Then
            a$ = LCase(Left(Combo1.Text, e - 1))
            If a$ = "message" Or a$ = "say" Or a$ = "talk" Or a$ = "changename" Then
            Else
                Combo1.AddItem Combo1.Text
            End If
        Else
            Combo1.AddItem Combo1.Text
        End If
    End If
    Combo1.Text = ""
End Sub

Private Sub Command8_Click()
Combo1.Clear

End Sub

Private Sub Form_Load()


On Error Resume Next

nm$ = "FormMainChat"
winash = Val(GetSetting("Server Assistant Client", "Window", nm$ + "winash", -1))

If winash <> -1 Then Me.Show

winmd = Val(GetSetting("Server Assistant Client", "Window", nm$ + "winmd", 3))
If winmd <> 3 Then Me.WindowState = winmd

If Me.WindowState = 0 Then
    winh = GetSetting("Server Assistant Client", "Window", nm$ + "winh", -1)
    wint = GetSetting("Server Assistant Client", "Window", nm$ + "wint", -1)
    winl = GetSetting("Server Assistant Client", "Window", nm$ + "winl", -1)
    winw = GetSetting("Server Assistant Client", "Window", nm$ + "winw", -1) '

    If winh <> -1 Then Me.Height = winh
    If wint <> -1 Then Me.Top = wint
    If winl <> -1 Then Me.Left = winl
    If winw <> -1 Then Me.Width = winw
End If


Dim Combo22() As String
frmMain.Combo1.Clear

datf$ = App.Path + "\lastcmd.dat"


If CheckForFile(datf$) Then
    Open datf$ For Binary As #1

        Get #1, , Cnt

        ReDim Combo22(0 To Cnt)
        Get #1, , Combo22

    Close #1
End If

For i = 0 To UBound(Combo22)
    If Combo22(i) <> "" Then Combo1.AddItem Combo22(i)
Next i


'AddForm True, 375, 277, 0, 0, Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

datf$ = App.Path + "\lastcmd.dat"


If CheckForFile(datf$) Then Kill datf$

Dim Combo() As String
If frmMain.Combo1.ListCount > 0 Then
    ReDim Combo(0 To frmMain.Combo1.ListCount - 1)
    For i = 0 To frmMain.Combo1.ListCount - 1
        Combo(i) = frmMain.Combo1.List(i)
    Next i
End If

Open datf$ For Binary As #1
    Put #1, , (i - 1)
    Put #1, , Combo
Close #1

nm$ = "FormMainChat"
SaveSetting "Server Assistant Client", "Window", nm$ + "winmd", frmMain.WindowState
SaveSetting "Server Assistant Client", "Window", nm$ + "winh", frmMain.Height
SaveSetting "Server Assistant Client", "Window", nm$ + "wint", frmMain.Top
SaveSetting "Server Assistant Client", "Window", nm$ + "winl", frmMain.Left
SaveSetting "Server Assistant Client", "Window", nm$ + "winw", frmMain.Width

If ImGone Then Exit Sub
UnloadTime = True


End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = 1 Then Exit Sub

If Me.WindowState = 0 Then
    'If Me.Width < 5625 Then Me.Width = 5625
    'If Me.Height < 4155 Then Me.Height = 4155
End If
RichTextBox1.Width = Me.ScaleWidth - 4

Combo1.Width = Me.ScaleWidth - Combo1.Left - Command1.Width - Command8.Width - 16
Command8.Left = Combo1.Width + 4 + Combo1.Left + 4 + Command1.Width
Command1.Left = Combo1.Width + 4 + Combo1.Left

RichTextBox1.Height = Me.ScaleHeight - RichTextBox1.Top - 4 - Text1.Height - 8
Text1.Top = RichTextBox1.Height + RichTextBox1.Top + 4
Text1.Width = RichTextBox1.Width


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

    If Text1 <> "" Then
    
        SendPacket "SY", Text1
    End If
    Text1 = ""

    KeyAscii = 0

End If

End Sub
