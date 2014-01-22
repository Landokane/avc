VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Control"
   ClientHeight    =   5865
   ClientLeft      =   1755
   ClientTop       =   2535
   ClientWidth     =   7845
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "udp.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   391
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   523
   Begin VB.CommandButton Command8 
      Height          =   315
      Left            =   2400
      Picture         =   "udp.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Left            =   2400
      Picture         =   "udp.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   0
      Width           =   375
   End
   Begin MSComDlg.CommonDialog Dlg1 
      Left            =   4500
      Top             =   4260
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RT2 
      Height          =   735
      Left            =   10320
      TabIndex        =   17
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      TextRTF         =   $"udp.frx":0F56
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
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   4980
      Top             =   4260
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5460
      Top             =   4260
   End
   Begin VB.Frame Frame3 
      Caption         =   "Server"
      Height          =   2415
      Left            =   1500
      TabIndex        =   12
      Top             =   600
      Width           =   1275
      Begin VB.Timer Timer3 
         Interval        =   15000
         Left            =   60
         Top             =   1620
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Players"
         Height          =   315
         Left            =   60
         TabIndex        =   15
         Top             =   960
         Width           =   1155
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Scripts"
         Height          =   315
         Left            =   60
         TabIndex        =   14
         Top             =   600
         Width           =   1155
      End
      Begin VB.CommandButton Command6 
         Caption         =   "File Manager"
         Height          =   315
         Left            =   60
         TabIndex        =   13
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Log Detail"
      Height          =   2415
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   1455
      Begin VB.CommandButton Command10 
         Caption         =   "Freeze"
         Height          =   255
         Left            =   60
         TabIndex        =   18
         Top             =   2100
         Width           =   1335
      End
      Begin VB.CheckBox LogDetail 
         Caption         =   "Speech"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox LogDetail 
         Caption         =   "Team Speech"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   8
         Top             =   480
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox LogDetail 
         Caption         =   "Kills"
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox LogDetail 
         Caption         =   "Admin Speech"
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   6
         Top             =   720
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox LogDetail 
         Caption         =   "Goals"
         Height          =   255
         Index           =   4
         Left            =   60
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox LogDetail 
         Caption         =   "Misc"
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   4
         ToolTipText     =   "Name changes, sentry, etc"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "All"
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "None"
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   1800
         Width           =   675
      End
   End
   Begin MSWinsockLib.Winsock TCP1 
      Left            =   5940
      Top             =   4260
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton Command5 
         Caption         =   "Reconnect"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Connect"
         Height          =   315
         Left            =   60
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label lblUpdate 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   0
      TabIndex        =   16
      Top             =   3060
      Width           =   2775
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
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


Public OldWindowProc As Long











Private Sub Command1_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Command10_Click()

If ChatFrozen = True Then
    Command10.Caption = "Freeze"
    ChatFrozen = False
    
    'replace
    
    n = UBound(FreezeText)
    
    For i = 1 To n
        UnPackageMessage FreezeText(i)
    Next i
    ReDim FreezeText(0 To 0)
Else
    
    ChatFrozen = True
    Command10.Caption = "Continue"
    
End If


    
    


End Sub

Private Sub Command11_Click()


'MsgBox SelIcon





End Sub

Private Sub Command12_Click()
SendPacket "MP", ""


End Sub

Private Sub Command13_Click()

a$ = a$ + MDIForm1.Toolbar1.Buttons(Val(Text3)).Description + vbCrLf
a$ = a$ + Ts(MDIForm1.Toolbar1.Buttons(Val(Text3)).Visible) + vbCrLf
a$ = a$ + Ts(MDIForm1.Toolbar1.Buttons(Val(Text3)).Enabled) + vbCrLf
a$ = a$ + Ts(MDIForm1.Toolbar1.Buttons(Val(Text3)).Left) + vbCrLf
a$ = a$ + Ts(MDIForm1.Toolbar1.Buttons(Val(Text3)).Height) + vbCrLf
a$ = a$ + Ts(MDIForm1.Toolbar1.Buttons(Val(Text3)).Style) + vbCrLf
a$ = a$ + Ts(MDIForm1.Toolbar1.Buttons(Val(Text3)).Width) + vbCrLf
a$ = a$ + Ts(MDIForm1.Toolbar1.Buttons(Val(Text3)).Value) + vbCrLf
a$ = a$ + Ts(MDIForm1.Toolbar1.Buttons(Val(Text3)).Height) + vbCrLf
a$ = a$ + Ts(MDIForm1.Toolbar1.Buttons(Val(Text3)).Index) + vbCrLf
a$ = a$ + MDIForm1.Toolbar1.Buttons(Val(Text3)).Key + vbCrLf
a$ = a$ + Ts(MDIForm1.Toolbar1.Buttons(Val(Text3)).MixedState) + vbCrLf
a$ = a$ + Ts(MDIForm1.Toolbar1.Buttons(Val(Text3)).Top) + vbCrLf


MessBox a$

End Sub

Private Sub Command14_Click()

'MDIForm1.Toolbar1.Buttons.Item(1).
'MDIForm1.Toolbar1.Buttons.Remove 1

For i = 1 To MDIForm1.Toolbar1.Controls.Count

    If MDIForm1.Toolbar1.Buttons(i).Style = tbrSeparator Then MsgBox "a"

    


Next i

End Sub

Private Sub Command1aa_Click()
frmMain.Show

End Sub

Private Sub Command2_Click()
For i = 0 To LogDetail.Count - 1
    LogDetail(i).Value = 1
Next i

UpdateLogDetail


End Sub

Private Sub Command3_Click()
For i = 0 To LogDetail.Count - 1
    LogDetail(i).Value = 0
Next i

UpdateLogDetail
End Sub

Private Sub Command4_Click()

If Command4.Caption = "Connect" Then


    frmConnect.Show
    SendingFile = False
    FileRecieveMode = False
Else

    If TCP1.State <> sckClosed Then TCP1.Close
    Command4.Caption = "Connect"
    PrevConnection = False
    'Command5.Enabled = False
    'Command4.Enabled = True

End If

End Sub

Private Sub Command5_Click()

Form1.TCP1.Close
DoEvents
Command4.Caption = "Connect"
Form1.TCP1.Connect

End Sub

Private Sub Command6_Click()
MDIForm1.mnuAdminIn_Click 11


End Sub

Private Sub Command7_Click()
ButtonShowMode = 0
SendPacket "BS", ""

End Sub



Private Sub Command8_Click()
Unload MDIForm1
End Sub

Private Sub Command9_Click()
ShowPlayers = True
SendPacket "SU", ""
Form6.Show

End Sub

Private Sub Form_Load()

MDIForm1.mnuScripts.Enabled = False
MDIForm1.mnuAdmin.Enabled = False
MDIForm1.mnuSettings.Enabled = False
MDIForm1.mnuWindows.Enabled = False
MDIForm1.mnuMessages.Enabled = False
MDIForm1.mnuLogs.Enabled = False

'Command4_Click
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

    
AddForm True, 190, 283, 202, 295, Me

End Sub

Private Sub Form_Paint()
If Me.WindowState = 2 Then Me.WindowState = 0: Exit Sub
Static doneb4

If doneb4 = 0 Then
    frmConnect.Show
    frmConnect.Move Int(MDIForm1.Width / 2) - Int(frmConnect.Width / 2), Int(MDIForm1.Height / 2) - Int(frmConnect.Height / 2)
    frmConnect.SetFocus
End If

doneb4 = 1

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If ImGone Then Exit Sub

'UnloadTime = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

If Not UnloadTime Then Cancel = 1: Exit Sub

UnloadReady

nm$ = Form1.Name
SaveSetting "Server Assistant Client", "Window", nm$ + "winmd", Form1.WindowState
SaveSetting "Server Assistant Client", "Window", nm$ + "winh", Form1.Height
SaveSetting "Server Assistant Client", "Window", nm$ + "wint", Form1.Top
SaveSetting "Server Assistant Client", "Window", nm$ + "winl", Form1.Left
SaveSetting "Server Assistant Client", "Window", nm$ + "winw", Form1.Width



If ImGone Then Exit Sub
UnloadTime = True

End Sub

Private Sub LogDetail_Click(Index As Integer)

UpdateLogDetail
End Sub



Private Sub TCP1_Close()

AddEvent "**** Disconnected."
Command4.Caption = "Disconnect"
Command4_Click


MDIForm1.mnuScripts.Enabled = False
MDIForm1.mnuAdmin.Enabled = False
MDIForm1.mnuSettings.Enabled = False
MDIForm1.mnuWindows.Enabled = False
MDIForm1.mnuMessages.Enabled = False
MDIForm1.mnuLogs.Enabled = False
MDIForm1.Toolbar1.Enabled = False

TCP1.Close
LastKnownState = 0

End Sub

Private Sub TCP1_Connect()

'we are connected, send the HL message
frmMain.RichTextBox1.Text = ""
AddEvent "**** Connected, negotiating with host..."

EncryptedMode = False

SendPacket "X1", ""

'SendPacket "HL", ""
'UpdateLogDetail


End Sub

Private Sub TCP1_DataArrival(ByVal bytesTotal As Long)
'(254)(254)(254)(255)[CODE](255)[PARAMS](255)(253)(253)(253)

MDIForm1.mnuScripts.Enabled = True
MDIForm1.mnuAdmin.Enabled = True
MDIForm1.mnuSettings.Enabled = True
MDIForm1.mnuWindows.Enabled = True
MDIForm1.mnuMessages.Enabled = True
MDIForm1.mnuLogs.Enabled = True
MDIForm1.Toolbar1.Enabled = True

'lblUpdate = bytesTotal
'lblUpdate.Refresh

TCP1.GetData a$
'Debug.Print a$


startstr$ = Chr(254) + Chr(254) + Chr(254)
endstr$ = Chr(253) + Chr(253) + Chr(253)

RecData = RecData + a$

'frmFileBrowser.Label1.Caption = "Last packet was " + Ts(Len(a$)) + " bytes" + vbCrLf + "Total size: " + Ts(Len(RecData))


If SendSize > 0 Then
    MDIForm1.StatusBar1.Panels(4).Text = "Download Progress: " + Ts(Int((Len(RecData) / SendSize) * 100)) + "%"
End If
    

'If Len(RecData) > 200000 Then
'    RecData = ""
'    MessBox "OVERFLOW ERROR!"
'End If

Do
    e = InStr(1, RecData, startstr$)
    ee = InStr(e + 1, RecData, endstr$)
        
    If e And InStr(e + 1, RecData, endstr$) > 0 Then 'there is a whole line
    
        If e > 1 Then 'not at beginning
            RecData = Right(RecData, Len(RecData) - e + 1)
            e = InStr(1, RecData, startstr$)
        End If
    
        'extract
        f = InStr(e + 1, RecData, endstr$)
        
        If e > 0 And f > e And f > 0 Then
            a$ = Mid(RecData, e, f - e + 3)
                    
            If Len(RecData) - Len(a$) > 0 Then
                RecData = Right(RecData, Len(RecData) - Len(a$))
            Else
                RecData = ""
            End If


            If EncryptedMode Then
                    a$ = Right(a$, Len(a$) - 3)
                    a$ = Left(a$, Len(a$) - 3)
                    a$ = Encrypt(a$, LoginPass)
            End If
                
            Interprit a$
                    
        End If
    End If
nxtpacket:
Loop Until e = 0 Or ee = 0



End Sub

Private Sub TCP1_SendComplete()

MDIForm1.StatusBar1.Panels(3).Text = "Upload Progress: 100%"

End Sub

Private Sub TCP1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)

If bytesRemaining > 0 Then

    MDIForm1.StatusBar1.Panels(3).Text = "Upload Progress: " + Ts(Int((bytesSent / (bytesSent + bytesRemaining)) * 100)) + "%"
End If


End Sub



Private Sub Timer1_Timer()


If TCP1.State <> sckClosed And Command4.Caption <> "Disconnect" Then
    Command5.Enabled = True
    'Command4.Enabled = False
    Command4.Caption = "Disconnect"
    
End If

If TCP1.State = 0 And LastKnownState <> 0 Then
    AddEvent "**** Disconnected."
    Command4.Caption = "Disconnect"
    Command4_Click
    
    
    MDIForm1.mnuScripts.Enabled = False
    MDIForm1.mnuAdmin.Enabled = False
    MDIForm1.mnuSettings.Enabled = False
    MDIForm1.mnuWindows.Enabled = False
    MDIForm1.mnuMessages.Enabled = False
    MDIForm1.mnuLogs.Enabled = False
    MDIForm1.Toolbar1.Enabled = False
End If

If TCP1.State = sckError Then TCP1.Close

LastKnownState = TCP1.State

End Sub

Private Sub Timer2_Timer()

If TCP1.State = sckConnected Then
    
    'Decrease Map Time Remaining counter
    If SecondsLeft > 0 Then
        SecondsLeft = SecondsLeft - 1
        UpdateLabel
    End If

    'See if its needed to send back file that was edited
    
    If EditMode = True Then
        wn$ = "temp1.txt - Notepad"
        If CheckWindowThere(wn$) = False Then
            EditMode = False
            'send file back
            
            frmFileBrowser.CmdEditComplete(0).Enabled = True
            frmFileBrowser.CmdEditComplete(1).Enabled = True

            
'            ms = MessBox("When you are done editing the file " + EditFile + "," + vbCrLf + "click YES to update it on the server," + vbCrLf + "or click NO not to update.", vbQuestion + vbYesNo, "File Edit")
'            If ms = vbYes Then PackageFileSend EditFileTemp, EditFile
        End If
    End If
    
    If EmailCheckCounter > 0 Then EmailCheckCounter = EmailCheckCounter - 1
    If EmailCheckCounter = 0 And EncryptedMode = True Then
        SendPacket "M.", ""
        EmailCheckCounter = 60
    End If
    
Else
    If PrevConnection = True Then
    
        Command5_Click
        
    
    End If
End If

' see if we are away

Dim Loc As POINTAPI
nn = GetCursorPos(Loc)

If Loc.X = LastMouseX And Loc.Y = LastMouseY Then
    SecondsAway = SecondsAway + 1
Else
    SecondsAway = 0
    
    If AutoAwayReturn And MyAwayMode > 0 And AutoSet = True Then
        
        MyAwayMode = 0
        MyAwayMsg = ""
        UpdateAwayMode
    End If
    
    AutoSet = False
    
End If

If SecondsAway Mod 60 = 0 And SecondsAway >= 120 Then
    ' Tell the server how long we've been idle.
    SendPacket "ID", Ts(SecondsAway)
End If


If SetAway10Min And SecondsAway = 600 And AwayMode = 0 Then
    MyAwayMode = 1
    MyAwayMsg = "I've been away from the computer for 10 or more minutes!"
    UpdateAwayMode
    AutoSet = True
End If

If SetNA20Min And SecondsAway = 1200 And AwayMode < 2 Then
    MyAwayMode = 2
    MyAwayMsg = "I've been away from the computer for 20 or more minutes!"
    UpdateAwayMode
    AutoSet = True
End If

LastMouseX = Loc.X
LastMouseY = Loc.Y

End Sub

Private Sub Timer3_Timer()
If EncryptedMode = True Then SendPacket "U2", ""
End Sub
