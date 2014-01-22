VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Chat Window"
   ClientHeight    =   3780
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9930
   ClipControls    =   0   'False
   Icon            =   "chatonly.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3780
   ScaleWidth      =   9930
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   8640
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   9480
      Top             =   1680
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   60
      TabIndex        =   13
      Text            =   "Chat Here!"
      ToolTipText     =   "Chat Box"
      Top             =   3420
      Width           =   7035
   End
   Begin VB.Frame Frame3 
      Caption         =   "Server"
      Height          =   615
      Left            =   8640
      TabIndex        =   12
      Top             =   1020
      Width           =   1275
      Begin VB.CommandButton Command9 
         Caption         =   "Players"
         Height          =   315
         Left            =   60
         TabIndex        =   14
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3315
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   60
      Width           =   7035
   End
   Begin VB.Frame Frame2 
      Caption         =   "Log Detail"
      Height          =   2415
      Left            =   7140
      TabIndex        =   1
      Top             =   0
      Width           =   1455
      Begin VB.CheckBox LogDetail 
         Caption         =   "Goals"
         Height          =   255
         Index           =   4
         Left            =   60
         TabIndex        =   16
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox LogDetail 
         Caption         =   "Speech"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox LogDetail 
         Caption         =   "Team Speech"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   7
         Top             =   480
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox LogDetail 
         Caption         =   "Kills"
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox LogDetail 
         Caption         =   "Admin Speech"
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   5
         Top             =   720
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox LogDetail 
         Caption         =   "Misc"
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   4
         ToolTipText     =   "Name changes, sentry, etc"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Select All"
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   1740
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Select None"
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   2040
         Width           =   1335
      End
   End
   Begin MSWinsockLib.Winsock TCP1 
      Left            =   9060
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection"
      Height          =   975
      Left            =   8640
      TabIndex        =   0
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command5 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   11
         Top             =   600
         Width           =   1155
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Connect"
         Height          =   315
         Left            =   60
         TabIndex        =   10
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Label lblUpdate 
      BorderStyle     =   1  'Fixed Single
      Height          =   1275
      Left            =   7140
      TabIndex        =   15
      Top             =   2460
      Width           =   2775
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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



Private Sub Command10_Click()
frmMap.Show



End Sub

Private Sub Command11_Click()

'array height
Dim h As Integer

'array width
Dim w As Integer

'for..next numbers
Dim x1 As Integer, y1 As Integer
Dim n1 As Integer, n2 As Integer, n3 As Integer, n4 As Integer


w = 3
h = 3

'the array
Dim Arr() As Integer
ReDim Arr(1 To w, 1 To h)

'Fill the array with sequencial numbers
a = 1
For Y = 1 To h
    For X = 1 To w
        Arr(X, Y) = a
        a = a + 1
    Next X
Next Y

'better :)

'ill add a TEST printout of the array, so you can see whats in it, kay?

For Y = 1 To h
    For X = 1 To w
        b$ = b$ + Trim(Str(Arr(X, Y))) + "  "
    Next X
    b$ = b$ + vbCrLf
Next Y

MessBox b$

'the string
Dim String1 As String

' start points:
'
'  -----2-----
'  |         |
'  1  ARRAY  3
'  |         |
'  |         |
'  -----4-----
'

'start positions for first loop:
n1 = 1
n2 = 1
n3 = w
n4 = h

'first for thingy dingy


Do

    'going across the top...
    For x1 = n1 To n3
        String1 = String1 + Trim(Str(Arr(x1, n2))) + ", "
    Next x1

    'shrink the borders
    n2 = n2 + 1

    'going down the right side...
    For y1 = n2 To n4
        String1 = String1 + Trim(Str(Arr(n3, y1))) + ", "
    Next y1

    'shrink the borders
    n3 = n3 - 1

    'going across the bottom...
    If n3 > n1 Then
        For x1 = n3 To n1 Step -1
            String1 = String1 + Trim(Str(Arr(x1, n4))) + ", "
        Next x1
    End If
    'shrink the borders
    n4 = n4 - 1

    'going up the left side...
    'If n4 > n2 Then
        For y1 = n4 To n2 Step -1
            String1 = String1 + Trim(Str(Arr(n1, y1))) + ", "
        Next y1
    'End If
    'shrink the borders
    n1 = n1 + 1

Loop Until n1 > n3 Or n2 > n4

MessBox String1

'ready to test it?
'Here goes nothing!!



End Sub

Private Sub Command2_Click()
For I = 0 To LogDetail.Count - 1
    LogDetail(I).Value = 1
Next I

UpdateLogDetail


End Sub

Private Sub Command3_Click()
For I = 0 To LogDetail.Count - 1
    LogDetail(I).Value = 0
Next I

UpdateLogDetail
End Sub

Private Sub Command4_Click()
frmConnect.Show
SendingFile = False
FileRecieveMode = False

End Sub

Private Sub Command5_Click()
If TCP1.State <> sckClosed Then TCP1.Close
Command5.Enabled = False
Command4.Enabled = True


End Sub

Private Sub Command6_Click()
MDIForm1.mnuAdminIn_Click 11


End Sub

Private Sub Command7_Click()
SendPacket "BS", ""

End Sub

Private Sub Command8_Click()
Combo1.Clear

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

Me.Width = 10065
Command4_Click

frmConnect.Move Int(MDIForm1.Width / 2) - Int(frmConnect.Width / 2), Int(MDIForm1.Height / 2) - Int(frmConnect.Height / 2)
frmConnect.Show

End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = 1 Then Exit Sub

If Me.WindowState = 0 Then
    If Me.Width < 5625 Then Me.Width = 5625
    If Me.Height < 4155 Then Me.Height = 4155
End If


Text1.Width = Me.Width - Text1.Left - Frame1.Width - Frame2.Width - 260

'RichTextBox1.Width = Me.Width - Text1.Left - Frame1.Width - Frame2.Width - 260
Combo1.Width = Me.Width - Combo1.Left - Frame1.Width - Frame2.Width - 260 - Command8.Width - 60
Command8.Left = Combo1.Width + 60 + Combo1.Left


Text1.Height = Me.Height - Text1.Top - 500 - Text2.Height - 60

'RichTextBox1.Height = Me.Height - Text1.Top - 500 - Text2.Height - 60

Text2.Top = Text1.Top + Text1.Height + 60

Text2.Width = Text1.Width

Frame2.Left = Text1.Left + Text1.Width + 60
Frame1.Left = Frame2.Left + Frame2.Width + 60
Frame3.Left = Frame1.Left
lblUpdate.Left = Frame2.Left
lblUpdate.Top = Frame2.Top + Frame2.Height + 45


'Text1.Width = Int(Text1.Width / 2) - 60

'RichTextBox1.Width = Text1.Width
'RichTextBox1.Left = Text1.Left


End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveCommands
End


End Sub

Private Sub lblUpdate_Click()

End Sub

Private Sub LogDetail_Click(Index As Integer)

UpdateLogDetail
End Sub



Private Sub TCP1_Close()
AddEvent "**** Disconnected."
Command5_Click


MDIForm1.mnuScripts.Enabled = False
MDIForm1.mnuAdmin.Enabled = False
MDIForm1.mnuSettings.Enabled = False
MDIForm1.mnuWindows.Enabled = False
MDIForm1.mnuMessages.Enabled = False
MDIForm1.mnuLogs.Enabled = False

TCP1.Close
End Sub

Private Sub TCP1_Connect()

'we are connected, send the HL message
SendPacket "HL", ""
UpdateLogDetail
Form1.Text1 = ""
'Form1.'RichTextBox1 = ""
AddEvent "**** Connected..."


End Sub

Private Sub TCP1_DataArrival(ByVal bytesTotal As Long)
'(254)(254)(254)(255)[CODE](255)[PARAMS](255)(253)(253)(253)

MDIForm1.mnuScripts.Enabled = True
MDIForm1.mnuAdmin.Enabled = True
MDIForm1.mnuSettings.Enabled = True
MDIForm1.mnuWindows.Enabled = True
MDIForm1.mnuMessages.Enabled = True
MDIForm1.mnuLogs.Enabled = True

'lblUpdate = bytesTotal
'lblUpdate.Refresh

TCP1.GetData a$

startstr$ = Chr(254) + Chr(254) + Chr(254)
endstr$ = Chr(253) + Chr(253) + Chr(253)

RecData = RecData + a$

'frmFileBrowser.Label1.Caption = "Last packet was " + Ts(Len(a$)) + " bytes" + vbCrLf + "Total size: " + Ts(Len(RecData))


If SendSize > 0 Then
    MDIForm1.StatusBar1.Panels(3).Text = "Download Progress: " + Ts(Int((Len(RecData) / SendSize) * 100)) + "%"
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
            
            Interprit a$
                    
        End If
    End If
    
Loop Until e = 0 Or ee = 0



End Sub

Private Sub TCP1_SendComplete()

MDIForm1.StatusBar1.Panels(2).Text = "Upload Progress: 100%"

End Sub

Private Sub TCP1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)

If bytesRemaining > 0 Then

    MDIForm1.StatusBar1.Panels(2).Text = "Upload Progress: " + Ts(Int((bytesSent / (bytesSent + bytesRemaining)) * 100)) + "%"
End If


End Sub

Private Sub Text2_GotFocus()
If Text2 = "Chat Here!" Then Text2 = ""

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    SendPacket "SY", Text2
    Text2 = ""
    KeyAscii = 0
    
End If

End Sub

Private Sub Timer1_Timer()

If TCP1.State <> sckClosed And Command5.Enabled = False Then
    Command5.Enabled = True
    Command4.Enabled = False
End If

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
    If EmailCheckCounter = 0 Then
        SendPacket "M.", ""
        EmailCheckCounter = 60
    End If
    
End If
End Sub
