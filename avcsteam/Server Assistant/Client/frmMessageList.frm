VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMessageList 
   Caption         =   "Message List"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   Icon            =   "frmMessageList.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   7710
   Begin VB.CommandButton Command7 
      Caption         =   "Check Mailbox"
      Height          =   375
      Left            =   4500
      TabIndex        =   8
      Top             =   3060
      Width           =   1215
   End
   Begin VB.CommandButton cmdMsg 
      Caption         =   "Mark as Read"
      Height          =   375
      Index           =   3
      Left            =   2880
      TabIndex        =   7
      Top             =   3060
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Close"
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   6180
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2955
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   5212
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Subject"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "From"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "To"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Read"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.CommandButton cmdMsg 
      Caption         =   "Mark as New"
      Height          =   375
      Index           =   2
      Left            =   1620
      TabIndex        =   4
      Top             =   3060
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Compose New Message"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   3060
      Width           =   1935
   End
   Begin VB.CommandButton cmdMsg 
      Caption         =   "Delete"
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   3060
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3480
      Width           =   7635
   End
   Begin VB.CommandButton cmdMsg 
      Caption         =   "Reply"
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   3060
      Width           =   735
   End
End
Attribute VB_Name = "frmMessageList"
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


Public Sub RefreshMessageList()

k = ListView1.SortKey
ListView1.Sorted = False

With ListView1
    With .ListItems

        For i = 1 To .Count
            If .Item(i).Selected = True Then a$ = .Item(i).Text + " " + .Item(i).SubItems(1) + " " + .Item(i).SubItems(2) + " " + .Item(i).SubItems(3)
        Next i

        .Clear
        
        'start adding
        
        For i = 1 To NumMessages
            .Add i
            With .Item(i)
                .Tag = Ts(i)
                    
                .Text = Messages(i).MsgSubj
                .SubItems(1) = Messages(i).MsgFrom
                .SubItems(2) = Messages(i).MsgFor
                .SubItems(3) = Format(Messages(i).MsgTimeSent, "dd/mm/yyyy hh:mm:ss AMPM")
                
                If CheckBit2(Messages(i).Flags, 1) Then .SubItems(4) = "No"
                If CheckBit2(Messages(i).Flags, 1) = False Then .SubItems(4) = "Yes"
                
            End With
        Next i
    


        ListView1.SortKey = k
        ListView1.Sorted = True
        
        
        
        For i = 1 To .Count
            .Item(i).Selected = False
            If a$ = .Item(i).Text + " " + .Item(i).SubItems(1) + " " + .Item(i).SubItems(2) + " " + .Item(i).SubItems(3) Then .Item(i).Selected = True
        Next i
    
    End With
End With

End Sub

Private Sub DisplayMessage(MsgNum)


Text1 = "From: " + Messages(MsgNum).MsgFrom + vbCrLf
Text1 = Text1 + "To: " + Messages(MsgNum).MsgFor + vbCrLf
Text1 = Text1 + "Date: " + Format(Messages(MsgNum).MsgTimeSent, "dd/mm/yyyy hh:mm:ss AMPM") + vbCrLf
Text1 = Text1 + "Subject: " + Messages(MsgNum).MsgSubj + vbCrLf
Text1 = Text1 + vbCrLf
Text1 = Text1 + vbCrLf
Text1 = Text1 + Messages(MsgNum).MsgText

If CheckBit2(Messages(MsgNum).Flags, 1) Then cmdMsg_Click 3


End Sub

Private Sub cmdMsg_Click(Index As Integer)

again:

For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems.Item(i).Selected = True Then
        j = i
        b = Val(ListView1.ListItems.Item(j).Tag)
        
        If Index = 0 Then
        
            frmNewMessage.ReplyTo = Messages(b).MsgFrom
            frmNewMessage.ReplySubj = "Re: " + ReplaceString(Messages(b).MsgSubj, "Re: ", "")
            frmNewMessage.ReplyText = vbCrLf + vbCrLf + vbCrLf + vbCrLf + Messages(b).MsgFrom + " wrote:" + vbCrLf + "> " + ReplaceString(Messages(b).MsgText, vbCrLf, vbCrLf + "> ")
        
            SendPacket "M5", ""
            
            Exit For
        
        ElseIf Index = 1 Then
            
            'cd = MessBox("Are you sure you want to delete this message?", vbYesNo, "Delete Message")
            SendPacket "M2", Ts(Messages(b).MsgID)
            ListView1.ListItems.Remove j
            GoTo again
            
        ElseIf Index = 2 Then
            
            SendPacket "M4", Ts(Messages(b).MsgID)
            If CheckBit2(Messages(b).Flags, 1) = False Then Messages(b).Flags = Messages(b).Flags + 2
            ys = 1
            
        
        ElseIf Index = 3 Then
            SendPacket "M3", Ts(Messages(b).MsgID)
            If CheckBit2(Messages(b).Flags, 1) Then Messages(b).Flags = Messages(b).Flags - 2
            ys = 1
        End If
    End If
Next i

If ys = 1 Then RefreshMessageList


End Sub

Private Sub cmdReply_Click(Index As Integer)



End Sub

Private Sub Command1_Click()
Open App.Path + "\mailtest.txt" For Append As #1

    For i = 1 To NumMessages
        Print #1, vbCrLf + vbCrLf + vbCrLf + vbCrLf + "----------"
        Print #1, Messages(i).Flags
        Print #1, "For: " + Messages(i).MsgFor
        Print #1, "From: " + Messages(i).MsgFrom
        Print #1, "Subject: " + Messages(i).MsgSubj
        Print #1, "ID: " + Ts(Messages(i).MsgID)
        Print #1, "Date: " + Format(Messages(i).MsgTimeSent, "dd/mm/yyyy hh:mm:ss AMPM")
        Print #1, vbCrLf + Messages(i).MsgText
    Next i
Close #1


End Sub

Private Sub Command3_Click()
SendPacket "M5", ""

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()
Unload Me

End Sub

Private Sub Command7_Click()
SendPacket "M6", ""

End Sub

Private Sub Command7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then MDIForm1.PopupMenu MDIForm1.mnuAdminEmail

End Sub

Public Sub MnuClick(Index)

If Index = 0 Then SendPacket "M7", ""

End Sub

Private Sub Form_Load()

RefreshMessageList
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

ListView1.Width = w - (ListView1.Left * 2) - 120
Text1.Width = ListView1.Width

h2 = Int((h - Command7.Height - Command5.Height - Command5.Height - 240) / 2)

ListView1.Height = h2

cmdMsg(0).Top = ListView1.Top + ListView1.Height + 60
cmdMsg(1).Top = ListView1.Top + ListView1.Height + 60
cmdMsg(2).Top = ListView1.Top + ListView1.Height + 60
cmdMsg(3).Top = ListView1.Top + ListView1.Height + 60
Command7.Top = ListView1.Top + ListView1.Height + 60
Command3.Top = ListView1.Top + ListView1.Height + 60

Text1.Top = Command3.Top + Command3.Height + 60
Text1.Height = h2
Command5.Top = Text1.Top + Text1.Height + 60

Command5.Left = w - Command5.Width - 120
On Error Resume Next
nm$ = Me.Name
SaveSetting "Server Assistant Client", "Window", nm$ + "winmd", Me.WindowState
SaveSetting "Server Assistant Client", "Window", nm$ + "winh", Me.Height
SaveSetting "Server Assistant Client", "Window", nm$ + "wint", Me.Top
SaveSetting "Server Assistant Client", "Window", nm$ + "winl", Me.Left
SaveSetting "Server Assistant Client", "Window", nm$ + "winw", Me.Width

End Sub

Private Sub ListView1_Click()


On Error Resume Next

a$ = ListView1.SelectedItem
If a$ = "" Then Exit Sub

For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems.Item(i).Selected = True Then j = i: Exit For
Next i

b = Val(ListView1.ListItems.Item(j).Tag)

DisplayMessage b

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

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
