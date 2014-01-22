VERSION 5.00
Begin VB.Form frmBans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Bans"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   Icon            =   "frmBans.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   8205
   Begin VB.CommandButton Command5 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   555
      Left            =   4740
      TabIndex        =   6
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   555
      Left            =   5880
      TabIndex        =   5
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   3375
      Left            =   4680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   60
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
      Height          =   555
      Left            =   2640
      TabIndex        =   3
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   555
      Left            =   60
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove"
      Height          =   555
      Left            =   1320
      TabIndex        =   1
      Top             =   3480
      Width           =   1275
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4575
   End
End
Attribute VB_Name = "frmBans"
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

Public Sub ReList()

List1.Clear

For i = 1 To NumCurrBans
    List1.AddItem ServerBans(i).Name + " - " + ServerBans(i).UIDs
    List1.ItemData(List1.NewIndex) = i
Next i

End Sub

Private Sub Command1_Click()


a = List1.ListIndex
If a = -1 Then Exit Sub

e = List1.ItemData(a)

NumCurrBans = NumCurrBans - 1

For i = e To NumCurrBans

    ServerBans(i).BannedAt = ServerBans(i + 1).BannedAt
    ServerBans(i).BanTime = ServerBans(i + 1).BanTime
    ServerBans(i).EntryName = ServerBans(i + 1).EntryName
    ServerBans(i).IP = ServerBans(i + 1).IP
    ServerBans(i).Map = ServerBans(i + 1).Map
    ServerBans(i).Name = ServerBans(i + 1).Name
    ServerBans(i).RealName = ServerBans(i + 1).RealName
    ServerBans(i).Reason = ServerBans(i + 1).Reason
    ServerBans(i).UIDs = ServerBans(i + 1).UIDs

Next i

ReList

End Sub

Private Sub Command3_Click()


a = List1.ListIndex
If a = -1 Then Exit Sub

e = List1.ItemData(a)


Load frmEditBan

frmEditBan.EditBanNum = e
frmEditBan.ShowInfo

End Sub

Private Sub Command4_Click()

'update banlist

a$ = ""

For i = 1 To NumCurrBans
    a$ = a$ + Chr(251)
    a$ = a$ + ServerBans(i).BannedAt + Chr(250)
    a$ = a$ + ServerBans(i).BanTime + Chr(250)
    a$ = a$ + ServerBans(i).EntryName + Chr(250)
    a$ = a$ + ServerBans(i).IP + Chr(250)
    a$ = a$ + ServerBans(i).Map + Chr(250)
    a$ = a$ + ServerBans(i).Name + Chr(250)
    a$ = a$ + ServerBans(i).RealName + Chr(250)
    a$ = a$ + ServerBans(i).Reason + Chr(250)
    a$ = a$ + ServerBans(i).UIDs + Chr(250)
    a$ = a$ + Chr(251)
Next i

SendPacket "GR", a$

Unload Me

End Sub

Private Sub Command5_Click()
Unload Me

End Sub

Private Sub Form_Load()
ReList

End Sub

Private Sub List1_Click()

a = List1.ListIndex
If a = -1 Then Exit Sub

e = List1.ItemData(a)

Text1 = ""

If ServerBans(e).BannedAt <> "" Then
    Text1 = "Banned At:" + vbCrLf
    Text1 = Text1 + ServerBans(e).BannedAt + vbCrLf + vbCrLf
End If
If ServerBans(e).BanTime <> "" Then
    Text1 = Text1 + "Banned for: (0 is infinite)" + vbCrLf
    Text1 = Text1 + ServerBans(e).BanTime + vbCrLf + vbCrLf
End If
If ServerBans(e).Name <> "" Then

    Text1 = Text1 + "Name:" + vbCrLf
    Text1 = Text1 + ServerBans(e).Name + vbCrLf + vbCrLf
End If
If ServerBans(e).EntryName <> "" Then

    Text1 = Text1 + "Entry Name:" + vbCrLf
    Text1 = Text1 + ServerBans(e).EntryName + vbCrLf + vbCrLf
End If
If ServerBans(e).RealName <> "" Then

    Text1 = Text1 + "Real Name:" + vbCrLf
    Text1 = Text1 + ServerBans(e).RealName + vbCrLf + vbCrLf
End If
If ServerBans(e).Reason <> "" Then
    Text1 = Text1 + "Reason:" + vbCrLf
    Text1 = Text1 + ServerBans(e).Reason + vbCrLf + vbCrLf
End If
If ServerBans(e).IP <> "" Then
    Text1 = Text1 + "IP:" + vbCrLf
    Text1 = Text1 + ServerBans(e).IP + vbCrLf + vbCrLf
End If
If ServerBans(e).Map <> "" Then
    Text1 = Text1 + "Map:" + vbCrLf
    Text1 = Text1 + ServerBans(e).Map + vbCrLf + vbCrLf
End If
If ServerBans(e).UIDs <> "" Then
    Text1 = Text1 + "UniqueID's:" + vbCrLf
    Text1 = Text1 + ServerBans(e).UIDs + vbCrLf + vbCrLf
End If




End Sub
