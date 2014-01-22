VERSION 5.00
Begin VB.Form frmKickBan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kick-Ban List"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "udp9.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6630
   Begin VB.CommandButton Command3 
      Caption         =   "Set"
      Height          =   315
      Left            =   3480
      TabIndex        =   16
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Ban kicked player immediately"
      Height          =   195
      Left            =   3480
      TabIndex        =   15
      Top             =   2100
      Width           =   2475
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Kick by UID"
      Height          =   195
      Left            =   3480
      TabIndex        =   14
      Top             =   1860
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Kick by Clan"
      Height          =   195
      Left            =   3480
      TabIndex        =   13
      Top             =   1620
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Kick by Name"
      Height          =   195
      Left            =   3480
      TabIndex        =   12
      Top             =   1380
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   4200
      TabIndex        =   11
      Top             =   720
      Width           =   2355
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   4200
      TabIndex        =   10
      Top             =   360
      Width           =   2355
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   4200
      TabIndex        =   9
      Top             =   0
      Width           =   2355
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3060
      Width           =   675
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Del"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   3060
      Width           =   675
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   3060
      Width           =   795
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   255
      Left            =   2640
      TabIndex        =   0
      Top             =   3060
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Flags"
      Height          =   195
      Left            =   3480
      TabIndex        =   8
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "UniqueID"
      Height          =   195
      Left            =   3480
      TabIndex        =   7
      Top             =   780
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Clan"
      Height          =   195
      Left            =   3480
      TabIndex        =   6
      Top             =   420
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   3480
      TabIndex        =   5
      Top             =   60
      Width           =   420
   End
End
Attribute VB_Name = "frmKickBan"
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

Private Sub UpdateList()
List1.Clear

For i = 1 To NumKickBans

    List1.AddItem KickBans(i).Name
    List1.ItemData(List1.NewIndex) = i

Next i


End Sub

Private Sub Command1_Click()

'add

n$ = InBox("Enter Name:", "New Kick-Ban", "")

n$ = Trim(n$)

If n$ = "" Then Exit Sub

For i = 1 To NumKickBans
    If n$ = KickBans(i).Name Then Exit Sub
Next i

NumKickBans = NumKickBans + 1
b = NumKickBans

KickBans(b).Name = n$
KickBans(b).Type = 1

UpdateList

End Sub

Private Sub Command2_Click()

a = List1.ListIndex
If a = -1 Then Exit Sub

'delete

b = MessBox("Are you sure you want to remove this kick-ban?", vbYesNo + vbQuestion, "Delete kick-ban")

If b = vbYes Then

    c = List1.ItemData(a)
    NumKickBans = NumKickBans - 1
    
    For i = c To NumKickBans

        KickBans(i).Clan = KickBans(i + 1).Clan
        KickBans(i).Name = KickBans(i + 1).Name
        KickBans(i).Type = KickBans(i + 1).Type
        KickBans(i).UID = KickBans(i + 1).UID
        
    Next i

    UpdateList
    
End If



End Sub

Private Sub Command3_Click()

a = List1.ListIndex
If a = -1 Then Exit Sub

e = List1.ItemData(a)

KickBans(e).Clan = Text2
KickBans(e).Name = Text1
KickBans(e).UID = Text3

If Check1 = 1 Then b = b + 1
If Check2 = 1 Then b = b + 2
If Check3 = 1 Then b = b + 4
If Check4 = 1 Then b = b + 8

KickBans(e).Type = b


End Sub

Private Sub Command4_Click()
Unload Me

End Sub

Private Sub Command5_Click()

PackageKickBans

Unload Me

End Sub

Private Sub Form_Load()
UpdateList

End Sub

Sub ShowPlayer()

a = List1.ListIndex
If a = -1 Then Exit Sub

e = List1.ItemData(a)

Text1 = KickBans(e).Name
Text2 = KickBans(e).Clan
Text3 = KickBans(e).UID

Check1 = 0
Check2 = 0
Check3 = 0
Check4 = 0

If CheckBit2(KickBans(e).Type, 0) Then Check1 = 1
If CheckBit2(KickBans(e).Type, 1) Then Check2 = 1
If CheckBit2(KickBans(e).Type, 2) Then Check3 = 1
If CheckBit2(KickBans(e).Type, 3) Then Check4 = 1

End Sub

Private Sub List1_Click()
ShowPlayer

End Sub
