VERSION 5.00
Begin VB.Form frmMembers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Clan Properties"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Clan Settings"
      Height          =   1635
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6735
      Begin VB.CheckBox Flags 
         Caption         =   "Remove clan tag from non-registered members"
         Height          =   195
         Index           =   5
         Left            =   2880
         TabIndex        =   22
         Top             =   1380
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   19
         Top             =   180
         Width           =   1635
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         Top             =   540
         Width           =   1635
      End
      Begin VB.CheckBox Flags 
         Caption         =   "Kick Imitators and force join"
         Height          =   195
         Index           =   0
         Left            =   2880
         TabIndex        =   17
         Top             =   180
         Width           =   2355
      End
      Begin VB.CheckBox Flags 
         Caption         =   "Members cannot be kickvoted"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   16
         Top             =   420
         Width           =   3015
      End
      Begin VB.CheckBox Flags 
         Caption         =   "Members do not get autokicked due to spam"
         Height          =   195
         Index           =   2
         Left            =   2880
         TabIndex        =   15
         Top             =   660
         Width           =   3555
      End
      Begin VB.CheckBox Flags 
         Caption         =   "Members recieve double points on Mapvote"
         Height          =   195
         Index           =   3
         Left            =   2880
         TabIndex        =   14
         Top             =   900
         Width           =   3555
      End
      Begin VB.CheckBox Flags 
         Caption         =   "Members recieve double points on Kickvote"
         Height          =   195
         Index           =   4
         Left            =   2880
         TabIndex        =   13
         Top             =   1140
         Width           =   3555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Clan Tag"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Join Password"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Members"
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   6735
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   2835
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3540
         TabIndex        =   7
         Top             =   180
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   540
         Width           =   3135
      End
      Begin VB.TextBox Text5 
         Height          =   765
         Left            =   2940
         TabIndex        =   5
         Top             =   1200
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   2700
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   2700
         Width           =   1155
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Set Changes"
         Height          =   435
         Left            =   5280
         TabIndex        =   2
         Top             =   2040
         Width           =   1395
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   2940
         TabIndex        =   11
         Top             =   240
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Last IP"
         Height          =   195
         Left            =   2940
         TabIndex        =   10
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Unique IDs"
         Height          =   195
         Left            =   2940
         TabIndex        =   9
         Top             =   960
         Width           =   795
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Close"
      Height          =   315
      Left            =   5640
      TabIndex        =   0
      Top             =   4740
      Width           =   1095
   End
End
Attribute VB_Name = "frmMembers"
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

Dim Chosen As Integer
Dim LoadMode As Boolean


Private Sub UpdateList()

List1.Clear

For i = 1 To Clans(Chosen).NumMembers
    List1.AddItem Clans(Chosen).Members(i).Name
    e = List1.NewIndex
    List1.ItemData(e) = i
Next i

Text3 = ""
Text4 = ""
Text5 = ""


End Sub

Private Sub Command1_Click()

a$ = InBox("Please enter the member name:", "Add Member", "")
a$ = Trim(a$)
If a$ = "" Then Exit Sub

b$ = InBox("Please enter the members uniqueid:", "Add Member", "")
b$ = Trim(b$)
If b$ = "" Then Exit Sub

Clans(Chosen).NumMembers = Clans(Chosen).NumMembers + 1
n = Clans(Chosen).NumMembers
Clans(Chosen).Members(n).Name = a$
Clans(Chosen).Members(n).UIN = b$
Clans(Chosen).Members(n).LastIP = ""
UpdateList

End Sub

Private Sub Command2_Click()

a = List1.ListIndex
If a = -1 Then Exit Sub
e = List1.ItemData(a)

'remove

Clans(Chosen).NumMembers = Clans(Chosen).NumMembers - 1

For i = e To Clans(Chosen).NumMembers
    
    Clans(Chosen).Members(i).LastIP = Clans(Chosen).Members(i + 1).LastIP
    Clans(Chosen).Members(i).Name = Clans(Chosen).Members(i + 1).Name
    Clans(Chosen).Members(i).UIN = Clans(Chosen).Members(i + 1).Name
Next i

UpdateList

End Sub

Private Sub Command3_Click()

a = List1.ListIndex
If a = -1 Then Exit Sub
e = List1.ItemData(a)

Clans(Chosen).Members(e).Name = Text3
Clans(Chosen).Members(e).UIN = Text5


End Sub

Private Sub Command4_Click()
Unload Me

End Sub

Private Sub Flags_Click(Index As Integer)
If LoadMode = True Then Exit Sub

For i = 0 To 5
    If Flags(i) = 1 Then a = a + (2 ^ i)
Next i

Clans(Chosen).Flags = a

End Sub

Private Sub Form_Load()

Flags(5).Enabled = DllEnabled

Chosen = ChosenClan
Text1 = Clans(Chosen).Clan
Text2 = Clans(Chosen).JoinPass


LoadMode = True
For i = 0 To 5
    If CheckBit2(Clans(Chosen).Flags, i) Then Flags(i) = 1
Next i
LoadMode = False

UpdateList

End Sub

Private Sub List1_Click()

'show info

a = List1.ListIndex
If a = -1 Then Exit Sub
e = List1.ItemData(a)

Text3 = Clans(Chosen).Members(e).Name
Text4 = Clans(Chosen).Members(e).LastIP
Text5 = Clans(Chosen).Members(e).UIN

End Sub

Private Sub Text1_Change()
If Trim(Text1) = "" Then Exit Sub
If Trim(Text1) = Clans(Chosen).Clan Then Exit Sub

Clans(Chosen).Clan = Text1
End Sub

Private Sub Text2_Change()
If Trim(Text2) = "" Then Exit Sub
If Trim(Text2) = Clans(Chosen).JoinPass Then Exit Sub

Clans(Chosen).JoinPass = Text2
End Sub
