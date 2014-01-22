VERSION 5.00
Begin VB.Form frmClans 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clans"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2130
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   2130
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   2880
      Width           =   1035
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2880
      Width           =   1035
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
      Height          =   255
      Left            =   1500
      TabIndex        =   3
      Top             =   2580
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      Height          =   255
      Left            =   660
      TabIndex        =   2
      Top             =   2580
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2580
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   2115
   End
End
Attribute VB_Name = "frmClans"
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

For i = 1 To NumClans
    List1.AddItem Clans(i).Clan
    e = List1.NewIndex
    List1.ItemData(e) = i
Next i

End Sub

Private Sub Command1_Click()

a$ = InBox("Please enter the tag found in all clan members names:", "Add Clan", "")
a$ = Trim(a$)
If a$ = "" Then Exit Sub

b$ = InBox("Please enter the join password for this clan:", "Add Clan", "")
b$ = Trim(b$)
If b$ = "" Then Exit Sub

NumClans = NumClans + 1
n = NumClans
Clans(n).Clan = a$
Clans(n).JoinPass = b$
Clans(n).Flags = 0


UpdateList

End Sub

Private Sub Command2_Click()

a = List1.ListIndex
If a = -1 Then Exit Sub
e = List1.ItemData(a)

NumClans = NumClans - 1

For i = e To NumClans

    Clans(i).Clan = Clans(i + 1).Clan
    Clans(i).JoinPass = Clans(i + 1).JoinPass
    Clans(i).NumMembers = Clans(i + 1).NumMembers
    Clans(i).Flags = Clans(i + 1).Flags
    For j = 1 To Clans(i + 1).NumMembers
    
        Clans(i).Members(j).LastIP = Clans(i + 1).Members(j).LastIP
        Clans(i).Members(j).Name = Clans(i + 1).Members(j).Name
        Clans(i).Members(j).UIN = Clans(i + 1).Members(j).UIN
    Next j
Next i

UpdateList


End Sub





Private Sub Command3_Click()

a = List1.ListIndex

If a = -1 Then Exit Sub
ChosenClan = List1.ItemData(a)

frmMembers.Show

End Sub

Private Sub Command4_Click()

PackageClans
Unload Me


End Sub

Private Sub Command5_Click()
Unload Me

End Sub

Private Sub Form_Load()
UpdateList

End Sub
