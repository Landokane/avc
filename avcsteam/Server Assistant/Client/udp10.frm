VERSION 5.00
Begin VB.Form frmUserList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Users List"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   Icon            =   "udp10.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   255
      Left            =   3660
      TabIndex        =   6
      Top             =   3300
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2820
      TabIndex        =   5
      Top             =   3300
      Width           =   795
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   3300
      Width           =   675
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Del"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   3300
      Width           =   675
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3300
      Width           =   675
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4395
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmUserList"
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

Public Sub UpdateList()
List1.Clear

For i = 1 To NumUsers

    List1.AddItem Users(i).Name
    List1.ItemData(List1.NewIndex) = i

Next i


End Sub

Private Sub Command1_Click()

'add

n$ = InBox("Enter Name:", "New User", "Joe")
p$ = InBox("Enter Password:", "New User", "Joe")

n$ = Trim(n$)
p$ = Trim(p$)

If n$ = "" Then Exit Sub
If p$ = "" Then Exit Sub

For i = 1 To NumUsers
    If n$ = Users(i).Name Then Exit Sub
Next i

NumUsers = NumUsers + 1
b = NumUsers

Users(b).Name = n$
Users(b).PassWord = p$
Users(b).Allowed = "users"
Users(b).Flags = 2


UpdateList

End Sub

Private Sub Command2_Click()

a = List1.ListIndex
If a = -1 Then Exit Sub

'delete

b = MessBox("Are you sure you want to remove this user?", vbYesNo + vbQuestion, "Delete User")

If b = vbYes Then

    c = List1.ItemData(a)
    NumUsers = NumUsers - 1
    
    For i = c To NumUsers

        Users(i).Name = Users(i + 1).Name
        Users(i).PassWord = Users(i + 1).PassWord
        Users(i).Allowed = Users(i + 1).Allowed
        Users(i).Flags = Users(i + 1).Flags
        
    Next i

    UpdateList
    
End If



End Sub

Private Sub Command3_Click()
a = List1.ListIndex
If a = -1 Then Exit Sub

b = List1.ItemData(a)

UserEditNum = b

Form11.Show


End Sub

Private Sub Command4_Click()
Unload Me

End Sub

Private Sub Command5_Click()

SendUserEdit
Unload Me


End Sub

Private Sub Form_Load()
UpdateList

End Sub

