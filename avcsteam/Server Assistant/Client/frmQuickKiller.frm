VERSION 5.00
Begin VB.Form frmQuickKiller 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quick Killer"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   Icon            =   "frmQuickKiller.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   3405
   Begin VB.CommandButton Command2 
      Caption         =   "Done"
      Height          =   315
      Left            =   1740
      TabIndex        =   2
      Top             =   3660
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Kill"
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   3660
      Width           =   1635
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   60
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   3315
   End
End
Attribute VB_Name = "frmQuickKiller"
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

Sub ReList()

List1.Clear

For i = 1 To NumCommands

    List1.AddItem Commands(i).Group + " - " + Commands(i).Name
    a = List1.NewIndex
    List1.ItemData(a) = i

Next i

Me.Caption = "Quick Killer - " + Ts(NumCommands)

End Sub

Private Sub Command1_Click()

again:
For i = 0 To List1.ListCount - 1

    b = List1.ItemData(i)
    c = List1.Selected(i)
    
    If c Then
        
        'kill it
        
        KillScript b
        
        'now, re-number the entire list.
        
        For j = 0 To List1.ListCount - 1
    
            d = List1.ItemData(j)
            If d > b Then
                List1.ItemData(j) = d - 1
            End If
        Next j
        
        List1.RemoveItem i
        GoTo again
    End If

Next i


ReList



End Sub

Private Sub Command2_Click()

Form3.Show
Unload Me


End Sub

Private Sub Form_Load()
ReList
End Sub


Sub KillScript(Num)

For i = Num To NumCommands - 1
    Commands(i).Exec = Commands(i + 1).Exec
    Commands(i).Name = Commands(i + 1).Name
    Commands(i).NumParams = Commands(i + 1).NumParams
    Commands(i).MustHave = Commands(i + 1).MustHave
    Commands(i).ScriptName = Commands(i + 1).ScriptName
    Commands(i).AutoMakeVars = Commands(i + 1).AutoMakeVars
    Commands(i).Group = Commands(i + 1).Group
    Commands(i).LogExec = Commands(i + 1).LogExec
    Commands(i).Unused1 = Commands(i + 1).Unused1
    Commands(i).unused2 = Commands(i + 1).unused2
    Commands(i).unused3 = Commands(i + 1).unused3
    Commands(i).ScriptID = Commands(i + 1).ScriptID
    Commands(i).Unused5 = Commands(i + 1).Unused5
    Commands(i).NumButtons = Commands(i + 1).NumButtons
    Commands(i).Changed = Commands(i + 1).Changed
    ReDim Commands(i).Buttons(0 To Commands(i + 1).NumButtons)
    For j = 1 To Commands(i + 1).NumButtons
        Commands(i).Buttons(j) = Commands(i + 1).Buttons(j)
    Next j
Next i

NumCommands = NumCommands - 1
ReDim Preserve Commands(0 To NumCommands)


End Sub
