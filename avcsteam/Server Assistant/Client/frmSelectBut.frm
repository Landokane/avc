VERSION 5.00
Begin VB.Form frmSelectBut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select A Script"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3360
   Icon            =   "frmSelectBut.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   3360
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1740
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Run 
      Caption         =   "Run"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   2760
      Width           =   1635
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   300
      Width           =   3255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select a script to start."
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1575
   End
End
Attribute VB_Name = "frmSelectBut"
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

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()

'display

For i = 1 To NumCommands
    If Commands(i).ScriptName <> "" Then
        doit = 1
        
        If Commands(i).NumButtons > 0 Then
            If Commands(i).Buttons(1).Type = 3 Then doit = 0
        End If
        
        If doit = 1 Then
            List1.AddItem Commands(i).ScriptName
            e = List1.NewIndex
            List1.ItemData(e) = i
        End If
    End If
Next i



End Sub

Private Sub Run_Click()

a = List1.ListIndex
If a = -1 Then Exit Sub
e = List1.ItemData(a)


frmControlFill.ButIndex = e
frmControlFill.Draw
'frmControlFill.Show

Unload Me

End Sub
