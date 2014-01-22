VERSION 5.00
Begin VB.Form frmSetAway 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Away Mode"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4680
   Icon            =   "frmSetAway.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   4680
   Begin VB.Frame Frame1 
      Caption         =   "Away Mode Options"
      Height          =   1695
      Left            =   60
      TabIndex        =   5
      Top             =   2760
      Width           =   4575
      Begin VB.CheckBox Check3 
         Caption         =   "Automatically return from automatically set Away Mode when returning to computer"
         Height          =   375
         Left            =   60
         TabIndex        =   8
         Top             =   1200
         Width           =   4275
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Automatically set N/A Mode after 20 minutes of inactivity"
         Height          =   375
         Left            =   60
         TabIndex        =   7
         Top             =   720
         Width           =   4275
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Automatically set Away Mode when I leave the computer unattended for more than 10 minutes"
         Height          =   375
         Left            =   60
         TabIndex        =   6
         Top             =   240
         Width           =   4275
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1500
      TabIndex        =   4
      Top             =   4500
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   4500
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   420
      Width           =   4575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Away mode:"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "frmSetAway"
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

Dim lod As Boolean


Private Sub Check1_Click()

updateReg

End Sub

Sub updateReg()
If lod Then Exit Sub

If Check1 = 1 Then SetAway10Min = True
If Check1 = 0 Then SetAway10Min = False
If Check2 = 1 Then SetNA20Min = True
If Check2 = 0 Then SetNA20Min = False
If Check3 = 1 Then AutoAwayReturn = True
If Check3 = 0 Then AutoAwayReturn = False



SaveSetting "Server Assistant", "Settings", "AwayMode10Min", Check1
SaveSetting "Server Assistant", "Settings", "AwayMode20Min", Check2
SaveSetting "Server Assistant", "Settings", "AutoReturn", Check3

End Sub

Private Sub Check2_Click()
updateReg

End Sub

Private Sub Check3_Click()
updateReg

End Sub

Private Sub Combo1_Change()

Index = Combo1.ListIndex
If Index = 1 Then Text1 = "I am away from my computer!"
If Index = 2 Then Text1 = "I am away from my computer, and won't be back until much later!"
If Index = 3 Then Text1 = "I am sleeping!"
If Index = 4 Then Text1 = "I am playing Half-Life!"
If Index = 5 Then Text1 = "I am stuffing my face with food!"


End Sub

Private Sub Command1_Click()

AutoSet = False
MyAwayMode = Combo1.ListIndex
MyAwayMsg = Text1
UpdateAwayMode
Unload Me



End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()

lod = True
If SetAway10Min Then Check1 = 1
If SetNA20Min Then Check2 = 1
If AutoAwayReturn Then Check3 = 1

DoEvents


lod = False


Combo1.AddItem GetAwayName(0)
Combo1.AddItem GetAwayName(1)
Combo1.AddItem GetAwayName(2)
Combo1.AddItem GetAwayName(3)
Combo1.AddItem GetAwayName(4)
Combo1.AddItem GetAwayName(5)

Combo1.ListIndex = MyAwayMode
Text1 = MyAwayMsg

End Sub

