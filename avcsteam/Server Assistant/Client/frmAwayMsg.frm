VERSION 5.00
Begin VB.Form frmAwayMsg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Away Message"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4680
   Icon            =   "frmAwayMsg.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   3300
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   2355
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   4635
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4515
   End
End
Attribute VB_Name = "frmAwayMsg"
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

