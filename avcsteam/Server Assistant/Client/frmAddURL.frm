VERSION 5.00
Begin VB.Form frmAddURL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add a URL"
   ClientHeight    =   1260
   ClientLeft      =   10830
   ClientTop       =   8925
   ClientWidth     =   5160
   Icon            =   "frmAddURL.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   5160
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   840
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   600
      TabIndex        =   2
      Text            =   "Google"
      Top             =   420
      Width           =   4515
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Text            =   "www.google.com"
      Top             =   60
      Width           =   4515
   End
   Begin VB.Label Label2 
      Caption         =   "Text"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "URL:"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "frmAddURL"
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

frmAdminChat.Text2 = frmAdminChat.Text2 & "[url=" & Text1 & "]" & Text2 & "[/url]"
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()


Me.Left = frmAdminChat.Left + 500
Me.Top = frmAdminChat.Top + 500

End Sub
