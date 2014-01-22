VERSION 5.00
Begin VB.Form frmWebInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Web Log Info"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   Icon            =   "frmWebInfo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Log Detail"
      Height          =   1275
      Left            =   60
      TabIndex        =   6
      Top             =   720
      Width           =   5535
      Begin VB.CheckBox Chck 
         Caption         =   "Joins / Leaves"
         Height          =   195
         Index           =   6
         Left            =   2880
         TabIndex        =   13
         Top             =   720
         Width           =   1515
      End
      Begin VB.CheckBox Chck 
         Caption         =   "Team Changes"
         Height          =   195
         Index           =   5
         Left            =   2880
         TabIndex        =   12
         Top             =   480
         Width           =   1515
      End
      Begin VB.CheckBox Chck 
         Caption         =   "Class Changes"
         Height          =   195
         Index           =   4
         Left            =   2880
         TabIndex        =   11
         Top             =   240
         Width           =   1515
      End
      Begin VB.CheckBox Chck 
         Caption         =   "Name Changes"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1515
      End
      Begin VB.CheckBox Chck 
         Caption         =   "Goals"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   915
      End
      Begin VB.CheckBox Chck 
         Caption         =   "Kills"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   915
      End
      Begin VB.CheckBox Chck 
         Caption         =   "Speech (global / team / admin)"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2595
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit Web Colors..."
      Height          =   315
      Left            =   3960
      TabIndex        =   5
      Top             =   0
      Width           =   1635
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Enable Web Logging"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Text            =   "c:\sierra\half-life"
      Top             =   360
      Width           =   4275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2820
      TabIndex        =   0
      Top             =   2040
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Web Logs Path"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   420
      Width           =   1215
   End
End
Attribute VB_Name = "frmWebInfo"
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

Web.LogPath = Text1
If Check1 = 1 Then Web.Enabled = True
If Check1 = 0 Then Web.Enabled = False

For i = Chck.Count - 1 To 0 Step -1
    b = 2 ^ i
    If Chck(i) Then a = a + b
Next i

Web.LogFlags = a

PackageWebInfo
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
SendPacket "WL", ""
End Sub

Private Sub Form_Load()
Text1 = Web.LogPath
If Web.Enabled = True Then Check1 = 1
If Web.Enabled = False Then Check1 = 0

a = Web.LogFlags
For i = Chck.Count - 1 To 0 Step -1
    b = 2 ^ i
    If a >= b Then a = a - b: Chck(i) = 1
Next i

End Sub
