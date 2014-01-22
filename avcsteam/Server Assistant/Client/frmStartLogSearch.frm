VERSION 5.00
Begin VB.Form frmStartLogSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log Search"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmStartLogSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   60
      TabIndex        =   18
      Top             =   2940
      Width           =   5535
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Search subfolders"
      Height          =   195
      Left            =   60
      TabIndex        =   17
      Top             =   2460
      Value           =   1  'Checked
      Width           =   1875
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Search for exact phrase"
      Height          =   195
      Left            =   60
      TabIndex        =   16
      Top             =   3540
      Value           =   -1  'True
      Width           =   5415
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Search each word seperatly"
      Height          =   195
      Left            =   60
      TabIndex        =   15
      Top             =   3780
      Width           =   5415
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Must contain all words"
      Enabled         =   0   'False
      Height          =   195
      Left            =   300
      TabIndex        =   14
      Top             =   4020
      Value           =   1  'Checked
      Width           =   3975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Search only SAY and SAY_TEAM lines"
      Height          =   195
      Left            =   60
      TabIndex        =   13
      Top             =   3300
      Width           =   3915
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Search between certain dates:"
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   660
      Width           =   2475
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dated Search"
      Enabled         =   0   'False
      Height          =   1515
      Left            =   60
      TabIndex        =   5
      Top             =   900
      Width           =   5535
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   480
         Width           =   4575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Select"
         Height          =   315
         Left            =   4740
         TabIndex        =   8
         Top             =   480
         Width           =   675
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Select"
         Height          =   315
         Left            =   4740
         TabIndex        =   7
         Top             =   1080
         Width           =   675
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Search logs after and including:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2235
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Search logs before and including:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   2370
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Begin Search"
      Height          =   315
      Left            =   4140
      TabIndex        =   4
      Top             =   5100
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   300
      Width           =   5535
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Specify custom search path:"
      Height          =   195
      Left            =   60
      TabIndex        =   19
      Top             =   2700
      Width           =   2010
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enter a Search Criterion into the above box and click Begin Search."
      Height          =   555
      Left            =   60
      TabIndex        =   3
      Top             =   4500
      Width           =   5535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Progress:"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   4260
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Search Criteria:"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1080
   End
End
Attribute VB_Name = "frmStartLogSearch"
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

Dim FromDay As Date
Dim ToDay As Date

Private Sub Check1_Click()
If Check1 = 1 Then Frame1.Enabled = True
If Check1 = 0 Then Frame1.Enabled = False

End Sub

Private Sub Command1_Click()
Text1 = Trim(Text1)
If Text1 = "" Then MessBox "You must specify a search parameter!", vbCritical, "Error Beginning Search": Exit Sub

Command1.Enabled = False
Text1.Enabled = False
Label3 = "Starting Search..."

Frame1.Enabled = False
Check1.Enabled = False
Check3.Enabled = False
Text4.Enabled = False
Check2.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
Check4.Enabled = False

' LS
PackageSearchStart Text1, Check1, FromDay, ToDay, Check4, Text4, Option2, Check4, Check2

End Sub


Private Sub Command2_Click()
FromDay = CalenBox(FromDay, "Search Logs After:")
Text2 = Format(FromDay, "dddd, mmmm d, yyyy hh:mm:ss AMPM")

End Sub

Private Sub Command3_Click()

ToDay = CalenBox(ToDay, "Search Logs Before:")
Text3 = Format(ToDay, "dddd, mmmm d, yyyy hh:mm:ss AMPM")

End Sub

Private Sub Form_Load()

FromDay = Now - 1
ToDay = Now

Text2 = Format(FromDay, "dddd, mmmm d, yyyy hh:mm:ss AMPM")
Text3 = Format(ToDay, "dddd, mmmm d, yyyy hh:mm:ss AMPM")


End Sub

Private Sub Option1_Click()
Check4.Enabled = True

End Sub

Private Sub Option2_Click()
Check4.Enabled = False

End Sub
