VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTransferProgress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Transfer"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   Icon            =   "frmTransferProgress.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   6900
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   4740
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   660
      Width           =   2115
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   660
      Width           =   2115
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4740
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   360
      Width           =   2115
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   360
      Width           =   2115
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   60
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Abort"
      Height          =   315
      Left            =   5640
      TabIndex        =   1
      Top             =   1380
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Time Left"
      Height          =   195
      Left            =   3540
      TabIndex        =   6
      Top             =   720
      Width           =   660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Transfer Rate"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   660
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Bytes Complete"
      Height          =   195
      Left            =   3540
      TabIndex        =   4
      Top             =   420
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Size"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   420
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   285
   End
End
Attribute VB_Name = "frmTransferProgress"
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
SendPacket "F.", ""
FileStop = True

End Sub

