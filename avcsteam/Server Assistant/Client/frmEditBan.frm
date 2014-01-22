VERSION 5.00
Begin VB.Form frmEditBan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Ban"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   Icon            =   "frmEditBan.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4635
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2580
      TabIndex        =   19
      Top             =   3360
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   3540
      TabIndex        =   18
      Top             =   3360
      Width           =   1035
   End
   Begin VB.TextBox Text9 
      Height          =   315
      Left            =   1380
      TabIndex        =   16
      Top             =   2940
      Width           =   3195
   End
   Begin VB.TextBox Text8 
      Height          =   315
      Left            =   1380
      TabIndex        =   14
      Top             =   2580
      Width           =   3195
   End
   Begin VB.TextBox Text7 
      Height          =   315
      Left            =   1380
      TabIndex        =   12
      Top             =   1860
      Width           =   3195
   End
   Begin VB.TextBox Text6 
      Height          =   315
      Left            =   1380
      TabIndex        =   10
      Top             =   1500
      Width           =   3195
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   1380
      TabIndex        =   8
      Top             =   1140
      Width           =   3195
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   1380
      TabIndex        =   6
      Top             =   780
      Width           =   3195
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1380
      TabIndex        =   4
      Top             =   2220
      Width           =   3195
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1380
      TabIndex        =   2
      Top             =   420
      Width           =   3195
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1380
      TabIndex        =   0
      Top             =   60
      Width           =   3195
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "UniqueID's"
      Height          =   195
      Left            =   60
      TabIndex        =   17
      Top             =   3000
      Width           =   780
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Reason"
      Height          =   195
      Left            =   60
      TabIndex        =   15
      Top             =   2640
      Width           =   555
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Real Name"
      Height          =   195
      Left            =   60
      TabIndex        =   13
      Top             =   1920
      Width           =   795
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   1560
      Width           =   420
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Map"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   1200
      Width           =   315
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "IP"
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   840
      Width           =   150
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Entry Name"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   2280
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Ban Time"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Banned At"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   750
   End
End
Attribute VB_Name = "frmEditBan"
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

Public EditBanNum As Integer

Public Sub ShowInfo()

i = EditBanNum

Text1 = ServerBans(i).BannedAt
Text2 = ServerBans(i).BanTime
Text3 = ServerBans(i).EntryName
Text4 = ServerBans(i).IP
Text5 = ServerBans(i).Map
Text6 = ServerBans(i).Name
Text7 = ServerBans(i).RealName
Text8 = ServerBans(i).Reason
Text9 = ServerBans(i).UIDs

End Sub



Private Sub Command1_Click()

i = EditBanNum
ServerBans(i).BannedAt = Text1
ServerBans(i).BanTime = Text2
ServerBans(i).EntryName = Text3
ServerBans(i).IP = Text4
ServerBans(i).Map = Text5
ServerBans(i).Name = Text6
ServerBans(i).RealName = Text7
ServerBans(i).Reason = Text8
ServerBans(i).UIDs = Text9

frmBans.ReList

Unload Me

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

