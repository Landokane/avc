VERSION 5.00
Begin VB.Form frmServerInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Setup"
   ClientHeight    =   6210
   ClientLeft      =   5115
   ClientTop       =   4530
   ClientWidth     =   5670
   Icon            =   "server1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   5670
   Begin VB.Frame Frame1 
      Caption         =   "HLDS"
      Height          =   3255
      Left            =   60
      TabIndex        =   15
      Top             =   2580
      Width           =   5535
      Begin VB.CheckBox Check2 
         Caption         =   "Auto-restart server when it goes down"
         Height          =   255
         Left            =   60
         TabIndex        =   26
         Top             =   2940
         Width           =   4035
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   60
         TabIndex        =   25
         Text            =   "c:\hlserver"
         Top             =   2640
         Width           =   5355
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Text            =   "+maxplayers 12 +map well -game tfc"
         Top             =   1080
         Width           =   5355
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Text            =   "c:\hlserver\hlds.exe"
         Top             =   480
         Width           =   5355
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Folder that HLDS is in:"
         Height          =   195
         Left            =   60
         TabIndex        =   24
         Top             =   2400
         Width           =   1590
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "-game tfc"
         Height          =   195
         Left            =   300
         TabIndex        =   23
         Top             =   2160
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "+maxplayers 18"
         Height          =   195
         Left            =   300
         TabIndex        =   22
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "+map WELL"
         Height          =   195
         Left            =   300
         TabIndex        =   21
         Top             =   1680
         Width           =   885
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Examples:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Command line arguements:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   1905
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Complete path to HLDS (ex: c:\hlserver\hlds.exe)"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Use Server Assistant to control HLDS"
      Height          =   255
      Left            =   60
      TabIndex        =   14
      Top             =   2280
      Width           =   4035
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1920
      TabIndex        =   12
      Text            =   "209.166.22.5"
      Top             =   1860
      Width           =   1995
   End
   Begin VB.TextBox Text6 
      Height          =   315
      Left            =   1920
      TabIndex        =   11
      Text            =   "tfc"
      Top             =   420
      Width           =   3675
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   1920
      TabIndex        =   9
      Text            =   "25500"
      Top             =   1500
      Width           =   1995
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   315
      Left            =   2040
      TabIndex        =   7
      Top             =   5880
      Width           =   3615
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Text            =   "26000"
      Top             =   1140
      Width           =   1995
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Text            =   "27015"
      Top             =   780
      Width           =   1995
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Text            =   "c:\sierra\half-life"
      Top             =   60
      Width           =   3675
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Server IP"
      Height          =   195
      Left            =   60
      TabIndex        =   13
      Top             =   1920
      Width           =   660
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Rcon Listen Port"
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   1560
      Width           =   1185
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Remote Connect Port"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   1200
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Game Server Port"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   840
      Width           =   1260
   End
   Begin VB.Label Label2 
      Caption         =   "Game Directory"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   480
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Half-Life Path"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "frmServerInfo"
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
Server.HLPath = Text1
Server.GamePath = Text6
Server.RconListenPort = Text4
Server.ServerPort = Ts(Val(Text2))
Server.LocalConnectPort = Ts(Val(Text5))
Server.LocalIP = Text3


ServerStart.HLDSPath = Text7
ServerStart.CommandLine = Text8
ServerStart.HLDSDir = Text9

ServerStart.AutoRestart = IIf(Check2 = 1, True, False)
ServerStart.UseFeature = IIf(Check1 = 1, True, False)

PackageServerInfo
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text3 = Server.LocalIP
Text1 = Server.HLPath
Text6 = Server.GamePath
Text2 = Server.ServerPort
Text4 = Server.RconListenPort
Text5 = Server.LocalConnectPort

Text7 = ServerStart.HLDSPath
Text8 = ServerStart.CommandLine
Text9 = ServerStart.HLDSDir

If ServerStart.AutoRestart Then Check2 = 1
If ServerStart.UseFeature Then Check1 = 1

End Sub

