VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Controller"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3720
      Top             =   240
   End
   Begin VB.Frame Frame1 
      Caption         =   "Network Control"
      Height          =   1035
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   4335
      Begin MSComctlLib.Slider Slider1 
         Height          =   435
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   767
         _Version        =   393216
         LargeChange     =   100
         Max             =   2000
         TickFrequency   =   100
      End
      Begin VB.Label Label1 
         Caption         =   "Simulated Lag in Milliseconds:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Initiate a Game"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If NumGames = 0 Then StartAGame

End Sub

Private Sub Form_Load()

w = Screen.Width / 2
h = Screen.Height * (1 / 3)

l = w - (Me.Width / 2)
t = h - (Me.Height / 2)

Me.Move l, t

End Sub

Private Sub Timer1_Timer()
CheckPacketBuffer
End Sub
