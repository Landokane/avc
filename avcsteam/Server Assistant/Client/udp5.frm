VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form5 
   Caption         =   "Update Remote Server"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4110
   Icon            =   "udp5.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   2580
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Update Options"
      Height          =   1215
      Left            =   60
      TabIndex        =   2
      Top             =   960
      Width           =   3975
      Begin VB.CommandButton Command4 
         Caption         =   "Upload EXE"
         Height          =   435
         Left            =   1980
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Upload Commands"
         Height          =   435
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1875
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Download Commands"
         Height          =   435
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1875
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check Server"
      Height          =   435
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   1875
   End
   Begin MSWinsockLib.Winsock UDP1 
      Left            =   3660
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   26000
      LocalPort       =   26001
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   2280
      Width           =   3975
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

UDP1.SendData Chr(254)

'b$ = Form1.Text5
'a$ = CompileData

'If a$ = b$ Then MsgBox "Identical!"
'If a$ <> b$ Then MsgBox "NOT Identical!"

End Sub

Private Sub Command3_Click()

'UDP1.RemoteHost = Text1
'
'a$ = CompileData
'
'UDP1.SendData a$
'UDP1.SendData Chr(254)

End Sub

Private Sub Command4_Click()

'a$ = CompileEXE


'UDP1.SendData Chr(254)

End Sub

Private Sub Form_Load()

'UDP1.RemoteHost = Server.LogIP
'UDP1.RemotePort = Server.LocalFilePort
'UDP1.LocalPort = Server.RemoteFilePort



End Sub
