VERSION 5.00
Begin VB.Form frmHalfConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configure Half-Life"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "frmHalfConfig.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4605
   Begin VB.CheckBox Check2 
      Caption         =   "Automatically activate ""In-Game"" mode when joining a game"
      Height          =   435
      Left            =   60
      TabIndex        =   7
      Top             =   1920
      Width           =   4035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2340
      TabIndex        =   6
      Top             =   2520
      Width           =   2235
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   2520
      Width           =   2235
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Quit Server Assistant Client upon joining game"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   1620
      Width           =   3555
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   60
      TabIndex        =   3
      Top             =   1260
      Width           =   4515
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Text            =   "C:\Sierra\Half-Life\hl.exe"
      Top             =   300
      Width           =   4515
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Extra Command-Line Arguements"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   660
      Width           =   4500
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Half-Life Executable Path"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1800
   End
End
Attribute VB_Name = "frmHalfConfig"
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

Text1 = Trim(Text1)

If CheckForFile(Text1) = False Then
    MsgBox "The HL Executable was not found."
    Exit Sub
End If

a$ = App.Path + "\assisthl.dat"

h = FreeFile
If CheckForFile(a$) Then Kill a$

    

Dim erg(1 To 10) As String

erg(1) = Text1
erg(2) = Text2
erg(3) = Ts(Check1)
erg(4) = Ts(Check2)

HLEXEPath = erg(1)
HLExtraArgs = erg(2)
HLQuitSA = Check1
HLSetAway = Check2

Open a$ For Binary As h
    Put #h, , erg
    
Close h

Unload Me


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
Label2 = "Extra Command-Line Arguements:" + vbCrLf + "Note: You need not specify -game, +connect, or -console."
Text1 = HLEXEPath
Text2 = HLExtraArgs
Check1 = Val(HLQuitSA)
Check2 = Val(HLSetAway)

End Sub
