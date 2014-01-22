VERSION 5.00
Begin VB.Form frmCarRace 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin - Car Racer"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   Icon            =   "frmCarRace.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   8070
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer tmrRollDice 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3480
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Stop"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      Height          =   1755
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.Label lblYourScore2 
         Caption         =   "0"
         Height          =   195
         Left            =   1080
         TabIndex        =   11
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lblScore2 
         Caption         =   "0"
         Height          =   195
         Left            =   720
         TabIndex        =   10
         Top             =   960
         Width           =   555
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2160
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Label lblYourScore 
         Caption         =   "Your Score:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label lblScore 
         Caption         =   "Score:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   555
      End
      Begin VB.Label lblWho 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Playing with:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Image picCar2 
      Height          =   480
      Left            =   240
      Picture         =   "frmCarRace.frx":0442
      Top             =   3240
      Width           =   480
   End
   Begin VB.Image picCar1 
      Height          =   480
      Left            =   240
      Picture         =   "frmCarRace.frx":0884
      Top             =   2520
      Width           =   480
   End
   Begin VB.Label lblYourTurn 
      Alignment       =   2  'Center
      Caption         =   "It's Your Turn!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FFFF&
      BorderStyle     =   2  'Dash
      X1              =   120
      X2              =   7920
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FFFF&
      X1              =   120
      X2              =   7920
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      X1              =   120
      X2              =   7920
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   8
      Top             =   1800
      Width           =   5175
   End
   Begin VB.Label lblDice 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Move Space:"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   360
      Top             =   2280
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   7680
      Top             =   2280
      Width           =   135
   End
   Begin VB.Shape shpFinish 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   6960
      Top             =   2280
      Width           =   135
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   120
      Top             =   2280
      Width           =   7815
   End
End
Attribute VB_Name = "frmCarRace"
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

Public GameDataNum As Integer
Public OppIndex As Integer
Public Opponent As String
Public IsOpponent As Boolean 'when false, means I'm the host
Public GameID As Integer
Public RemoteGameID As Integer
Dim MyScore As Integer

Dim NumOfTurns As Integer

Public Sub StartGame()

If IsOpponent Then
    'tell the other guy i am here
    SendIt "HI", "Ready to play!"
    lblStatus = "You move after opponent moves."
Else

End If

End Sub

Public Sub GameInterprit(Txt As String, IndexFrom As Integer, WhoFrom As String)

lblWho = WhoFrom

'gets the stuff
'(244)(244)(244)(245)[CODE](245)[PARAMS](245)(243)(243)(243)

e = InStr(1, Txt, Chr(245))
f = InStr(e + 1, Txt, Chr(245))
Debug.Print TheWord
If e > 0 And f > e And f > 0 Then
    'code
    a$ = Mid(Txt, e + 1, f - e - 1)
    e = f
    f = InStrRev(Txt, Chr(245))
    
    If e > 0 And f > e And f > 0 Then
        'params
        p$ = Mid(Txt, e + 1, f - e - 1)
        'decode the encoded shtuff
    End If
End If

'GAME SPECIFIC PACKETS

If a$ = "HI" Then
    lblStatus = "You start moving first."
    NumOfTurns = 0
    tmrRollDice.Enabled = True
    Command1.Enabled = True
    lblYourTurn.Visible = True
    picCar1.Left = 240
    picCar2.Left = 240
End If

If a$ = "WD" Then
    TheWord = p$
End If

If a$ = "CU" Then
    lblClue = "Clue: " + p$
End If

If a$ = "DO" Then
    Dollars = Val(p$)
    StartDollars = Dollars
End If

If a$ = "MY" Then 'update on HIS score
    lblScore = "Score: " + p$
End If

If a$ = "OW" Then
    lblScore2 = p$
    MessBox "You have lost the race!", , "You Lose", , 5
    lblStatus = "You have lost the race!"
End If

If a$ = "MS" Then
    If IsOpponent = True Then
        picCar1.Left = picCar1.Left + (120 * p$)
        lblStatus = WhoFrom & " moved " & p$ & " spaces."
    Else
        picCar2.Left = picCar2.Left + (120 * p$)
        lblStatus = WhoFrom & " moved " & p$ & " spaces."
    End If
End If

If a$ = "YT" Then
    tmrRollDice.Enabled = True
    Command1.Enabled = True
    lblYourTurn.Visible = True
End If
        

'standard messages
If a$ = "QU" Then
    MessBox "Opponent has left the game!"
    Unload Me
End If

If a$ = "N!" Then
    MessBox "Opponent does not wish to play!"
    Unload Me
End If

End Sub

Sub MoveMyCar()

If IsOpponent = False Then
    picCar1.Left = picCar1.Left + (120 * lblDice)
    SendIt "MS", lblDice
    lblStatus = "You moved " & lblDice & " spaces."
Else
    picCar2.Left = picCar2.Left + (120 * lblDice)
    SendIt "MS", lblDice
    lblStatus = "You moved " & lblDice & " spaces."
End If

If IsOpponent = True Then
    If picCar2.Left >= shpFinish.Left Then
        lblStatus = "You have won the race!"
        MessBox "You have won the race!", , "You Win", , 5
        lblYourScore2 = lblYourScore2 + NumOfTurns
        SendIt "OW", lblYourScore2
        cmdReset.Visible = True
    Else
        SendIt "YT", ""
    End If
Else
    If picCar1.Left >= shpFinish.Left Then
        lblStatus = "You have won the race!"
        MessBox "You have won the race!", , "You Win", , 5
        lblYourScore2 = lblYourScore2 + NumOfTurns
        SendIt "OW", lblYourScore2
        cmdReset.Visible = True
    Else
        SendIt "YT", ""
    End If
End If

End Sub

Sub ResetGame()

SendIt "HI", ""

tmrRollDice.Enabled = False
Command1.Enabled = False
lblYourTurn.Visible = False
picCar1.Left = 240
picCar2.Left = 240

NumOfTurns = 0

End Sub

Private Sub SendIt(Cde As String, Params As String)

SendGamePacket Cde, Params, OppIndex, GameID, RemoteGameID

End Sub

Private Sub cmdReset_Click()

ResetGame
cmdReset.Visible = False

End Sub

Private Sub Command1_Click()

Command1.Enabled = False
tmrRollDice.Enabled = False
lblYourTurn.Visible = False

NumOfTurns = NumOfTurns + 1

MoveMyCar

End Sub

Private Sub Form_Load()

NumOfTurns = 0
Command1.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

SendIt "QU", ""
RemoveGameData GameDataNum

End Sub

Private Sub tmrRollDice_Timer()

minr = 1
maxr = 10
randno = minr + Fix(Rnd * (maxr - minr + 1))
lblDice = randno

End Sub
