VERSION 5.00
Begin VB.Form frmGame2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin - Hangman"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "frmGame2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6690
   Begin VB.CommandButton Command1 
      Caption         =   "Next Game"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   2820
      Width           =   1035
   End
   Begin VB.CommandButton ButLet 
      Caption         =   "A"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   1560
      TabIndex        =   7
      Top             =   1920
      Width           =   315
   End
   Begin VB.PictureBox picMan 
      AutoRedraw      =   -1  'True
      Height          =   1695
      Left            =   60
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      Height          =   1755
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   2295
      Begin VB.Label Label1 
         Caption         =   "Playing with:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   180
         Width           =   1815
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
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblScore 
         Caption         =   "Score:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label lblYourScore 
         Caption         =   "Your Score:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   1275
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2160
         Y1              =   1260
         Y2              =   1260
      End
   End
   Begin VB.Label lblDollars 
      Height          =   255
      Left            =   2700
      TabIndex        =   10
      Top             =   2820
      Width           =   3795
   End
   Begin VB.Label lblWord 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2460
      TabIndex        =   9
      Top             =   180
      Width           =   4095
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblClue 
      Caption         =   "Clue:"
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
      Left            =   1620
      TabIndex        =   8
      Top             =   3180
      Width           =   4755
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   3660
      Width           =   6615
   End
End
Attribute VB_Name = "frmGame2"
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


'game specific:
Dim IStarted As Boolean
Dim TheWord As String
Dim Dollars As Integer
Dim StartDollars As Integer

Public Sub StartGame()

If IsOpponent Then
    'tell the other guy i am here
    SendIt "HI", "Ready to play!"
Else

End If

End Sub

Sub LoadLetters()

For i = 1 To 25

    Load ButLet(i)
    
    With ButLet(i)
    
        
        .Caption = Chr(i + 65)
        .Enabled = False
        .Visible = True
        
        If i > 12 Then
            .Top = ButLet(0).Top + ButLet(0).Height
            .Left = ButLet(0).Left + (ButLet(0).Width * (i - 13))
        Else
            .Left = ButLet(0).Left + (ButLet(0).Width * i)
        End If
    
    End With

Next i


End Sub

Sub MyTurn()

    EnableButtons
    MakeWord
    lblStatus = "Select a letter!"
    ShowDollars
    DrawMan

End Sub

Sub IAmStarting()

Command1.Enabled = False
SendIt "IS", ""

a$ = InBox("Please enter a word or a phrase:", "Enter Word")
b$ = InBox("Please enter a clue (optional):", "Enter Clue")
c$ = InBox("Please enter starting number of dollars: ", "Enter Dollars", "7")

agina:

TheWord = a$
h = CheckWord


If h < 4 Then
    MessBox "You must use at least 4 letters!"
    a$ = InBox("Please enter a word or a phrase:", "Enter Word")
    GoTo agina
End If


If b$ = "" Then b$ = "(No Clue)"
If Val(c$) <= 0 Then c$ = "7"
If Val(c$) < 3 Then c$ = "3"
If Val(c$) > 12 Then c$ = "12"


SendIt "WD", a$
SendIt "CU", b$
SendIt "DO", c$

lblClue = "Clue: " + b$


EnableButtons

'tell him he can go
SendIt "YT", b$

lblStatus = "Opponent is Playing."
lf = MakeWord
Dollars = Val(c$)
StartDollars = Dollars

ShowDollars
DrawMan

End Sub

Sub DisableButtons()

For i = 0 To 25
    ButLet(i).Enabled = False
Next i

End Sub

Sub EnableButtons()

For i = 0 To 25
    ButLet(i).Enabled = True
Next i

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
    IStarted = True
    IAmStarting
End If

If a$ = "IS" Then
    lblStatus = "Opponent is deciding word..."
    DisableButtons
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

If a$ = "YT" Then MyTurn
If a$ = "TD" Then TakeDollar


If a$ = "CK" Then
    ButLet(Val(p$)).Enabled = False
    h = MakeWord
End If

If a$ = "IL" Then
    'he lost
    MessBox "You opponent has lost!", , , True
    
    'its my turn now
    lblStatus = "Opponent is deciding word..."
    lblWord = ""
    lblClue = "Clue:"
    Dollars = 0
    picMan.Cls
    DisableButtons
    IStarted = False
End If

If a$ = "IW" Then
    'he lost
    MessBox "You opponent has won!", , , True
    
    'its my turn now
    lblStatus = "Opponent is deciding word..."
    lblWord = ""
    lblClue = "Clue:"
    Dollars = 0
    picMan.Cls
    DisableButtons
    IStarted = False
    
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

Sub TakeDollar()

Dollars = Dollars - 1

'redraw the hangman
ShowDollars

DrawMan

End Sub

Sub DrawMan()
'draw order:
'  1    2     3       4      5      6
'head, body, L leg, R leg, L arm, R arm,
'  7       8       9     10     11     12
'L hand, R hand, mouth, L eye, R eye, nose

'Dim B(1 To 12)
'
'If StartDollars = 3 Then
'    B(1) = 1: B(2) = 1: B(3) = 1: B(4) = 1: B(5) = 2: B(6) = 2
'    B(7) = 2: B(8) = 2: B(9) = 3: B(10) = 3: B(11) = 3: B(12) = 3
'End If
'
'If StartDollars = 4 Then
'    B(1) = 1: B(2) = 1: B(3) = 1: B(4) = 1: B(5) = 2: B(6) = 2
'    B(7) = 3: B(8) = 3: B(9) = 3: B(10) = 4: B(11) = 4: B(12) = 3
'End If
'
'If StartDollars = 5 Then
'    B(1) = 1: B(2) = 1: B(3) = 1: B(4) = 1: B(5) = 2: B(6) = 2
'    B(7) = 2: B(8) = 2: B(9) = 3: B(10) = 3: B(11) = 3: B(12) = 3
'End If

c = StartDollars - Dollars

m = Int(picMan.Width / Screen.TwipsPerPixelX / 2)
picMan.Cls

'frame

picMan.Line (m, 10)-(m, 4)
picMan.Line (m, 4)-(10, 4)
picMan.Line (10, 4)-(10, 100)
picMan.Line (5, 100)-(m + 30, 100)


If c >= 1 Or c = StartDollars Then
    picMan.Circle (m, 17), 7
End If

If c >= 2 Or c = StartDollars Then
    picMan.Line (m, 24)-(m, 50)
End If

If c >= 3 Or c = StartDollars Then
    picMan.Line (m, 50)-(m - 10, 60)
End If

If c >= 4 Or c = StartDollars Then
    picMan.Line (m, 50)-(m + 10, 60)
End If

If c >= 5 Or c = StartDollars Then
    picMan.Line (m, 30)-(m - 5, 36)
End If

If c >= 6 Or c = StartDollars Then
    picMan.Line (m, 30)-(m + 5, 36)
End If

If c >= 7 Or c = StartDollars Then
    picMan.Circle (m - 5, 36), 2
End If

If c >= 8 Or c = StartDollars Then
    picMan.Circle (m + 5, 36), 2
End If

If c >= 9 Or c = StartDollars Then
    picMan.Line (m - 2, 19)-(m + 2, 19)
End If

If c >= 10 Or c = StartDollars Then
    picMan.Circle (m - 3, 13), 1
End If

If c >= 11 Or c = StartDollars Then
    picMan.Circle (m + 3, 13), 1
End If

If c >= 12 Or c = StartDollars Then
    picMan.Line (m, 15)-(m, 17)
End If



End Sub

Sub ShowDollars()

lblDollars = "Dollars Left: " + Ts(Dollars)

End Sub

Sub ClickOn(Num As Integer)

'disable the button

lf1 = MakeWord
ButLet(Num).Enabled = False
lf2 = MakeWord

'tell him i clicked here
SendIt "CK", Ts(Num)

'check if i clicked a letter that wasnt in it
If lf2 = lf1 Then ' i did
    
    SendIt "TD", ""
    TakeDollar
    
    'see if we lost
    
    If Dollars <= 0 Then
        'we lost
        
        SendIt "IL", ""
        MyScore = MyScore - 1
        SendIt "MY", Ts(MyScore)
        ShowScore
        lblWord = TheWord
        MessBox "You have lost the game!", , , True
        
        DoEvents
        IStarted = True
        Command1.Enabled = True
        lblStatus = "Click NEXT GAME to begin a new game!"
    End If
End If

'check if I won
If lf2 = 0 And Dollars > 0 Then
    'i won! yay!
    
    MyScore = MyScore + 1
    SendIt "MY", Ts(MyScore)
    SendIt "IW", ""
    ShowScore
    lblWord = TheWord
    
    MessBox "You have won the game!", , , True
    
    IStarted = True
    Command1.Enabled = True
    lblStatus = "Click NEXT GAME to begin a new game!"
End If

End Sub

Sub ShowScore()

lblYourScore = "Your Score: " + Ts(MyScore)
End Sub

Function MakeWord() As Integer
For i = 1 To Len(TheWord)
    
    a$ = Mid(TheWord, i, 1)
    b = Asc(UCase(a$)) - 65
    If b >= 0 And b <= 25 Then
        If ButLet(b).Enabled = True Then
            w$ = w$ + "-"
            d = d + 1
        Else
            w$ = w$ + a$
        End If
    Else
        w$ = w$ + a$
    End If
Next i

lblWord = w$
MakeWord = d

End Function

Function CheckWord() As Integer
For i = 1 To Len(TheWord)
    
    a$ = Mid(TheWord, i, 1)
    b = Asc(UCase(a$)) - 65
    If b >= 0 And b <= 25 Then
        d = d + 1
    End If
Next i

CheckWord = d
End Function


Private Sub SendIt(Cde As String, Params As String)

SendGamePacket Cde, Params, OppIndex, GameID, RemoteGameID

End Sub

Private Sub ButLet_Click(Index As Integer)

If IStarted = False Then ClickOn Index

End Sub

Private Sub Command1_Click()

IAmStarting

End Sub

Private Sub Form_Load()

LoadLetters

End Sub

Private Sub Form_Unload(Cancel As Integer)

SendIt "QU", ""
RemoveGameData GameDataNum

End Sub
