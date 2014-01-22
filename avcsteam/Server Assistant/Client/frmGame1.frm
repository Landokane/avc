VERSION 5.00
Begin VB.Form frmGame1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin - Tic Tac Toe"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "frmGame1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4800
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1995
      Begin VB.CommandButton ClikBut 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   8
         Left            =   1320
         TabIndex        =   9
         Top             =   1440
         Width           =   555
      End
      Begin VB.CommandButton ClikBut 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   7
         Left            =   720
         TabIndex        =   8
         Top             =   1440
         Width           =   555
      End
      Begin VB.CommandButton ClikBut 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   555
      End
      Begin VB.CommandButton ClikBut 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   5
         Left            =   1320
         TabIndex        =   6
         Top             =   840
         Width           =   555
      End
      Begin VB.CommandButton ClikBut 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   4
         Left            =   720
         TabIndex        =   5
         Top             =   840
         Width           =   555
      End
      Begin VB.CommandButton ClikBut 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   555
      End
      Begin VB.CommandButton ClikBut 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   2
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   555
      End
      Begin VB.CommandButton ClikBut 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   1
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   555
      End
      Begin VB.CommandButton ClikBut 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   555
      End
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
      Left            =   120
      TabIndex        =   14
      Top             =   2220
      Width           =   4575
   End
   Begin VB.Line Line1 
      X1              =   2160
      X2              =   4200
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblYourScore 
      Caption         =   "Your Score:"
      Height          =   195
      Left            =   2280
      TabIndex        =   13
      Top             =   1380
      Width           =   1275
   End
   Begin VB.Label lblScore 
      Caption         =   "Score:"
      Height          =   195
      Left            =   2280
      TabIndex        =   12
      Top             =   900
      Width           =   1275
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
      Left            =   2280
      TabIndex        =   11
      Top             =   420
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Playing with:"
      Height          =   195
      Left            =   2160
      TabIndex        =   10
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmGame1"
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


'game specific:
Dim MyScore As Integer
Dim IAmLetter As String
Dim IStarted As Boolean


Public Sub StartGame()

If IsOpponent Then
    
    'tell the other guy i am here
    SendIt "HI", "Ready to play!"
    Randomize
    v = Int(Rnd * 2) + 1
    If v = 1 Then 'i am X
        Me.Caption = "Admin - Tic Tac Toe - You are X"
        SendIt "UR", "O"
        IAmLetter = "X"
        CheckIfX
    Else
        Me.Caption = "Admin - Tic Tac Toe - You are O"
        SendIt "UR", "X"
        IAmLetter = "O"
    End If
    
Else



End If

End Sub



Sub CheckIfX()
'person who is X starts
    If IAmLetter = "X" Then
        SendIt "NY", "asd"
        MyTurn
        IStarted = True
    End If

End Sub

Sub MyTurn()

    lblStatus = "It's YOUR turn!"
    
    lblWho = Opponent
    lblYourScore = "Your Score: " + Ts(MyScore)
    
    For i = 0 To 8
        ClikBut(i).Enabled = True
    Next i


End Sub

Sub NotMyTurn()

    lblStatus = "Waiting for opponent to move..."
    
    lblWho = Opponent
    lblYourScore = "Your Score: " + Ts(MyScore)
    
    For i = 0 To 8
        ClikBut(i).Enabled = False
    Next i


End Sub

Public Sub GameInterprit(Txt As String, IndexFrom As Integer, WhoFrom As String)

'gets the stuff
'(244)(244)(244)(245)[CODE](245)[PARAMS](245)(243)(243)(243)

e = InStr(1, Txt, Chr(245))
f = InStr(e + 1, Txt, Chr(245))

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
    'MsgBox "Got a hi packet"
    
    'start the game
    'pick a number.

    
    
    
End If

If a$ = "MY" Then 'update on HIS score
    lblScore = "Score: " + p$
End If

If a$ = "UR" Then 'what letter am I
    IAmLetter = p$
    Me.Caption = "Admin - Tic Tac Toe - You are " + p$
    CheckIfX
End If

If a$ = "NY" Then NotMyTurn
If a$ = "YT" Then MyTurn

If a$ = "CK" Then
    n = Val(p$)
    If IAmLetter = "X" Then d$ = "O"
    If IAmLetter = "O" Then d$ = "X"
    
    ClikBut(n).Caption = d$
    
    If CheckIfWin(d$) Then
        
        'i lost
        MessBox "You lost the game!"
        ClearBoard
        
    Else
        MyTurn

    End If
    
End If

If a$ = "CG" Then
    MessBox "Cat's Game!"
    ClearBoard
End If

If a$ = "QU" Then
    MessBox "Opponent has left the game!"
    Unload Me
End If

If a$ = "N!" Then
    MessBox "Opponent does not wish to play!"
    Unload Me
End If

End Sub

Sub ClickOn(Num As Integer)

    If ClikBut(Num).Caption <> "" Then Exit Sub
    
    ClikBut(Num).Caption = IAmLetter
    
    For i = 0 To 8
        ClikBut(i).Enabled = False
        If ClikBut(i).Caption <> "" Then j = j + 1
    Next i

    SendIt "CK", Ts(Num)
    
    If CheckIfWin(IAmLetter) Then
        
        'i won
        MessBox "You won the game!"
        ClearBoard
        MyScore = MyScore + 1
        SendIt "MY", Ts(MyScore)
        
        If IStarted Then
            IStarted = False
            NotMyTurn
            SendIt "YT", ""
        Else
            IStarted = True
            MyTurn
            SendIt "NY", ""
        End If
    ElseIf j = 9 Then
    
        MessBox "Cat's Game!"
        SendIt "CG", ""
        ClearBoard
        Randomize
        If IStarted Then
            IStarted = False
            NotMyTurn
            SendIt "YT", ""
        Else
            IStarted = True
            MyTurn
            SendIt "NY", ""
        End If

    Else
        NotMyTurn
    End If
    
    
    
    
End Sub

Sub ClearBoard()

    For i = 0 To 8
        ClikBut(i).Caption = ""
    Next i

End Sub

Function CheckIfWin(Letter As String) As Boolean

'0 1 2
'3 4 5
'6 7 8


If ClikBut(0).Caption = Letter And ClikBut(1).Caption = Letter And ClikBut(2).Caption = Letter Then win = 1
If ClikBut(3).Caption = Letter And ClikBut(4).Caption = Letter And ClikBut(5).Caption = Letter Then win = 1
If ClikBut(6).Caption = Letter And ClikBut(7).Caption = Letter And ClikBut(8).Caption = Letter Then win = 1

If ClikBut(0).Caption = Letter And ClikBut(3).Caption = Letter And ClikBut(6).Caption = Letter Then win = 1
If ClikBut(1).Caption = Letter And ClikBut(4).Caption = Letter And ClikBut(7).Caption = Letter Then win = 1
If ClikBut(2).Caption = Letter And ClikBut(5).Caption = Letter And ClikBut(8).Caption = Letter Then win = 1

If ClikBut(0).Caption = Letter And ClikBut(4).Caption = Letter And ClikBut(8).Caption = Letter Then win = 1
If ClikBut(2).Caption = Letter And ClikBut(4).Caption = Letter And ClikBut(6).Caption = Letter Then win = 1

If win = 1 Then CheckIfWin = True


End Function


Private Sub Label2_Click()

End Sub

Private Sub ClikBut_Click(Index As Integer)

ClickOn Index

End Sub

Private Sub Command1_Click()


End Sub

Private Sub SendIt(Cde As String, Params As String)

SendGamePacket Cde, Params, OppIndex, GameID, RemoteGameID

End Sub

Private Sub Form_Unload(Cancel As Integer)

SendIt "QU", ""

RemoveGameData GameDataNum


End Sub
