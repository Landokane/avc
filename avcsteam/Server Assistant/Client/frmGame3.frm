VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGame3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin - Battleship"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   Icon            =   "frmGame3.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   416
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   527
   Begin VB.CommandButton Command4 
      Caption         =   "Begin Next Game"
      Height          =   555
      Left            =   3000
      TabIndex        =   13
      Top             =   4020
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Battle Command - Ship Placement"
      Height          =   1875
      Left            =   600
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CommandButton Command3 
         Caption         =   "Done!"
         Height          =   315
         Left            =   1740
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   60
         ScaleHeight     =   105
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   109
         TabIndex        =   11
         Top             =   240
         Width           =   1635
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next Ship"
         Height          =   315
         Left            =   1740
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Rotate"
         Height          =   315
         Left            =   1740
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   4020
      ScaleHeight     =   257
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   7
      Top             =   60
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   60
      ScaleHeight     =   257
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   6
      Top             =   60
      Width           =   3855
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5640
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame3.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame3.frx":091A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame3.frx":0CAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame3.frx":0FFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame3.frx":137E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame3.frx":170E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame3.frx":1AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame3.frx":1E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame3.frx":21A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame3.frx":2506
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame3.frx":287E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame3.frx":2BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame3.frx":2F6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame3.frx":32E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame3.frx":3666
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame3.frx":39E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame3.frx":3D72
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame3.frx":4102
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame3.frx":4492
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGame3.frx":4822
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      Height          =   1755
      Left            =   5580
      TabIndex        =   1
      Top             =   4020
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "ITS YOUR TURN!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   60
      TabIndex        =   14
      Top             =   5100
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Line Line2 
      X1              =   264
      X2              =   264
      Y1              =   4
      Y2              =   260
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
      TabIndex        =   0
      Top             =   5880
      Width           =   7815
   End
End
Attribute VB_Name = "frmGame3"
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

Private Declare Function BitBlt Lib "gdi32" _
         (ByVal hDestDC As Long, _
          ByVal X As Long, _
          ByVal Y As Long, _
          ByVal nWidth As Long, _
          ByVal nHeight As Long, _
          ByVal hSrcDC As Long, _
          ByVal xSrc As Long, _
          ByVal ySrc As Long, _
          ByVal dwRop As Long) As Long

'game specific:
Dim IStarted As Boolean
Dim MyField(1 To 16, 1 To 16) As Integer
Dim HisField(1 To 16, 1 To 16) As Integer
Dim MyBoats(1 To 5, 1 To 3) As Integer  ' 2nd param: 1 = x, 2 = y, 3 = orientation
Dim MyFieldBoats(1 To 16, 1 To 16) As Integer
Dim HisFieldBoats(1 To 16, 1 To 16) As Integer


Dim PlaceMode As Boolean
Dim CurrentlyPlacing As Integer '1 to 5
Dim CurrentRotation As Integer '1 to 2
Dim BoatData(1 To 5, 1 To 5) As Integer
Dim RotTable(0 To 20) As Integer
Dim GetLen(1 To 5) As Integer
Dim OppDonePlacing As Boolean
Dim ItsMyTurn As Boolean
Dim BombImages(0 To 20) As Integer
Dim GameInProgress As Boolean
'ORIENTDATION:
'1 = LEFT TO RIGHT       OXXXX
'2 = TOP TO BOTTOM.    O
'                      X
'                      X
'data

Sub InitializeField()

For X = 1 To 16
    For Y = 1 To 16
        MyField(X, Y) = 3 'water
        HisField(X, Y) = 3 'water
        HisFieldBoats(X, Y) = 0
        MyFieldBoats(X, Y) = 0
        
        If X < 6 And Y < 4 Then MyBoats(X, Y) = 0
        
    Next Y
Next X



End Sub

Sub DrawMyField()

For X = 1 To 16
    For Y = 1 To 16
        
        'draw coords
        X1 = (X - 1) * 16
        Y1 = (Y - 1) * 16
                
        n = MyField(X, Y)
        
        ImageList1.ListImages(n).Draw Picture1.hDC, X1, Y1

    Next Y
Next X
Picture1.Refresh
End Sub

Sub DrawHisField(Optional Mde As Integer)

For X = 1 To 16
    For Y = 1 To 16
        
        'draw coords
        X1 = (X - 1) * 16
        Y1 = (Y - 1) * 16
                
        n = HisField(X, Y)
        
        'make sure we dont draw boats :)
        If n = 3 Or n = 5 Or n = 6 Or (n >= 8 And n <= 13) Then
            m = n
        Else
            m = 3
        End If
        'm = n
        
        If Mde = 1 Then m = n
        
        ImageList1.ListImages(m).Draw Picture2.hDC, X1, Y1
    Next Y
Next X
Picture2.Refresh
End Sub

Public Sub StartGame()

If IsOpponent Then
    'tell the other guy i am here
    GameInProgress = True
    SendIt "HI", "Ready to play!"
    PlaceBoats
Else

End If

End Sub

Sub ContinueGame()

'tell opponent we are done
PackageMyField 2
PackageMyField 1

'see if he is done
If OppDonePlacing Then
    If IStarted Then MyTurn
Else
    lblStatus = "Waiting for opponent to finish placing ships..."
End If

End Sub

Sub MyTurn()
SendIt "MT", ""

lblStatus = "Select a bomb location on the RIGHT grid!"
ItsMyTurn = True
Label2.Visible = True

End Sub

Sub BombLocation(X, Y)

    If GameInProgress = False Then Exit Sub
    
    Dim Bombed As Boolean
    Dim Hit As Boolean
    
    'make sure we havent bombed here before
    n = HisField(X, Y)
    
    If n <> 6 And n <> 5 And (n < 8 Or n > 13) Then
        'Valid bomb location!
        p = HisFieldBoats(X, Y)
        
        'See if we actually hit anything here.
        If p > 0 Then
            'yes.
            'get boat number.
            bn = Val(Left(Ts(p), 1))
            'get part number
            pn = Val(Mid(Ts(p), 2, 1))
            'set this to be "bombed"
            p = (bn * 100) + (pn * 10) + 1
            
            HisFieldBoats(X, Y) = p
            
            'now, set field to look bombed
            HisField(X, Y) = 6
            Bombed = True
            Hit = True
            'now, check if this boat is dead.
            
            For X1 = 1 To 16
                For Y1 = 1 To 16
                    'look for "alive" sections of boat.
                    bn2 = Val(Left(Ts(HisFieldBoats(X1, Y1)), 1))
                    bm2 = Val(Mid(Ts(HisFieldBoats(X1, Y1)), 3, 1))
                    
                    If bn2 = bn And bm2 = 0 Then alive = 1
                    If bm2 = 1 Then anyalive = anyalive + 1
                Next Y1
            Next X1
            
            If alive = 1 Then 'boat not destroyed
            
            
            Else
                'boat is fully destroyed
                'make the map show the whole boat
                bombedall = bn
            
                For X1 = 1 To 16
                    For Y1 = 1 To 16
                        'look for "alive" sections of boat.
                        bn2 = Val(Left(Ts(HisFieldBoats(X1, Y1)), 1))
                        pn2 = Val(Mid(Ts(HisFieldBoats(X1, Y1)), 2, 1))
                        
                        If bn2 = bn Then
                            pn2 = pn2 + 7
                            HisField(X1, Y1) = pn2
                        End If
                    Next Y1
                Next X1
            End If
        Else
            'no, put the splash picture
            HisField(X, Y) = 5
            Bombed = True
        End If
    Else
        'we already shot here
    
    End If

DrawHisField

a$ = Ts(X)
If Len(a$) = 1 Then a$ = "0" + a$
b$ = Ts(Y)
If Len(b$) = 1 Then b$ = "0" + b$
a$ = a$ + b$

If Bombed = True And Hit = True Then
    PlayWaveRes 101
ElseIf Bombed = True Then
    PlayWaveRes 102
End If



If Bombed = True And bombedall = 0 Then 'tell him we bombed
    ItsMyTurn = False
    SendIt "BM", a$
    lblStatus = "Opponent is deciding bomb location..."
    Label2.Visible = False
End If

If Bombed = True And bombedall > 0 Then 'tell him we bombed and destroyed
    ItsMyTurn = False
    
    TellBomb bombedall, 1
    
    SendIt "BD", a$ + Ts(bombedall)
    lblStatus = "Opponent is deciding bomb location..."
    Label2.Visible = False
End If



If Bombed = True And Hit = True Then 'all enemy ships dead
    If anyalive = 18 Then
    ItsMyTurn = False
    MyScore = MyScore + 1
    ShowScore
    
    SendIt "MY", Ts(MyScore)
    SendIt "IW", ""
    GameInProgress = False
    Label2.Visible = False
    MessBox "You have won the game!"
    
    If IStarted Then Command4.Visible = True
    
    End If
End If



End Sub

Sub TellBomb(n, m)

If n = 1 Then a$ = "Minesweeper!"
If n = 2 Then a$ = "Destroyer!"
If n = 3 Then a$ = "Cruiser!"
If n = 4 Then a$ = "Battleship!"
If n = 5 Then a$ = "Carrier!"
    
If m = 1 Then b$ = "You sunk the " + a$
If m = 2 Then b$ = "Your opponent sunk your " + a$

MessBox b$, , , True

End Sub

Function NextShip() As Boolean

'find a ship that hasnt been placed

If CurrentlyPlacing >= 5 Then CurrentlyPlacing = 0
CurrentlyPlacing = CurrentlyPlacing + 1

DrawBoat

End Function

Sub PlaceBoats()

lblStatus = "Select Location for SHIPS by clicking on LEFT GRID"
CurrentlyPlacing = 1
CurrentRotation = 1
PlaceMode = True
Frame2.Visible = True
DrawBoat

End Sub

Sub DrawBoat()

m = CurrentlyPlacing
o = CurrentRotation

'draw
Picture3.Cls

For i = 1 To 5
    n = BoatData(m, i)
        
    If n <> 0 Then
        
        X = (i - 1) * 16
        If o = 2 Then
            n = RotTable(n)
            Y = X
            X = 0
        End If
        ImageList1.ListImages(n).Draw Picture3.hDC, X, Y
    End If
Next i

Picture3.Refresh

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
    f = InStr(e + 1, Txt, Chr(245))
    
    If e > 0 And f > e And f > 0 Then
        'params
        p$ = Mid(Txt, e + 1, f - e - 1)
        'decode the encoded shtuff
    End If
End If

'GAME SPECIFIC PACKETS

If a$ = "HI" Then
    IStarted = True
    GameInProgress = True
    'allow this player to place his boats
    PlaceBoats
End If

If a$ = "NG" Then
    IStarted = True
    GameInProgress = True
    'allow this player to place his boats
    InitializeField
    DrawMyField
    DrawHisField
    PlaceBoats
End If


If a$ = "MF" Then
    'oppenent is giving me his field data
    UnPackageHisField p$, 1
    
    OppDonePlacing = True
    If PlaceMode = True Then
        lblStatus = "Opponent has finished placing ships and is waiting for you!"
    Else
        'start the game here
        If IStarted Then MyTurn
    End If
End If

If a$ = "MB" Then
    'oppenent is giving me his boat data
    UnPackageHisField p$, 2
End If

If a$ = "BM" And GameInProgress = True Then
    'opponent has bombed here
    X = Val(Left(p$, 2))
    Y = Val(Right(p$, 2))
    
    n = MyField(X, Y)
    n = BombImages(n)
    
    If n = 5 Then PlayWaveRes 102
    If n <> 5 Then PlayWaveRes 101
    
    
    MyField(X, Y) = n
    DrawMyField
    
    MyTurn
End If

If a$ = "BD" And GameInProgress = True Then
    'opponent has bombed here, and sunk our ship
    X = Val(Left(p$, 2))
    Y = Val(Mid(p$, 3, 2))
    bn = Val(Mid(p$, 5, 1))
    
    n = MyField(X, Y)
    n = BombImages(n)
    
    If n = 5 Then PlayWaveRes 102
    If n <> 5 Then PlayWaveRes 101
    
    
    MyField(X, Y) = n
    DrawMyField
        
    TellBomb bn, 2
        
    MyTurn
End If

If a$ = "MY" Then 'update on HIS score
    lblScore = "Score: " + p$
End If

If a$ = "IW" Then
    'he won
    If IStarted Then Command4.Visible = True
    MessBox "You have lost the game!"
    GameInProgress = False
    DrawHisField 1
End If
    
If a$ = "YT" Then MyTurn
If a$ = "MT" Then lblStatus = "Opponent is deciding bomb location..."

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

Sub ShowScore()

lblYourScore = "Your Score: " + Ts(MyScore)
End Sub


Private Sub SendIt(Cde As String, Params As String)

SendGamePacket Cde, Params, OppIndex, GameID, RemoteGameID

End Sub

Private Sub Command1_Click()

NextShip


End Sub

Private Sub Command2_Click()
If CurrentRotation = 1 Then
    CurrentRotation = 2
Else
    CurrentRotation = 1
End If
DrawBoat

End Sub

Private Sub Command3_Click()


For i = 1 To 5
    If MyBoats(i, 1) = 0 Then j = i: Exit For
Next i

If j > 0 Then
    MessBox "You must place all 5 ships!"
Else
    'all placed
    PlaceMode = False
    Frame2.Visible = False
    ContinueGame
End If

End Sub

Private Sub Command4_Click()
Command4.Visible = False
IStarted = False

SendIt "NG", ""

IStarted = True
GameInProgress = True
'allow this player to place his boats
InitializeField
DrawMyField
DrawHisField
PlaceBoats


End Sub


Private Sub Form_Load()
LoadBoatData
InitializeField
DrawMyField
DrawHisField

End Sub

Private Sub Form_Unload(Cancel As Integer)

SendIt "QU", ""
RemoveGameData GameDataNum

End Sub

Sub LoadBoatData()

n = 1
BoatData(n, 1) = 20
BoatData(n, 2) = 16

n = 2
BoatData(n, 1) = 20
BoatData(n, 2) = 14
BoatData(n, 3) = 16

n = 3
BoatData(n, 1) = 20
BoatData(n, 2) = 7
BoatData(n, 3) = 14
BoatData(n, 4) = 16

n = 4
BoatData(n, 1) = 20
BoatData(n, 2) = 7
BoatData(n, 3) = 15
BoatData(n, 4) = 16

n = 5
BoatData(n, 1) = 20
BoatData(n, 2) = 14
BoatData(n, 3) = 14
BoatData(n, 4) = 14
BoatData(n, 5) = 16


RotTable(7) = 18
RotTable(14) = 19
RotTable(15) = 1
RotTable(16) = 2
RotTable(20) = 17
RotTable(8) = 11
RotTable(9) = 12
RotTable(10) = 13

GetLen(1) = 2
GetLen(2) = 3
GetLen(3) = 4
GetLen(4) = 4
GetLen(5) = 5

BombImages(1) = 12
BombImages(2) = 13
BombImages(3) = 5
BombImages(7) = 9
BombImages(14) = 9
BombImages(15) = 9
BombImages(16) = 10
BombImages(17) = 11
BombImages(18) = 12
BombImages(19) = 12
BombImages(20) = 8


End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'find out what "tile" was clicked on

If X > Picture1.Width Then Exit Sub
If Y > Picture1.Height Then Exit Sub
If X < 0 Then Exit Sub
If Y < 0 Then Exit Sub

X1 = Int(X / 16) + 1
Y1 = Int(Y / 16) + 1

If X1 > 16 Or Y1 > 16 Then Exit Sub

'now what do we do?
If PlaceMode = True Then
    
    'draw the ship here
    PlaceShip X1, Y1


Else


End If


End Sub

Sub PlaceShip(X, Y)

'puts the ship here
n = CurrentlyPlacing

l = GetLen(n)

'check new co-ords are valid
If CurrentRotation = 1 Then
    X2 = X + l - 1
    If X2 > 16 Then Exit Sub
Else
    Y2 = Y + l - 1
    If Y2 > 16 Then Exit Sub
End If




'first, remove the prev location of this boat.

X1 = MyBoats(n, 1)
Y1 = MyBoats(n, 2)

If X1 > 0 And Y1 > 0 Then
    
    For i = 1 To l
        
        X2 = X1
        Y2 = Y1
        
        If MyBoats(n, 3) = 1 Then X2 = X2 + (i - 1)
        If MyBoats(n, 3) = 2 Then Y2 = Y2 + (i - 1)
        
        'Remove
        MyField(X2, Y2) = 3
        
        MyFieldBoats(X2, Y2) = 0
    Next i
End If

'Check if new pos is valid.

bd = 0
For i = 1 To l
    
    X2 = X
    Y2 = Y
    
    If CurrentRotation = 1 Then X2 = X2 + (i - 1)
    If CurrentRotation = 2 Then Y2 = Y2 + (i - 1)
    
    'place
    If MyField(X2, Y2) <> 3 Then bd = 1
Next i

If bd = 0 Then
    X1 = X
    Y1 = Y
End If
 

'now place in new pos.
If X1 > 0 And Y1 > 0 Then
    For i = 1 To l
        
        X2 = X1
        Y2 = Y1
        
        If CurrentRotation = 1 Then X2 = X2 + (i - 1)
        If CurrentRotation = 2 Then Y2 = Y2 + (i - 1)
        
        'place
        m = BoatData(n, i)
        If CurrentRotation = 2 Then m = RotTable(m)
        MyField(X2, Y2) = m
        
        p = BombImages(m)
        p = p - 7
        p = p * 10
        p = p + (100 * n)
        MyFieldBoats(X2, Y2) = p
    Next i
End If
'set

MyBoats(n, 1) = X1
MyBoats(n, 2) = Y1
MyBoats(n, 3) = CurrentRotation
DrawMyField

End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'find out what "tile" was clicked on

If X > Picture1.Width Then Exit Sub
If Y > Picture1.Height Then Exit Sub
If X < 0 Then Exit Sub
If Y < 0 Then Exit Sub

X1 = Int(X / 16) + 1
Y1 = Int(Y / 16) + 1

If X1 > 16 Or Y1 > 16 Then Exit Sub

If ItsMyTurn Then
    'bomb here
    BombLocation X1, Y1

End If

End Sub

Sub PackageMyField(Mde)

'compiles and sends the real player info
'generic format for array items:
'(255)DATA1(254)DATA2(254)DATA3(254)(255)(255)DATA1(254)DATA2(254)DATA3(254)(255) etc

'compile it

For X = 1 To 16
    a$ = a$ + Chr(241)
    For Y = 1 To 16
        If Mde = 1 Then a$ = a$ + Ts(MyField(X, Y)) + Chr(240)
        If Mde = 2 Then a$ = a$ + Ts(MyFieldBoats(X, Y)) + Chr(240)
    Next Y
    a$ = a$ + Chr(241)
Next X

'all set, send it
If Mde = 1 Then SendIt "MF", a$
If Mde = 2 Then SendIt "MB", a$

End Sub

Sub UnPackageHisField(p$, Mde)
'extracts field data from string

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(241))
    f = InStr(e + 1, p$, Chr(241))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(240))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                m$ = Mid(a$, g, h - g)
                
                If Mde = 1 Then HisField(i, j) = Val(m$)
                If Mde = 2 Then HisFieldBoats(i, j) = Val(m$)

            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0
If Mde = 1 Then DrawHisField

End Sub

