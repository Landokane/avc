Attribute VB_Name = "GameModule"
' **********************************************************
' **********************************************************
' **********************************************************
'             DO NOT MODIFY THE CODE IN THIS FILE
'             DO NOT MODIFY THE CODE IN THIS FILE
'             DO NOT MODIFY THE CODE IN THIS FILE
' **********************************************************
' **********************************************************
' **********************************************************

'Game Data type
Public Type typGameData
    GameForm As Object
    Index As Integer
    Opponent As String
End Type

Public Type typPacketBuffer
    Cde As String
    Params As String
    TimeStarted As Long
    ToWho As Integer
    GameID As Integer
    HisGameID As Integer
End Type


'Game Data array
Public GameData() As typGameData
Public NumGames As Integer

'Packet Buffer
Public PacketBuffer() As typPacketBuffer



Sub Main()

'Main ini sub.

ReDim PacketBuffer(0 To 0)
Form1.Show

End Sub

Sub CheckPacketBuffer()

'check the buffer and send any that have been "lagged"

n = UBound(PacketBuffer)
b = Form1.Slider1.Value / 1000

i = 1
Do Until i > n
        
    
    
    tm = Timer - PacketBuffer(i).TimeStarted

    If tm < 0 Or tm > b Then 'send
        
        Cde$ = PacketBuffer(i).Cde
        Params$ = PacketBuffer(i).Params
        ToWho = PacketBuffer(i).ToWho
        GameID = PacketBuffer(i).GameID
        HisGameID = PacketBuffer(i).HisGameID
        
        'Remove the entry
        For j = i To n - 1
            PacketBuffer(j).Cde = PacketBuffer(j + 1).Cde
            PacketBuffer(j).Params = PacketBuffer(j + 1).Params
            PacketBuffer(j).ToWho = PacketBuffer(j + 1).ToWho
            PacketBuffer(j).GameID = PacketBuffer(j + 1).GameID
            PacketBuffer(j).HisGameID = PacketBuffer(j + 1).HisGameID
            PacketBuffer(j).TimeStarted = PacketBuffer(j + 1).TimeStarted
        Next j
        
        n = n - 1
        ReDim Preserve PacketBuffer(0 To n)
        i = i - 1
        
        ActualSendGamePacket Cde$, Params$, CInt(ToWho), CInt(GameID), HisGameID
        
        
    End If
    
    i = i + 1
    If i > n Then Exit Do
Loop



End Sub


Sub UnpackageGameRequest()


    NumGames = NumGames + 1
    ReDim Preserve GameData(0 To NumGames)
    
    Randomize
    t = Int(Rnd * 30000) + 1

    GameData(NumGames).Index = 2
    GameData(NumGames).Opponent = "Right Side"
    
    Set GameData(NumGames).GameForm = New frmGame2
    
    GameData(NumGames).GameForm.Opponent = "Left Side"
    GameData(NumGames).GameForm.OppIndex = 2
    GameData(NumGames).GameForm.GameID = t
    GameData(NumGames).GameForm.RemoteGameID = 1000
    GameData(NumGames).GameForm.IsOpponent = True
    GameData(NumGames).GameForm.GameDataNum = NumGames
    GameData(NumGames).GameForm.Show

w = Screen.Width * (3 / 4)
h = Screen.Height * (2 / 3)
l = w - (GameData(NumGames).GameForm.Width / 2)
t = h - (GameData(NumGames).GameForm.Height / 2)
GameData(NumGames).GameForm.Move l, t

    
    GameData(NumGames).GameForm.StartGame


End Sub

Sub RemoveGameData(Num)

For i = Num To NumGames - 1

    Set GameData(i).GameForm = GameData(i + 1).GameForm
    GameData(i).Index = GameData(i + 1).Index
    GameData(i).Opponent = GameData(i + 1).Opponent
    GameData(i).GameForm.GameDataNum = i

Next i
NumGames = NumGames - 1

ReDim Preserve GameData(0 To NumGames)

End Sub

Sub StartAGame()

b$ = "Left Side"

NumGames = NumGames + 1
ReDim Preserve GameData(0 To NumGames)
Randomize
t = 1000

Set GameData(NumGames).GameForm = New frmGame2
GameData(NumGames).GameForm.GameID = t
GameData(NumGames).GameForm.GameDataNum = NumGames
GameData(NumGames).GameForm.Opponent = b$
GameData(NumGames).GameForm.Show
GameData(NumGames).GameForm.lblStatus = "Waiting for " + b$ + "'s response..."


w = Screen.Width * (1 / 4)
h = Screen.Height * (2 / 3)
l = w - (GameData(NumGames).GameForm.Width / 2)
t = h - (GameData(NumGames).GameForm.Height / 2)
GameData(NumGames).GameForm.Move l, t

UnpackageGameRequest

End Sub

Sub SendGamePacket(Cde As String, Params As String, ToWho As Integer, GameID As Integer, HisGameID)

'Buffer it.
n = UBound(PacketBuffer) + 1
ReDim Preserve PacketBuffer(0 To n)

PacketBuffer(n).Cde = Cde
PacketBuffer(n).Params = Params
PacketBuffer(n).ToWho = ToWho
PacketBuffer(n).GameID = GameID
PacketBuffer(n).HisGameID = HisGameID
PacketBuffer(n).TimeStarted = Timer

End Sub


Sub ActualSendGamePacket(Cde As String, Params As String, ToWho As Integer, GameID As Integer, HisGameID)

a$ = Chr(244) + Chr(244) + Chr(244) + Chr(245) + Cde + Chr(245) + Params + Chr(245) + Chr(243) + Chr(243) + Chr(243)

WhoFrom$ = "Right Side"
IndexFrom = 2
If GameID = 1000 Then WhoFrom$ = "Left Side": IndexFrom = 1

For i = 1 To NumGames
    If GameData(i).GameForm.GameID = HisGameID Then
        If GameID <> 0 Then GameData(i).GameForm.RemoteGameID = GameID
                
        GameData(i).GameForm.GameInterprit (a$), (CInt(IndexFrom)), (WhoFrom$)
        Exit Sub
    End If
Next i

End Sub

' Needed functions

Function MessBox(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String, Optional ShowMode As Boolean, Optional TimeToShow As Integer) As Long

MessBox = MsgBox(Prompt, Buttons, Title)

End Function

Function InBox(Prompt As String, Optional Title As String, Optional Default As String) As String

InBox = InputBox(Prompt, Title, Default)

End Function

Function Ts(a) As String
    Ts = Trim(Str(a))
End Function

