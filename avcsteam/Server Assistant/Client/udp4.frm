VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   Icon            =   "udp4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3510
   Begin VB.CommandButton Command3 
      Caption         =   "Remove"
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   660
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Save Password"
      Height          =   195
      Left            =   840
      TabIndex        =   11
      Top             =   1860
      Width           =   1755
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   840
      TabIndex        =   10
      Top             =   300
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   840
      TabIndex        =   7
      Top             =   1140
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   2340
      TabIndex        =   9
      Top             =   2160
      Width           =   1155
   End
   Begin VB.TextBox Text4 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1500
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   840
      TabIndex        =   5
      Text            =   "26000"
      Top             =   660
      Width           =   915
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   1560
      Width           =   690
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Username"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Port"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   720
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "IP"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   300
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Connect to:"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   825
   End
End
Attribute VB_Name = "frmConnect"
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

Private Type LastConnects
    IP As String
    Port As Long
    Name As String
    PassWord As String
    SavePass As Boolean
End Type

Private LastCon() As LastConnects
Private NumLastCon As Integer
'Private InfoShow As Integer

Function GetIP() As String

a$ = Combo1.Text
e = InStr(a$, ":")

If e Then
    a$ = Left(a$, e - 1)
End If

GetIP = a$

End Function

Private Sub Check1_Click()



a$ = GetIP + ":" + Ts(Val(Text2))
For i = 1 To NumLastCon
    If LastCon(i).IP + ":" + Ts(LastCon(i).Port) = a$ Then j = i: Exit For
Next i

If j > 0 Then
    If Check1 = 1 Then LastCon(j).SavePass = True
    If Check1 = 0 Then LastCon(j).SavePass = False
    
End If

End Sub

Private Sub Combo1_Change()

If InStr(1, Combo1.Text, ":") Then
    a = Combo1.ListIndex
    If a = -1 Then Exit Sub
    e = Combo1.ItemData(a)
    
    
    Combo1.Text = LastCon(e).IP
End If



End Sub

Private Sub Combo1_Click()
'Show this information

a = Combo1.ListIndex
If a = -1 Then Exit Sub
e = Combo1.ItemData(a)

DoEvents

Text2 = Ts(LastCon(e).Port)
Text3 = LastCon(e).Name
If LastCon(e).SavePass Then
    Text4 = LastCon(e).PassWord
    Check1 = 1
Else
    Check1 = 0
    Text4 = ""
End If

DoEvents
'Combo1.Text = LastCon(e).IP
'Combo1.Text = "hehe"

'MessBox Combo1.Text + " - " + LastCon(e).IP

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    Dim CB As Long
    Dim FindString As String

    If KeyAscii < 32 Or KeyAscii > 127 Then Exit Sub

    If Combo1.SelLength = 0 Then
        FindString = Combo1.Text & Chr$(KeyAscii)
    Else
        FindString = Left$(Combo1.Text, Combo1.SelStart) & Chr$(KeyAscii)
    End If

    CB = SendMessage(Combo1.hwnd, CB_FINDSTRING, -1, ByVal FindString)

    If CB <> CB_ERR Then
        Combo1.ListIndex = CB
        Combo1.SelStart = Len(FindString)
        Combo1.SelLength = Len(Combo1.Text) - Combo1.SelStart
        KeyAscii = 0
    End If
End Sub

Sub SaveLastCon()


If CheckForFile(DataFile3) Then Kill DataFile3

Open DataFile3 For Binary As #1
    Put #1, , NumLastCon
    Put #1, , LastCon
Close #1

End Sub

Sub LoadLastCon()

If CheckForFile(DataFile3) Then
    Open DataFile3 For Binary As #1
        Get #1, , NumLastCon
        ReDim LastCon(0 To NumLastCon)
        Get #1, , LastCon
    Close #1
End If

End Sub


Private Sub Combo1_LostFocus()

If InStr(1, Combo1.Text, ":") Then
    a = Combo1.ListIndex
    If a = -1 Then Exit Sub
    e = Combo1.ItemData(a)
    
    Combo1.Text = LastCon(e).IP
End If

End Sub

Private Sub Command1_Click()

CB = SendMessage(Combo1.hwnd, CB_FINDSTRING, -1, ByVal Combo1.Text)

If CB = CB_ERR Then
    frmConnect.Combo1.AddItem Combo1.Text
End If

a$ = Combo1.Text
Dim X() As String
X = Split(a$, ":")
    

j = 0
a$ = GetIP + ":" + Ts(Val(Text2))
For i = 1 To NumLastCon
    If LastCon(i).IP + ":" + Ts(LastCon(i).Port) = a$ Then j = i: Exit For
Next i

If j > 0 Then
    If LastCon(j).SavePass Then LastCon(j).PassWord = Text4
    LastCon(j).Name = Text3
    LastCon(j).Port = Val(Text2)
    
    'swap with last one
    
    If j < NumLastCon Then
            
        Swap LastCon(NumLastCon).IP, LastCon(j).IP
        Swap LastCon(NumLastCon).Port, LastCon(j).Port
        Swap LastCon(NumLastCon).Name, LastCon(j).Name
        Swap LastCon(NumLastCon).PassWord, LastCon(j).PassWord
        Swap LastCon(NumLastCon).SavePass, LastCon(j).SavePass
    
    End If
Else
    'add
    NumLastCon = NumLastCon + 1
    ReDim Preserve LastCon(0 To NumLastCon)
    j = NumLastCon
    If Check1.Value = 1 Then LastCon(j).PassWord = Text4: LastCon(j).SavePass = True
    LastCon(j).Name = Text3
    LastCon(j).Port = Val(Text2)

    LastCon(j).IP = X(0)
End If


SaveLastCon

AttemptConnect X(0), Text2, Text3, Text4


Unload Me

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()

a$ = GetIP + ":" + Ts(Val(Text2))
For i = 1 To NumLastCon
    If LastCon(i).IP + ":" + Ts(LastCon(i).Port) = a$ Then j = i: Exit For
Next i

If j > 0 Then
    'remove entry
    
    For i = j To NumLastCon - 1
        
        LastCon(i).IP = LastCon(i + 1).IP
        LastCon(i).Name = LastCon(i + 1).Name
        LastCon(i).Port = LastCon(i + 1).Port
        LastCon(i).PassWord = LastCon(i + 1).PassWord
        LastCon(i).SavePass = LastCon(i + 1).SavePass
    
    Next i
    
    NumLastCon = NumLastCon - 1
    
    Combo1.Clear
    
    For i = 1 To NumLastCon
        Combo1.AddItem LastCon(i).IP + ":" + Ts(LastCon(i).Port)
        Combo1.ItemData(Combo1.NewIndex) = i
    Next i

    Combo1.Text = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Check1 = 0
    

End If

End Sub

Private Sub Form_Load()


If CheckForFile(DataFile2) Then

    Dim Combo2() As String
    ReDim Combo2(0 To 1000)
    frmConnect.Combo1.Clear

    Open DataFile2 For Binary As #1
        Get #1, , Combo2
    Close #1
    
    Kill DataFile2
    SaveLastCon

    For i = 0 To UBound(Combo2)
        NumLastCon = i
        ReDim Preserve LastCon(0 To i)
        
        LastCon(i).IP = Combo2(i)
    Next i
End If

'Load data into combobox
LoadLastCon

For i = 1 To NumLastCon
    Combo1.AddItem LastCon(i).IP + ":" + Ts(LastCon(i).Port)
    Combo1.ItemData(Combo1.NewIndex) = i
Next i

If frmConnect.Combo1.ListCount > 0 Then frmConnect.Combo1.ListIndex = frmConnect.Combo1.ListCount - 1

Me.Show

If Text4 = "" Then Text4.SetFocus
If Text3 = "" Then Text3.SetFocus

End Sub


Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
a$ = GetIP + ":" + Ts(Val(Text2))
j = 0
For i = 1 To NumLastCon
    If LastCon(i).IP + ":" + Ts(LastCon(i).Port) = a$ Then j = i: Exit For
Next i

If j > 0 Then
    LastCon(j).Port = Val(Text2)
End If

End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
a$ = GetIP + ":" + Ts(Val(Text2))
For i = 1 To NumLastCon
    If LastCon(i).IP + ":" + Ts(LastCon(i).Port) = a$ Then j = i: Exit For
Next i

If j > 0 Then
    LastCon(j).Name = Text3
End If

End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
a$ = GetIP + ":" + Ts(Val(Text2))
For i = 1 To NumLastCon
    If LastCon(i).IP + ":" + Ts(LastCon(i).Port) = a$ Then j = i: Exit For
Next i

If j > 0 Then
    If LastCon(j).SavePass Then LastCon(j).PassWord = Text4
End If
End Sub
