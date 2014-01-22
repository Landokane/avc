VERSION 5.00
Begin VB.Form frmSpeech 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Speech"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   Icon            =   "frmSpeech.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command8 
      Caption         =   "Replace All"
      Height          =   315
      Left            =   2760
      TabIndex        =   15
      Top             =   3600
      Width           =   1035
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Search"
      Height          =   315
      Left            =   1800
      TabIndex        =   14
      Top             =   3600
      Width           =   915
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   900
      TabIndex        =   13
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Send"
      Height          =   315
      Left            =   0
      TabIndex        =   12
      Top             =   3600
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   3915
      Left            =   5640
      TabIndex        =   3
      Top             =   -120
      Width           =   4335
      Begin VB.CommandButton Command2 
         Caption         =   "Remove"
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   3180
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   255
         Left            =   3660
         TabIndex        =   10
         Top             =   3180
         Width           =   615
      End
      Begin VB.ListBox List2 
         Height          =   2010
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1140
         Width           =   4215
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   60
         TabIndex        =   6
         Top             =   3540
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   60
         TabIndex        =   4
         Top             =   420
         Width           =   4215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Command"
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Top             =   3300
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Responses"
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   900
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Activation"
         Height          =   195
         Left            =   60
         TabIndex        =   5
         Top             =   180
         Width           =   705
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remove"
      Height          =   315
      Left            =   4020
      TabIndex        =   2
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   315
      Left            =   4920
      TabIndex        =   1
      Top             =   3600
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmSpeech"
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

Dim List1Sel As Integer
Dim List2Sel As Integer


Private Sub ShowItem(Num)

Text1 = Speech(Num).ClientText
Text2 = ""
UpdateList2 Num

End Sub

Sub UpdateList2(Num)

List2Sel = 0
List2.Clear
For i = 1 To Speech(Num).NumAnswers
    List2.AddItem Speech(Num).Answers(i)
    e = List2.NewIndex
    List2.ItemData(e) = i
Next i
Text2 = ""

End Sub

Sub UpdateList1()

List1Sel = 0
a = List1.ListIndex
List1.Clear
For i = 1 To NumSpeech

    g$ = ""
    If Speech(i).NumAnswers = 0 Then g$ = "  "
    List1.AddItem g$ + Speech(i).ClientText
    e = List1.NewIndex
    List1.ItemData(e) = i
Next i
If a < List1.ListCount Then List1.ListIndex = a

End Sub

Sub SaveItem()

If List1Sel = 0 Then Exit Sub

If Speech(List1Sel).ClientText <> Text1 Then
    Speech(List1Sel).ClientText = Text1
    UpdateList1
End If
SaveAnswer

End Sub

Sub SaveAnswer()
If List1Sel = 0 Then Exit Sub
If List2Sel = 0 Then Exit Sub


If Speech(List1Sel).Answers(List2Sel) <> Text2 Then
    Speech(List1Sel).Answers(List2Sel) = Text2
    UpdateList2 List1Sel

End If
End Sub

Private Sub Command1_Click()

If List1Sel = 0 Then Exit Sub

a$ = InBox("Command?", "New Answer", "say AutoAdmin: ")

If a$ = "" Then Exit Sub

Speech(List1Sel).NumAnswers = Speech(List1Sel).NumAnswers + 1
e = Speech(List1Sel).NumAnswers

ReDim Preserve Speech(List1Sel).Answers(0 To e)

Speech(List1Sel).Answers(e) = a$

UpdateList2 List1Sel

End Sub

Private Sub Command2_Click()

'remove an item

a = List2.ListIndex
If a = -1 Then Exit Sub
e = List2.ItemData(a)

If List1Sel = 0 Then Exit Sub

Speech(List1Sel).NumAnswers = Speech(List1Sel).NumAnswers - 1

For i = e To Speech(List1Sel).NumAnswers
    Speech(List1Sel).Answers(i) = Speech(List1Sel).Answers(i + 1)
Next i

ReDim Preserve Speech(List1Sel).Answers(0 To Speech(List1Sel).NumAnswers)

UpdateList2 List1Sel

End Sub

Private Sub Command3_Click()

a$ = InBox("Activation?", "New Speech", "help me")
If a$ = "" Then Exit Sub
NumSpeech = NumSpeech + 1
e = NumSpeech

ReDim Preserve Speech(0 To e)

Speech(e).NumAnswers = 0
Speech(e).ClientText = UCase(a$)

UpdateList1

End Sub

Private Sub Command4_Click()
'remove an item

a = List1.ListIndex
If a = -1 Then Exit Sub
e = List1.ItemData(a)

NumSpeech = NumSpeech - 1
For i = e To NumSpeech
    Speech(i).ClientText = Speech(i + 1).ClientText
    
    ReDim Preserve Speech(i).Answers(0 To Speech(i + 1).NumAnswers)
    For j = 1 To Speech(i + 1).NumAnswers
        Speech(i).Answers(j) = Speech(i + 1).Answers(j)
    Next j
    Speech(i).NumAnswers = Speech(i + 1).NumAnswers
Next i

ReDim Preserve Speech(0 To NumSpeech)

UpdateList1

End Sub

Private Sub Command5_Click()
PackageSpeech
Unload Me

End Sub

Private Sub Command6_Click()
Unload Me

End Sub

Private Sub Command7_Click()

a$ = InBox("Search For:", "Search Speech", "")

If a$ = "" Then Exit Sub



For i = List1.ListIndex + 1 To List1.ListCount - 1
    e = List1.ItemData(i)
    
    mt = 0
    If InStr(1, LCase(Speech(e).ClientText), LCase(a$)) Then mt = 1
    
    For j = 1 To Speech(e).NumAnswers
        If InStr(1, LCase(Speech(e).Answers(j)), LCase(a$)) Then mt = 1
    Next j
    
    If a$ = "notsay" Then
        mt = 0
        For j = 1 To Speech(e).NumAnswers
            If Left(LCase(Speech(e).Answers(j)), 3) <> "say" Then mt = 1
        Next j
    End If
    If mt = 1 Then Exit For

Next i

If mt = 1 Then
    List1.ListIndex = i
Else
    MessBox "Search String not found."
End If

End Sub

Private Sub Command8_Click()

a$ = InBox("Search For:", "Search & Replace Speech", "")
If a$ = "" Then Exit Sub

b$ = InBox("Replace With:", "Search & Replace Speech", "")
If b$ = "" Then Exit Sub

For i = 1 To NumSpeech
    For j = 1 To Speech(i).NumAnswers
        Speech(i).Answers(j) = ReplaceString(Speech(i).Answers(j), a$, b$)
    Next j
Next i

List1Sel = 0
UpdateList1


End Sub

Private Sub Form_Load()
List1Sel = 0
List2Sel = 0
UpdateList1


End Sub

Private Sub List1_Click()

a = List1.ListIndex
If a = -1 Then Exit Sub

e = List1.ItemData(a)
SaveItem

List1Sel = e
List2Sel = 0
ShowItem e

For i = 1 To NumSpeech
    If Speech(i).NumAnswers = 0 Then n = n + 1
    nn = nn + Speech(i).NumAnswers
Next i

Me.Caption = "Speech - " + Ts(NumSpeech) + " Total, " + Ts(n) + " undone, " + Ts(nn) + " answers"
'UpdateList1

End Sub

Private Sub List2_Click()

a = List2.ListIndex
If a = -1 Then Exit Sub

If List1Sel = 0 Then Exit Sub

e = List2.ItemData(a)
SaveAnswer

List2Sel = e

Text2 = Speech(List1Sel).Answers(List2Sel)


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

KeyAscii = Asc(UCase(Chr(KeyAscii)))


End Sub

Private Sub Text1_LostFocus()
SaveItem
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)


If KeyAscii = 13 Then
    If List1Sel = 0 Then Exit Sub
   
    a$ = Text2
    If a$ = "" Then Exit Sub
    If List2.ListIndex <> -1 Then Exit Sub
    
    Speech(List1Sel).NumAnswers = Speech(List1Sel).NumAnswers + 1
    e = Speech(List1Sel).NumAnswers
    
    ReDim Preserve Speech(List1Sel).Answers(0 To e)
    
    a$ = "say %a " + a$
    
    Speech(List1Sel).Answers(e) = a$
    
    UpdateList2 List1Sel
    'UpdateList1
    KeyAscii = 0
End If

End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)

'Debug.Print KeyCode, Shift

If Shift = 2 Then
    If KeyCode = 78 Then
        e = Text2.SelStart
        If e > 1 Then a$ = Left(Text2, e)
        If e < Len(Text2) Then b$ = Right(Text2, Len(Text2) - e)
        Text2 = a$ + "%n" + b$
        Text2.SelStart = Len(a$ + "%n")
        KeyCode = 0
        Shift = 0
    ElseIf KeyCode = 85 Then
        e = Text2.SelStart
        If e > 1 Then a$ = Left(Text2, e)
        If e < Len(Text2) Then b$ = Right(Text2, Len(Text2) - e)
        Text2 = a$ + "%u" + b$
        Text2.SelStart = Len(a$ + "%u")
        KeyCode = 0
        Shift = 0
    End If
End If


End Sub

Private Sub Text2_LostFocus()
SaveAnswer
'UpdateList1

End Sub
