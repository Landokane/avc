VERSION 5.00
Begin VB.Form frmReal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Real Players"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "frmReal.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command12 
      Caption         =   "Debug3"
      CausesValidation=   0   'False
      Height          =   435
      Left            =   2880
      TabIndex        =   39
      Top             =   4140
      Width           =   1035
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Debug2"
      Height          =   435
      Left            =   1920
      TabIndex        =   38
      Top             =   4140
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Delete Temp"
      Height          =   435
      Left            =   0
      TabIndex        =   33
      Top             =   4140
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Debug"
      Height          =   435
      Left            =   1260
      TabIndex        =   32
      Top             =   4140
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Search"
      Height          =   435
      Left            =   1920
      TabIndex        =   31
      Top             =   3660
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Settings"
      Height          =   4575
      Left            =   4320
      TabIndex        =   11
      Top             =   0
      Width           =   3675
      Begin VB.CheckBox Flag4 
         Caption         =   "Custom Flag 4"
         Height          =   255
         Left            =   1920
         TabIndex        =   37
         Top             =   4260
         Width           =   1695
      End
      Begin VB.CheckBox Flag3 
         Caption         =   "Custom Flag 3"
         Height          =   255
         Left            =   1920
         TabIndex        =   36
         Top             =   4020
         Width           =   1695
      End
      Begin VB.CheckBox Flag2 
         Caption         =   "Custom Flag 2"
         Height          =   255
         Left            =   60
         TabIndex        =   35
         Top             =   4260
         Width           =   1815
      End
      Begin VB.CheckBox Flag1 
         Caption         =   "Custom Flag 1"
         Height          =   255
         Left            =   60
         TabIndex        =   34
         Top             =   4020
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1500
         TabIndex        =   24
         Top             =   2160
         Width           =   2115
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   60
         TabIndex        =   23
         Top             =   360
         Width           =   3555
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1620
         Width           =   3555
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1020
         Width           =   3555
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Protect this Name"
         Height          =   255
         Left            =   60
         TabIndex        =   20
         Top             =   3000
         Width           =   1755
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Force Player to use Real Name"
         Height          =   255
         Left            =   60
         TabIndex        =   19
         Top             =   2520
         Width           =   2595
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Update Name"
         Height          =   255
         Left            =   2340
         TabIndex        =   18
         Top             =   720
         Width           =   1275
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Announce Real Name if join name differs"
         Height          =   255
         Left            =   60
         TabIndex        =   17
         Top             =   2760
         Width           =   3495
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Always ID to Join Name"
         Height          =   255
         Left            =   60
         TabIndex        =   16
         Top             =   3240
         Width           =   2055
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Not allowed to start kick votes"
         Height          =   255
         Left            =   60
         TabIndex        =   15
         Top             =   3480
         Width           =   2535
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Temporary Mode"
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   3720
         Width           =   2115
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2160
         Width           =   675
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   780
         TabIndex        =   12
         Top             =   2160
         Width           =   675
      End
      Begin VB.Line Line1 
         X1              =   60
         X2              =   3600
         Y1              =   3990
         Y2              =   3990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "UniqueID's"
         Height          =   195
         Left            =   1500
         TabIndex        =   30
         Top             =   1980
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Last Known Name"
         Height          =   195
         Left            =   60
         TabIndex        =   29
         Top             =   780
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Real Name"
         Height          =   195
         Left            =   60
         TabIndex        =   28
         Top             =   180
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Last Time Seen"
         Height          =   195
         Left            =   60
         TabIndex        =   27
         Top             =   1380
         Width           =   1110
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Connects"
         Height          =   195
         Left            =   60
         TabIndex        =   26
         Top             =   1980
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Points"
         Height          =   195
         Left            =   780
         TabIndex        =   25
         Top             =   1980
         Width           =   435
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Close"
      Height          =   255
      Left            =   4260
      TabIndex        =   10
      Top             =   5880
      Width           =   1035
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   4680
      Width           =   4215
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Sort By Name"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   0
      Value           =   -1  'True
      Width           =   1395
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Sort By Connects"
      Height          =   255
      Left            =   1380
      TabIndex        =   7
      Top             =   0
      Width           =   1635
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Sort By Points"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   1395
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Merge"
      Height          =   435
      Left            =   1260
      TabIndex        =   5
      Top             =   3660
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   435
      Left            =   3600
      TabIndex        =   4
      Top             =   3660
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   2880
      TabIndex        =   3
      Top             =   3660
      Width           =   675
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   435
      Left            =   600
      TabIndex        =   2
      Top             =   3660
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   3660
      Width           =   555
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmReal"
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

Dim Merge1 As Integer
Dim ListSel As Integer
Dim Chng As Integer



Private Sub UpdateList()

DoEvents

a = List1.ListIndex
List1.Clear

If Option1 Then
    For i = 1 To NumRealPlayers

        be = Len(Ts(RealPlayers(i).Points))
        If be > maxnum Then maxnum = be


    Next i

End If

For i = 1 To NumRealPlayers
    
    RealPlayers(i).UniqueID = Trim(RealPlayers(i).UniqueID)
    If Len(RealPlayers(i).UniqueID) > 2 Then
        If Right(RealPlayers(i).UniqueID, 1) <> ";" Then RealPlayers(i).UniqueID = RealPlayers(i).UniqueID + ";"
        If InStr(1, RealPlayers(i).UniqueID, ";;") Then RealPlayers(i).UniqueID = ReplaceString(RealPlayers(i).UniqueID, ";;", ";")
        
        ' Check to see if the last one is right.
        
again:
        e = InStrRev(RealPlayers(i).UniqueID, ";", -1)
        If e > 0 Then f = InStrRev(RealPlayers(i).UniqueID, ";", e - 1)
        
        If e > 0 And f > 0 Then
            b$ = Trim(Mid(RealPlayers(i).UniqueID, f + 1, e - f - 1))
            g = InStrRev(RealPlayers(i).UniqueID, ";", f - 1)
            
            b2$ = Trim(Mid(RealPlayers(i).UniqueID, g + 1, f - g - 1))
            
            If b2$ = b$ Then 'the last two are the same
            
                b3$ = Left(RealPlayers(i).UniqueID, f)
                RealPlayers(i).UniqueID = b3$
                numdone = numdone + 1
                GoTo again
            
            End If
        End If
        
    End If

    d$ = RealPlayers(i).RealName
    If CheckBit2(RealPlayers(i).Flags, 5) Then d$ = " <TEMP>    " + d$: numtm = numtm + 1
    If Option1 Then d$ = Numberize(Ts(RealPlayers(i).Points), CInt(maxnum)) + " - " + d$
    If Option2 Then d$ = Numberize(Ts(Val(RealPlayers(i).TimesSeen)), 4) + " - " + d$
    
    If Val(RealPlayers(i).TimesSeen) > 100 Then numrg = numrg + 1
    
    List1.AddItem d$
    List1.ItemData(List1.NewIndex) = i
    If RealPlayers(i).LastTime > 0 Then kk1 = kk1 + 1
    
    If FindReal <> "" Then
        If InStr(1, RealPlayers(i).UniqueID, FindReal) Then kk = i
    End If
Next i
If numdone > 0 Then MessBox "Fixed " + Ts(numdone) + " UIDs!"

FindReal = "JACK"
If a < List1.ListCount Then List1.ListIndex = a: ShowPlayer
If a >= List1.ListCount Then List1.ListIndex = List1.ListCount - 1: ShowPlayer

If FindReal = "JACK" Then FindReal = ""


Chng = 1
If kk > 0 Then
    For i = 0 To List1.ListCount - 1
        If List1.ItemData(i) = kk Then j = i: Exit For
    Next i
    
    If j > 0 Then List1.ListIndex = j
    ShowPlayer
    DoEvents
End If
Chng = 0
FindReal = ""

ListSel = List1.ListIndex

Me.Caption = "Real Players    " + Ts(NumRealPlayers) + " total, " + Ts(numtm) + " temporary, " + Ts(numrg) + " regulars."

End Sub

Private Sub Check1_Click()
'SetPlayer
End Sub

Private Sub Check2_Click()
'SetPlayer
End Sub

Private Sub Check3_Click()
'SetPlayer
End Sub

Private Sub Command1_Click()

'add

n$ = InBox("Enter Real Name:", "New Real Player", "")
n$ = Trim(n$)

u$ = InBox("Enter UniqueID:", "New Real Player", "")
u$ = Trim(u$)

If n$ = "" Then Exit Sub
If u$ = "" Then Exit Sub

For i = 1 To NumRealPlayers
    If n$ = RealPlayers(i).RealName Then Exit Sub
Next i

NumRealPlayers = NumRealPlayers + 1
ReDim Preserve RealPlayers(0 To NumRealPlayers)

b = NumRealPlayers

RealPlayers(b).LastName = ""
RealPlayers(b).Flags = 0
RealPlayers(b).LastTime = 0
RealPlayers(b).RealName = n$
RealPlayers(b).UniqueID = u$
RealPlayers(b).TimesSeen = "0"

UpdateList

End Sub

Private Sub Command10_Click()

'delete all temporary realplayers

b = MessBox("Delete all temporary RealPlayers?", vbYesNo + vbQuestion, "Delete Temporaries")


If b = vbYes Then

    i = 1
    Do
        If CheckBit2(RealPlayers(i).Flags, 5) Then
            'remove
                
            RemoveReal i
            i = i - 1
        End If
        
        i = i + 1
    
    Loop Until i > NumRealPlayers

    UpdateList

End If



End Sub

Private Sub Command11_Click()

Dim UnId() As String
List2.Clear

For i = 1 To NumRealPlayers
    
    a$ = RealPlayers(i).UniqueID
    UnId = Split(a$, ";")
    nm = UBound(UnId)
    
    b$ = ""
    For j = 0 To (nm - 1)
        c$ = Trim(UnId(j))
        If Left(c$, 1) = Chr(34) Then c$ = Right(c$, Len(c$) - 1)
        If Right(c$, 1) = Chr(34) Then c$ = Left(c$, Len(c$) - 1)
                
        ' do we keep this ID?
        
        If nm > 1 Then
            'if its 6 length, its prolly a new one... add it.
            If Len(c$) = 6 And j < (nm - 1) Then
                 b$ = b$ + Chr(34) + c$ + Chr(34) + "; "
                List2.AddItem RealPlayers(i).RealName + " - " + c$
                
                

                For k = 0 To List1.ListCount - 1
                    If List1.ItemData(k) = i Then m = k: Exit For
                Next k
                
                List2.ItemData(List2.NewIndex) = m
                
            ElseIf j = (nm - 1) Then
                 b$ = b$ + Chr(34) + c$ + Chr(34) + "; "
            End If
        Else
            'its the only one... use it.
            
            b$ = b$ + Chr(34) + c$ + Chr(34) + "; "
        End If
    Next j
        
    b$ = Trim(b$)
    RealPlayers(i).UniqueID = b$
    tot = tot + nm
Next i

MessBox Ts(tot) + " done!"

Me.Height = 6525

End Sub

Private Sub Command12_Click()

'agn:

'Dim TheDate As Date

'TheDate = CDate("March 1, 2001")

For i = 1 To NumRealPlayers

    RealPlayers(i).Points = "0"

 '   If RealPlayers(i).LastTime < TheDate Then
     ' remove
         
  '      RemoveReal i
   '     n = n + 1
 '       GoTo agn:
  '  End If
Next

UpdateList

'MessBox "Removed " + Ts(n) + " realplayers!"


End Sub

Private Sub Command2_Click()

a = List1.ListIndex
If a = -1 Then Exit Sub

'delete
b = MessBox("Are you sure you want to remove this real player?", vbYesNo + vbQuestion, "Delete Real Player")

If b = vbYes Then

    c = List1.ItemData(a)
    
    RemoveReal c

    UpdateList
    
End If



End Sub

Private Sub RemoveReal(c)
    
NumRealPlayers = NumRealPlayers - 1

For i = c To NumRealPlayers

    RealPlayers(i).LastName = RealPlayers(i + 1).LastName
    RealPlayers(i).RealName = RealPlayers(i + 1).RealName
    RealPlayers(i).UniqueID = RealPlayers(i + 1).UniqueID
    RealPlayers(i).Flags = RealPlayers(i + 1).Flags
    RealPlayers(i).LastTime = RealPlayers(i + 1).LastTime
    RealPlayers(i).Points = RealPlayers(i + 1).Points
    RealPlayers(i).TimesSeen = RealPlayers(i + 1).TimesSeen
    
Next i

ReDim Preserve RealPlayers(0 To NumRealPlayers)

End Sub

Private Sub Command3_Click()

a$ = InBox("Search for ID / Name:", "RealPlayer Search", "")

If a$ = "" Then Exit Sub

List2.Clear

For i = 1 To NumRealPlayers

    If InStr(1, RealPlayers(i).UniqueID, a$) Then
        List2.AddItem RealPlayers(i).RealName
        
        For j = 0 To List1.ListCount - 1
            If List1.ItemData(j) = i Then k = j: Exit For
        Next j
        List2.ItemData(List2.NewIndex) = k
    End If
Next i

For i = 1 To NumRealPlayers

    If InStr(1, LCase(RealPlayers(i).RealName), LCase(a$)) Then
        List2.AddItem RealPlayers(i).RealName
        For j = 0 To List1.ListCount - 1
            If List1.ItemData(j) = i Then k = j: Exit For
        Next j
        List2.ItemData(List2.NewIndex) = k
    End If
Next i

For i = 1 To NumRealPlayers

    If InStr(1, LCase(RealPlayers(i).LastName), LCase(a$)) Then
        List2.AddItem RealPlayers(i).RealName
        For j = 0 To List1.ListCount - 1
            If List1.ItemData(j) = i Then k = j: Exit For
        Next j
        List2.ItemData(List2.NewIndex) = k
    End If
Next i

Me.Height = 6525


End Sub

Private Sub SetPlayer()

If Chng = 1 Then Exit Sub
a = ListSel
If a = -1 Then Exit Sub

e = List1.ItemData(a)

RealPlayers(e).UniqueID = Text3
RealPlayers(e).Points = Ts(Val(Text6))

b = 0
If Check1.Value = 1 Then b = b + 1
If Check2.Value = 1 Then b = b + 2 ^ 1
If Check3.Value = 1 Then b = b + 2 ^ 2
If Check4.Value = 1 Then b = b + 2 ^ 3
If Check5.Value = 1 Then b = b + 2 ^ 4
If Check6.Value = 1 Then b = b + 2 ^ 5

If Flag1 = 1 Then b = b + 2 ^ 6
If Flag2 = 1 Then b = b + 2 ^ 7
If Flag3 = 1 Then b = b + 2 ^ 8
If Flag4 = 1 Then b = b + 2 ^ 9


RealPlayers(e).Flags = b


If Text1 <> RealPlayers(e).RealName Then
    RealPlayers(e).RealName = Text1
    'FindReal = RealPlayers(e).UniqueID
    upd = 1
End If

If upd = 1 Then UpdateList

End Sub


Private Sub Command4_Click()
Unload Me

End Sub

Private Sub Command5_Click()
SetPlayer

PackageRealPlayers
Unload Me

End Sub

Private Sub Command6_Click()

a = List1.ListIndex
If a = -1 Then Exit Sub
e = List1.ItemData(a)

If Merge1 = 0 Then
    Merge1 = e
    
    Command6.Caption = "With..."
    Me.Caption = "Select RealPlayer to merge with"
Else
    Me.Caption = "Real Players"
    'merge:
    RealPlayers(Merge1).UniqueID = RealPlayers(Merge1).UniqueID + RealPlayers(e).UniqueID + "; "
    RealPlayers(Merge1).Points = Ts(Val(RealPlayers(Merge1).Points) + Val(RealPlayers(e).Points))
    RealPlayers(Merge1).TimesSeen = Ts(Val(RealPlayers(Merge1).TimesSeen) + Val(RealPlayers(e).TimesSeen))

    'remove old
    RemoveReal e
    UpdateList
    Merge1 = 0
    Command6.Caption = "Merge"
End If

End Sub

Private Sub Command7_Click()
Text1 = Text2

End Sub

Private Sub Command8_Click()
Me.Height = 4980

End Sub

Private Sub Command9_Click()

List2.Clear

For i = 1 To NumRealPlayers
    
    e = 0
    n = 0
    Do
        e = InStr(e + 1, RealPlayers(i).UniqueID, ";")
        If e > 0 Then n = n + 1
    Loop Until e = 0
    
    If n > 3 Then
        List2.AddItem Ts(n) + " - " + RealPlayers(i).RealName
        
        For j = 0 To List1.ListCount - 1
            If List1.ItemData(j) = i Then k = j: Exit For
        Next j
        List2.ItemData(List2.NewIndex) = k
    End If
Next i


Me.Height = 6525

End Sub

Private Sub Form_GotFocus()
ListSel = -1

If FindReal <> "" Then UpdateList

End Sub

Private Sub Form_Load()
UpdateList
Check2.Enabled = DllEnabled
Check1.Enabled = DllEnabled

If CustomFlag1 <> "" Then Flag1.Caption = CustomFlag1
If CustomFlag2 <> "" Then Flag2.Caption = CustomFlag2
If CustomFlag3 <> "" Then Flag3.Caption = CustomFlag3
If CustomFlag4 <> "" Then Flag4.Caption = CustomFlag4



End Sub

Sub ShowPlayer(Optional listitm As Integer)

a = List1.ListIndex
If a = -1 Then Exit Sub

e = List1.ItemData(a)

If listitm Then e = listitm


Text1 = RealPlayers(e).RealName
Text2 = RealPlayers(e).LastName
Text3 = RealPlayers(e).UniqueID
Text5 = RealPlayers(e).TimesSeen
Text6 = RealPlayers(e).Points
If RealPlayers(e).LastTime <> 0 Then Text4 = Format(RealPlayers(e).LastTime, "ddd, mmm d yyyy, hh:mm:ss AMPM") Else Text4 = "Unknown"

If CheckBit2(RealPlayers(e).Flags, 0) = True Then Check1.Value = 1
If CheckBit2(RealPlayers(e).Flags, 1) = True Then Check2.Value = 1
If CheckBit2(RealPlayers(e).Flags, 2) = True Then Check3.Value = 1
If CheckBit2(RealPlayers(e).Flags, 3) = True Then Check4.Value = 1
If CheckBit2(RealPlayers(e).Flags, 4) = True Then Check5.Value = 1
If CheckBit2(RealPlayers(e).Flags, 5) = True Then Check6.Value = 1
If CheckBit2(RealPlayers(e).Flags, 0) = False Then Check1.Value = 0
If CheckBit2(RealPlayers(e).Flags, 1) = False Then Check2.Value = 0
If CheckBit2(RealPlayers(e).Flags, 2) = False Then Check3.Value = 0
If CheckBit2(RealPlayers(e).Flags, 3) = False Then Check4.Value = 0
If CheckBit2(RealPlayers(e).Flags, 4) = False Then Check5.Value = 0
If CheckBit2(RealPlayers(e).Flags, 5) = False Then Check6.Value = 0

'flag1 = 6, flag2 = 7, flag3 = 8, flag4 = 9
Flag1 = 0
Flag2 = 0
Flag3 = 0
Flag4 = 0

If CheckBit2(RealPlayers(e).Flags, 6) = True Then Flag1 = 1
If CheckBit2(RealPlayers(e).Flags, 7) = True Then Flag2 = 1
If CheckBit2(RealPlayers(e).Flags, 8) = True Then Flag3 = 1
If CheckBit2(RealPlayers(e).Flags, 9) = True Then Flag4 = 1

'FindReal = RealPlayers(e).UniqueID

End Sub

Private Sub List1_Click()
If FindReal <> "" Then Exit Sub

SetPlayer

ListSel = List1.ListIndex
ShowPlayer

End Sub

Private Sub List2_Click()

a = List2.ListIndex
If a = -1 Then Exit Sub

e = List2.ItemData(a)

'ShowPlayer CInt(e)
List1.ListIndex = e

End Sub

Private Sub Option1_Click()
UpdateList

End Sub

Private Sub Option2_Click()
UpdateList

End Sub

Private Sub Option3_Click()
UpdateList

End Sub
