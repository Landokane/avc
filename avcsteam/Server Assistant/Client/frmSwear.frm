VERSION 5.00
Begin VB.Form frmSwear 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bad Word List"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   Icon            =   "frmSwear.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "Use"
      Height          =   315
      Left            =   5160
      TabIndex        =   25
      Top             =   420
      Width           =   555
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   60
      Width           =   2115
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   3000
      TabIndex        =   20
      Top             =   1440
      Width           =   2715
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   435
      Left            =   2040
      TabIndex        =   19
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   1320
      TabIndex        =   18
      Top             =   5040
      Width           =   675
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   435
      Left            =   660
      TabIndex        =   17
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   435
      Left            =   0
      TabIndex        =   16
      Top             =   5040
      Width           =   615
   End
   Begin VB.Frame Frame5 
      Caption         =   "Removal"
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   4980
      Width           =   2715
      Begin VB.CheckBox Check 
         Caption         =   "Ban player permanently"
         Height          =   195
         Index           =   6
         Left            =   60
         TabIndex        =   13
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Warning"
      Height          =   495
      Left            =   3000
      TabIndex        =   10
      Top             =   4440
      Width           =   2715
      Begin VB.CheckBox Check 
         Caption         =   "Warn first time, kick second"
         Height          =   195
         Index           =   5
         Left            =   60
         TabIndex        =   11
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Filtering Options"
      Height          =   735
      Left            =   3000
      TabIndex        =   7
      Top             =   3660
      Width           =   2715
      Begin VB.CheckBox Check 
         Caption         =   "Filter non-letters (a#b$c -> abc)"
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   9
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox Check 
         Caption         =   "Use 1337 filter (3 -> E, 4 -> A)"
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   8
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Name Options"
      Height          =   735
      Left            =   3000
      TabIndex        =   5
      Top             =   2880
      Width           =   2715
      Begin VB.CheckBox Check 
         Caption         =   "Remove from NAME rather than kicking"
         Height          =   375
         Index           =   4
         Left            =   60
         TabIndex        =   6
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detection"
      Height          =   1035
      Left            =   3000
      TabIndex        =   2
      Top             =   1800
      Width           =   2715
      Begin VB.CheckBox Check 
         Caption         =   "Replace in SPEECH"
         Height          =   195
         Index           =   7
         Left            =   60
         TabIndex        =   22
         Top             =   240
         Width           =   2535
      End
      Begin VB.CheckBox Check 
         Caption         =   "Disallowed in NAME"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   4
         Top             =   720
         Width           =   2535
      End
      Begin VB.CheckBox Check 
         Caption         =   "Disallowed in SPEECH"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   3
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   2715
   End
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Line Line3 
      X1              =   5700
      X2              =   5100
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Line Line2 
      X1              =   5100
      X2              =   5100
      Y1              =   480
      Y2              =   780
   End
   Begin VB.Line Line1 
      X1              =   3000
      X2              =   5100
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label4 
      Caption         =   "Presets"
      Height          =   255
      Left            =   3000
      TabIndex        =   24
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label3 
      Caption         =   "Replacement"
      Height          =   255
      Left            =   3000
      TabIndex        =   21
      Top             =   1200
      Width           =   2475
   End
   Begin VB.Label Label2 
      Caption         =   "Word List"
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label Label1 
      Caption         =   "Word"
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   600
      Width           =   555
   End
End
Attribute VB_Name = "frmSwear"
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
Dim NoClick As Boolean

Private Type typBadPreset
    NumPreset As Integer
    Name As String
    BadWord() As String
    Replace() As String
End Type

Dim Presets(1 To 2) As typBadPreset

Sub LoadPreset()

nm = 1
Presets(nm).Name = "No Swearing at ALL"
Presets(nm).NumPreset = 29

ReDim Presets(nm).BadWord(1 To Presets(nm).NumPreset)
ReDim Presets(nm).Replace(1 To Presets(nm).NumPreset)

n = 0
n = n + 1
Presets(nm).BadWord(n) = "fuck"
Presets(nm).Replace(n) = "frick"
n = n + 1
Presets(nm).BadWord(n) = "shit"
Presets(nm).Replace(n) = "crap"
n = n + 1
Presets(nm).BadWord(n) = "ass"
Presets(nm).Replace(n) = "caboose"
n = n + 1
Presets(nm).BadWord(n) = "asshole"
Presets(nm).Replace(n) = "tent pole"
n = n + 1
Presets(nm).BadWord(n) = "tits"
Presets(nm).Replace(n) = "eyes"
n = n + 1
Presets(nm).BadWord(n) = "damn"
Presets(nm).Replace(n) = "darn"
n = n + 1
Presets(nm).BadWord(n) = "cum"
Presets(nm).Replace(n) = "stuff"
n = n + 1
Presets(nm).BadWord(n) = "whore"
Presets(nm).Replace(n) = "horse"
n = n + 1
Presets(nm).BadWord(n) = "bitch"
Presets(nm).Replace(n) = "fish"
n = n + 1
Presets(nm).BadWord(n) = "motherfucker"
Presets(nm).Replace(n) = "mamas boy"
n = n + 1
Presets(nm).BadWord(n) = "fucking"
Presets(nm).Replace(n) = "freaking"
n = n + 1
Presets(nm).BadWord(n) = "penis"
Presets(nm).Replace(n) = "larynx"
n = n + 1
Presets(nm).BadWord(n) = "cock"
Presets(nm).Replace(n) = "chicken"
n = n + 1
Presets(nm).BadWord(n) = "fucker"
Presets(nm).Replace(n) = "buster"
n = n + 1
Presets(nm).BadWord(n) = "gayass"
Presets(nm).Replace(n) = "silly"
n = n + 1
Presets(nm).BadWord(n) = "bastard"
Presets(nm).Replace(n) = "batter"
n = n + 1
Presets(nm).BadWord(n) = "wop"
Presets(nm).Replace(n) = "fella"
n = n + 1
Presets(nm).BadWord(n) = "dildo"
Presets(nm).Replace(n) = "hotdog"
n = n + 1
Presets(nm).BadWord(n) = "cunt"
Presets(nm).Replace(n) = "lollypop"
n = n + 1
Presets(nm).BadWord(n) = "goddamn"
Presets(nm).Replace(n) = "goshdarn"
n = n + 1
Presets(nm).BadWord(n) = "damnit"
Presets(nm).Replace(n) = "darnit"
n = n + 1
Presets(nm).BadWord(n) = "dammit"
Presets(nm).Replace(n) = "dangit"
n = n + 1
Presets(nm).BadWord(n) = "bitching"
Presets(nm).Replace(n) = "whining"
n = n + 1
Presets(nm).BadWord(n) = "vagina"
Presets(nm).Replace(n) = "lasagna"
n = n + 1
Presets(nm).BadWord(n) = "blowjob"
Presets(nm).Replace(n) = "dollar"
n = n + 1
Presets(nm).BadWord(n) = "pussy"
Presets(nm).Replace(n) = "toast"
n = n + 1
Presets(nm).BadWord(n) = "nigga"
Presets(nm).Replace(n) = "dude"
n = n + 1
Presets(nm).BadWord(n) = "nigger"
Presets(nm).Replace(n) = "dude"
n = n + 1
Presets(nm).BadWord(n) = "slut"
Presets(nm).Replace(n) = "slot"

nm = 2
Presets(nm).Name = "No Harsh Words"
Presets(nm).NumPreset = 9

ReDim Presets(nm).BadWord(1 To Presets(nm).NumPreset)
ReDim Presets(nm).Replace(1 To Presets(nm).NumPreset)

n = 0
n = n + 1
Presets(nm).BadWord(n) = "fuck"
Presets(nm).Replace(n) = "frick"
n = n + 1
Presets(nm).BadWord(n) = "asshole"
Presets(nm).Replace(n) = "tent pole"
n = n + 1
Presets(nm).BadWord(n) = "motherfucker"
Presets(nm).Replace(n) = "mamas boy"
n = n + 1
Presets(nm).BadWord(n) = "fucking"
Presets(nm).Replace(n) = "freaking"
n = n + 1
Presets(nm).BadWord(n) = "cock"
Presets(nm).Replace(n) = "chicken"
n = n + 1
Presets(nm).BadWord(n) = "fucker"
Presets(nm).Replace(n) = "buster"
n = n + 1
Presets(nm).BadWord(n) = "gayass"
Presets(nm).Replace(n) = "silly"
n = n + 1
Presets(nm).BadWord(n) = "nigga"
Presets(nm).Replace(n) = "dude"
n = n + 1
Presets(nm).BadWord(n) = "nigger"
Presets(nm).Replace(n) = "dude"



'next
'n = 0
'n = n + 1
'Presets(nm).BadWord(n) = ""
'Presets(nm).Replace(n) = ""






End Sub

Private Sub UpdateList()

DoEvents

a = List1.ListIndex
List1.Clear

For i = 1 To NumSwears
    List1.AddItem Swears(i).BadWord
    List1.ItemData(List1.NewIndex) = i
Next i

If a < List1.ListCount Then List1.ListIndex = a: ShowWord
If a >= List1.ListCount Then List1.ListIndex = List1.ListCount - 1: ShowWord

ListSel = List1.ListIndex

End Sub

Private Sub Command1_Click()

'add

n$ = InBox("Enter Word:", "New Word", "")
n$ = Trim(n$)

If n$ = "" Then Exit Sub

For i = 1 To NumSwears
    If n$ = Swears(i).BadWord Then Exit Sub
Next i

NumSwears = NumSwears + 1
ReDim Preserve Swears(0 To NumSwears)

b = NumSwears

Swears(b).BadWord = n$
Swears(b).Flags = 47

ListSel = -1
UpdateList

'List1.AddItem n$
'List1.ListIndex = List1.NewIndex

End Sub

Private Sub Command2_Click()

a = List1.ListIndex
If a = -1 Then Exit Sub

'delete
b = MessBox("Are you sure you want to remove this word?", vbYesNo + vbQuestion, "Delete Word")

If b = vbYes Then

    c = List1.ItemData(a)
    
    RemoveWord c
    
    ListSel = -1

    UpdateList
    
End If



End Sub

Private Sub RemoveWord(c)
    
NumSwears = NumSwears - 1

For i = c To NumSwears

    Swears(i).BadWord = Swears(i + 1).BadWord
    Swears(i).Flags = Swears(i + 1).Flags
    Swears(i).Replacement = Swears(i + 1).Replacement
    
Next i

ReDim Preserve Swears(0 To NumSwears)

End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    NumSwears = 0
    ReDim Swears(0 To 0)
    UpdateList
End If

End Sub

Private Sub Command3_Click()

Unload Me

End Sub

Private Function SetWord() As Boolean

If Chng = 1 Then Exit Function

a = ListSel
If a = -1 Then Exit Function
e = List1.ItemData(a)


'lets do an error check.

h$ = LCase(DeSpace(Text1))
g$ = LCase(DeSpace(Text2))

If Text2 = "" And Check(7).Value = 1 Then
    SetWord = True
    MessBox "You must specify a REPLACE text!", vbOKOnly + vbCritical, "Error with BAD WORD", True, 5
    Exit Function
End If

If h$ = "" And Check(7).Value = 1 Then
    SetWord = True
    MessBox "You may only use Alphabetic characters in the BAD WORD text!", vbOKOnly + vbCritical, "Error with BAD WORD", True, 5
    Exit Function
End If

If g$ = "" And Check(7).Value = 1 Then
    SetWord = True
    MessBox "You may only use Alphabetic characters in the REPLACE text!", vbOKOnly + vbCritical, "Error with BAD WORD", True, 5
    Exit Function
End If

If InStr(1, h$, g$) And Check(7).Value = 1 Then
    SetWord = True
    MessBox "The BAD WORD text cannot be in the REPLACE text!", vbOKOnly + vbCritical, "Error with BAD WORD", True, 5
    Exit Function
End If

If InStr(1, g$, h$) And Check(7).Value = 1 Then
    SetWord = True
    MessBox "The REPLACE text cannot be in the BAD WORD text!", vbOKOnly + vbCritical, "Error with BAD WORD", True, 5
    Exit Function
End If


b = 0
For i = 0 To Check.Count - 1
    If Check(i).Value = 1 Then b = b + (2 ^ i)
Next i

Swears(e).Flags = b
Swears(e).Replacement = Text2

If Text1 <> Swears(e).BadWord Then
    Swears(e).BadWord = Text1
    
    ListSel = -1
    UpdateList
End If

End Function

Function DeSpace(OrigText As String) As String
If DebugMode Then LastCalled = "DeSpace"

'De-Leets a string.


For i = 1 To Len(OrigText)
    a$ = Mid(OrigText, i, 1)
    b = Asc(a$)
    If (b >= Asc("A") And b <= Asc("Z")) Or (b >= Asc("a") And b <= Asc("z")) Then
        c$ = c$ + a$
    End If
Next i

DeSpace = c$

End Function

Private Sub Command4_Click()

SetWord

PackageSwears
Unload Me


End Sub

Private Sub Command5_Click()
a = Combo1.ListIndex + 1

If a <= 0 Then Exit Sub
'List1.Clear

'NumSwears = Presets(a).NumPreset

For i = 1 To Presets(a).NumPreset
        
    us = 1
    For j = 1 To NumSwears
        If LCase(Presets(a).BadWord(i)) = LCase(Swears(j).BadWord) Then us = 0
    Next j
    
    If us = 1 Then 'add
        NumSwears = NumSwears + 1
        ReDim Preserve Swears(0 To NumSwears)
        n = NumSwears
    
        Swears(n).BadWord = Presets(a).BadWord(i)
        Swears(n).Replacement = Presets(a).Replace(i)
        Swears(n).Flags = 2 ^ 0 + 2 ^ 1 + 2 ^ 2 + 2 ^ 3 + 2 ^ 4 + 2 ^ 5 + 2 ^ 7
    
    End If
Next i

ListSel = -1
UpdateList

End Sub

Private Sub Form_Load()
UpdateList
Check(4).Enabled = DllEnabled
Check(7).Enabled = DllEnabled

Msg$ = "Some words commonly accepted as unacceptable " + vbCrLf + "in normal verbal conversation " + vbCrLf + "(e.g. ""shit"", ""crap"", ""damn"", and even ""fuck"") " + vbCrLf + "are used extensively on the internet.  " + vbCrLf + vbCrLf + "Censoring these words is not recommended, " + vbCrLf + "as it can prevent normal gaming conversation, " + vbCrLf + "as can censoring extremely short words " + vbCrLf + "like ""tit"" (""What it does..."", ""Set it's..."", etc.)."

a4 = Val(GetSetting("Server Assistant", "Settings", "FirstReplace", "0"))

If a4 = 0 Then
    a5 = MessBox(Msg$, vbOKOnly, "Notice")
End If

SaveSetting "Server Assistant", "Settings", "FirstReplace", "1"

LoadPreset

For i = 1 To 2
    Combo1.AddItem Presets(i).Name


Next i

Combo1.ListIndex = 0

End Sub

Sub ShowWord()

a = List1.ListIndex
If a = -1 Then Exit Sub

e = List1.ItemData(a)

Text1 = Swears(e).BadWord
Text2 = Swears(e).Replacement

For i = 0 To Check.Count - 1
    If CheckBit2(Swears(e).Flags, i) = True Then Check(i).Value = 1
    If CheckBit2(Swears(e).Flags, i) = False Then Check(i).Value = 0
Next i

End Sub

Private Sub List1_Click()

If NoClick Then Exit Sub

bg = SetWord

If bg = True Then
    
    'error. return to current selection
    NoClick = True
    List1.ListIndex = ListSel

    DoEvents
    
    NoClick = False

Else
    ListSel = List1.ListIndex
    ShowWord
End If

End Sub
