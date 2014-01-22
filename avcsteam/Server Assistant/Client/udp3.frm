VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form3 
   Caption         =   "Commands Editor"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   Icon            =   "udp3.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   9000
   Begin VB.CommandButton Command4 
      Caption         =   "Quick Killer"
      Height          =   495
      Left            =   8040
      TabIndex        =   23
      Top             =   1800
      Width           =   615
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2055
      Left            =   60
      TabIndex        =   20
      Top             =   240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   3625
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Load Local"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6840
      TabIndex        =   14
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Save Locally"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6840
      TabIndex        =   13
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Replace"
      Height          =   255
      Left            =   5640
      TabIndex        =   18
      Top             =   2040
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   5685
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Edit Buttons"
      Height          =   255
      Left            =   5640
      TabIndex        =   12
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Choose Font"
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   3195
      Left            =   60
      TabIndex        =   17
      Top             =   2340
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5636
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"udp3.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Send..."
      Height          =   495
      Left            =   3240
      TabIndex        =   15
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Search"
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog ComDlg1 
      Left            =   8160
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Script Settings"
      Height          =   1575
      Left            =   4560
      TabIndex        =   4
      Top             =   180
      Width           =   4095
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   660
         Sorted          =   -1  'True
         TabIndex        =   22
         Text            =   "No Group"
         Top             =   900
         Width           =   3375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "MUST have this many parameters"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Text            =   "0"
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Group"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label4 
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1260
         Width           =   3915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Number of required parameters"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Rename"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7560
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "udp3.frx":038D
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "udp3.frx":07DF
            Key             =   "file"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "udp3.frx":0C31
            Key             =   "open"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1440
      TabIndex        =   7
      Top             =   60
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Script List"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   690
   End
End
Attribute VB_Name = "Form3"
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

Dim DisableUpdate As Integer
Dim CurrIndex As Integer
Dim SetIndex As Integer
Dim KeyBuffer As String
Dim CommandData() As String
Dim GroupNames() As String

Private Sub LoadCommandData()

a$ = App.Path + "\commands.txt"

ReDim CommandData(1 To 3, 0 To 0)

If CheckForFile(a$) Then

    h = FreeFile
    Open a$ For Input As h
        
        Do While Not EOF(h)
            n = UBound(CommandData, 2)
            
            Line Input #h, b$
            b$ = Trim(b$)
            If b$ = "" Then GoTo nxt
            ReDim Preserve CommandData(1 To 3, 0 To n + 1)
            CommandData(1, n + 1) = b$
            
            Line Input #h, b$
            b$ = Trim(b$)
            If b$ = "" Then GoTo nxt
            CommandData(2, n + 1) = ReplaceString(b$, "\n", vbCrLf)
            
            Line Input #h, b$
            b$ = Trim(b$)
            If b$ = "" Then GoTo nxt
            CommandData(3, n + 1) = b$
            
nxt:
        Loop
    
    Close h
    
End If





End Sub


Private Sub ReList(Optional NextX As Integer)
'    SetIndex = -1
'    List1.Clear
'    For I = 1 To NumCommands
'        List1.AddItem Commands(I).Name
'        a = List1.NewIndex
'        List1.ItemData(a) = I
'        If I = CurrIndex Then
'            SetIndex = a
'        Else
'            If SetIndex > -1 Then
'                If a <= SetIndex Then SetIndex = SetIndex + 1
'            End If
'        End If
'
'    Next I

    Label3 = "Total Scripts: " + Ts(NumCommands)


    'do the tree view
    
    DisableUpdate = 1
    
    a = GetSelTag
    'ab$ = GetSelTagName
        
    If NextX <> 0 Then a = NextX  'ab$ <> Commands(a).Name And a <> 0 Then 'deleted
    '    a = GetSelTagAfter
    'End If
    
    TreeView1.Nodes.Clear
    
    'add groups
    
    ReDim GroupNames(0 To 0)
    
    For i = 1 To NumCommands
        
        If Trim(Commands(i).Group) = "" Then Commands(i).Group = "No Group"
        ad = 1
        n = UBound(GroupNames)
        For j = 0 To n
            If LCase(GroupNames(j)) = LCase(Commands(i).Group) Then ad = 0
        Next j
        
        If ad = 1 Then
            n = n + 1
            ReDim Preserve GroupNames(0 To n)
            GroupNames(n) = Commands(i).Group
        End If
    Next i

    'add groups
    
    bg$ = Combo1.Text
    Combo1.Clear
    
    n = UBound(GroupNames)
    
    Dim mNode As Node
    Dim mNode2 As Node
    Dim mNode3 As Node
    
    For i = 1 To n
        
        Set mNode = TreeView1.Nodes.Add(, , , , 1)
        mNode.Text = GroupNames(i)
        mNode.ExpandedImage = 3
        mNode.Tag = "-1"
        mNode.Sorted = True
                
        For j = 1 To NumCommands
            If LCase(Commands(j).Group) = LCase(GroupNames(i)) Then
                Set mNode2 = TreeView1.Nodes.Add(mNode, tvwChild, , , 2)
                mNode2.Text = Commands(j).Name
                mNode2.Tag = Ts(j)
                mNode2.Sorted = True
                If j = a Then Set mNode3 = mNode2
            End If
        Next j
        
        Combo1.AddItem GroupNames(i)
        
    Next i

    DisableUpdate = 0

    On Error Resume Next

    TreeView1.SelectedItem = mNode3
    
    CurrIndex = 0
    Combo1.Text = bg$
    
    TreeView1_Click
    


    TreeView1.Refresh

End Sub

Function GetSelTag() As Integer

    For i = 1 To TreeView1.Nodes.Count
        Set bbq = TreeView1.Nodes.Item(i).Child
        
        For j = 1 To TreeView1.Nodes.Item(i).Children
            
            If bbq.Selected = True Then GetSelTag = Val(bbq.Tag): Exit Function
            Set bbq = bbq.Next
        Next j
    Next i


End Function

Function GetSelTagName() As String

    For i = 1 To TreeView1.Nodes.Count
        Set bbq = TreeView1.Nodes.Item(i).Child
        
        For j = 1 To TreeView1.Nodes.Item(i).Children
            
            If bbq.Selected = True Then GetSelTagName = Val(bbq.Text): Exit Function
            Set bbq = bbq.Next
        Next j
    Next i


End Function

Function GetSelTagAfter() As Integer

    For i = 1 To TreeView1.Nodes.Count
        Set bbq = TreeView1.Nodes.Item(i).Child
        
        For j = 1 To TreeView1.Nodes.Item(i).Children
            
            If bbq.Selected = True Then
                Set bbq = bbq.Next
                On Error Resume Next
                GetSelTagAfter = Val(bbq.Tag): Exit Function
            End If
        Next j
    Next i


End Function

Private Sub SaveExec()
    'Saves the stuff in the textbox.
    If DisableUpdate = 0 Then
        If CurrIndex > 0 Then
            If Commands(CurrIndex).Exec <> Text1.Text Then Commands(CurrIndex).Changed = True
            Commands(CurrIndex).Exec = Text1.Text
            Commands(CurrIndex).NumParams = Val(Text2)
            Commands(CurrIndex).MustHave = Check1
            If Commands(CurrIndex).Group <> Combo1.Text Then Commands(CurrIndex).Group = Combo1.Text: chg = 1
        End If
    End If
    
    If chg = 1 Then
        Commands(CurrIndex).Changed = True
        ReList
    End If
End Sub

Sub RePaginate(SelStart, Txt As String)

'organizes the scripts in accordance to the contents
Dim Brck As Integer
Dim String1 As String
Dim String2 As String
Dim BrackLevel As String

'Const SB_THUMBPOSITION = 4
'Const SB_THUMBTRACK = 5

Const SIF_RANGE = &H1
Const SIF_PAGE = &H2
Const SIF_POS = &H4
Const SIF_DISABLENOSCROLL = &H8
Const SIF_TRACKPOS = &H10
Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
Const WM_VSCROLL = &H115
Const SB_ENDSCROLL = 8


Dim ScrollBar As SCROLLINFO

ScrollBar.cbSize = Len(ScrollBar)
ScrollBar.fMask = SIF_ALL

bc = GetScrollInfo(Text1.hwnd, SB_VERT, ScrollBar)

'bc = GetScrollPos(Text1.hwnd, SB_THUMBTRACK)

Commands(CurrIndex).Exec = Txt
String1 = Txt



startpos = SelStart
e = 0
fff = 0
Brck = 0
String2 = ""

Do
    e1 = InStrQuote(e + 1, String1, vbCrLf)
    e2 = InStrQuote(e + 1, String1, "{")
    e4 = InStrQuote(e + 1, String1, "}")
    
    flg = 0
    e = e1
    If e = 0 Then e = 100000000
    If e2 < e And e2 > 0 Then flg = 1: e = e2
    If e4 < e And e4 > 0 Then e = e4: flg = 2
    
    If flg = 1 Then Brck = Brck + 1
    If flg = 2 Then
        Brck = Brck - 1
        BrackLevel = ""
        If Brck > 0 Then BrackLevel = Space(Brck * 4)
    End If
    
    If flg = 0 Then 'no bracket, rather... an enter
        'get the line
        ln$ = Mid(String1, fff + 1, e - fff - 1)
        ll1 = Len(ln$)
        ln$ = Trim(ln$)
        lcut = ll1 - Len(ln$)
        
        'startpos = startpos + Len(BrackLevel) - lcut
        
        If startpos >= fff And startpos <= e Then
            endpos = Len(String2) + (startpos - fff) + Len(BrackLevel)
        End If
        
        'add the brack
        ln$ = BrackLevel + ln$
        'add this line
        String2 = String2 + ln$ + vbCrLf
        
        BrackLevel = ""
        If Brck > 0 Then BrackLevel = Space(Brck * 4)
            
        fff = e + 1
    End If
    
Loop Until e = 100000000

If Len(String2) > 2 Then String2 = Left(String2, Len(String2) - 2)


'Text1.Text = String2

Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Text1.SelText = String2

Text1.SelStart = endpos

ScrollBar.nPos = 75

vn = ScrollBar.nPos
'Do
'    bc = GetScrollInfo(Text1.hwnd, SB_VERT, ScrollBar)
'    If ScrollBar.nPos < vn Then bc3 = SendMessage(Text1.hwnd, WM_VSCROLL, 1, "")'

'Loop Until (bc = False Or bc3 = False) Or ScrollBar.nPos >= vn

'Debug.Print bc, bc2, bc3, ScrollBar.nPos, ScrollBar.nMax, ScrollBar.nMin, ScrollBar.cbSize, ScrollBar.nPage, ScrollBar.nTrackPos, ScrollBar.fMask


'bc2 = SetScrollPos(Text1.hwnd, SB_THUMBPOSITION, bc, True)
'Debug.Print bc, bc2



End Sub


Private Sub Check1_Click()
    SaveExec
End Sub

Private Sub Combo1_LostFocus()
'SaveExec
End Sub

Private Sub Command1_Click()
    a$ = InBox("Please specify the new script name:", "Add New Script", "")
    a$ = Trim(a$)
    If a$ = "" Then Exit Sub
    If InStr(1, a$, " ") Then MessBox "No spaces allowed!", vbCritical, "Error": Exit Sub
    
    grp$ = InBox("Please specify the group:", "Add New Script", "No Group")
       
    For i = 1 To NumCommands
        If LCase(a$) = LCase(Commands(i).Name) Then
            MessBox "Can't have duplicate scripts!"
            Exit Sub
        End If
    Next i
    
    DisableUpdate = 1
    NumCommands = NumCommands + 1
    ReDim Preserve Commands(0 To NumCommands)
    Commands(NumCommands).Name = a$
    Commands(NumCommands).Exec = ""
    Commands(NumCommands).NumParams = 0
    Commands(NumCommands).ScriptName = ""
    Commands(NumCommands).Group = grp$
    Commands(NumCommands).NumButtons = 0
    
    'create an ID.
    
    Do
        Randomize
        newid = Int(Rnd * 30000) + 1
        usd = 0
        For i = 1 To NumCommands
            If newid = Commands(i).ScriptID Then usd = 1: Exit For
        Next i
    Loop Until usd = 0
    
    Commands(NumCommands).ScriptID = newid
    ReDim Commands(NumCommands).Buttons(0 To 0)
    ReList
    'a = List1.NewIndex
    'List1.ListIndex = a
    DoEvents
    DisableUpdate = 0
End Sub

Private Sub Command12_Click()

'PackageScripts

MDIForm1.PopupMenu MDIForm1.mnuSend

End Sub

Private Sub Command13_Click()
    f = List1.ListIndex
    If List1.ListCount = 0 Then Exit Sub
    If f = -1 Then Exit Sub
    
    PackageOneScripts CurrIndex
    
    
    
End Sub

Private Sub Command14_Click()

a$ = InBox("Replace all occurences of:", "Replace...", Text1.SelText)

If a$ = "" Then Exit Sub
b$ = InBox("With...", "With...", Text1.SelText)


c$ = Text1.Text

d$ = ReplaceString(c$, a$, b$)

Text1.Text = d$

e = MessBox("Keep change?", vbYesNo, "Replace")

If e = vbNo Then Text1.Text = c$

End Sub

Private Sub Command2_Click()


a = Val(TreeView1.SelectedItem.Tag)
If a = 0 Then Exit Sub

If a = -1 Then

    'renaming a group
    
    b$ = TreeView1.SelectedItem.Text
    
    c = MessBox("Are you sure you want to delete the group " + b$ + "?" + vbCrLf + "All scripts in this group will be moved to No Group.", vbQuestion + vbYesNo, "Delete Group")

    If c = vbYes Then
    
        For i = 1 To NumCommands
            
            If LCase(Commands(i).Group) = LCase(b$) Then
                Commands(i).Group = "No Group"
            End If
        Next i
    
        ReList
    End If
Else

    If CurrIndex > 0 Then
        
        b$ = Commands(CurrIndex).Name
        c = MessBox("Are you sure you want to delete " + Chr(34) + b$ + Chr(34) + "?", vbYesNo, "Delete Script")
        
        If c = vbYes Then
            If CurrIndex < NumCommands Then
                For i = CurrIndex To NumCommands - 1
                    Commands(i).Exec = Commands(i + 1).Exec
                    Commands(i).Name = Commands(i + 1).Name
                    Commands(i).NumParams = Commands(i + 1).NumParams
                    Commands(i).MustHave = Commands(i + 1).MustHave
                    Commands(i).ScriptName = Commands(i + 1).ScriptName
                    Commands(i).AutoMakeVars = Commands(i + 1).AutoMakeVars
                    Commands(i).Group = Commands(i + 1).Group
                    Commands(i).LogExec = Commands(i + 1).LogExec
                    Commands(i).Unused1 = Commands(i + 1).Unused1
                    Commands(i).unused2 = Commands(i + 1).unused2
                    Commands(i).unused3 = Commands(i + 1).unused3
                    Commands(i).ScriptID = Commands(i + 1).ScriptID
                    Commands(i).Unused5 = Commands(i + 1).Unused5
                    
                    
                    Commands(i).NumButtons = Commands(i + 1).NumButtons
                    Commands(i).Changed = Commands(i + 1).Changed
                    ReDim Commands(i).Buttons(0 To Commands(i + 1).NumButtons)
                    For j = 1 To Commands(i + 1).NumButtons
                        Commands(i).Buttons(j) = Commands(i + 1).Buttons(j)
                    Next j
                Next i
            End If
            
            NumCommands = NumCommands - 1
            ReDim Preserve Commands(0 To NumCommands)
            DisableUpdate = 1
            
            Text1.Text = ""
            Combo1.Text = ""
            Check1 = 0
            Label4 = ""
            
            nnx = GetSelTagAfter
            
            
            ReList CInt(nnx)
'            If NumCommands > 0 Then
'                If a <= NumCommands - 1 Then
'                    List1.ListIndex = a
'                Else
'                    List1.ListIndex = a - 1
'                End If
'            End If
            DoEvents
            DisableUpdate = 0
        End If
    End If
End If

End Sub

Private Sub Command3_Click()
    f = Val(TreeView1.SelectedItem.Tag)
    If f = 0 Then Exit Sub
    
    If f = -1 Then
    
        'renaming a group
        
        b$ = TreeView1.SelectedItem.Text
        
        c$ = InBox("Enter New Name for Group " + b$ + ":", "Rename Group", b$)
            
    
        If c$ <> "" And c$ <> b$ Then
        
            For i = 1 To NumCommands
                
                If LCase(Commands(i).Group) = LCase(b$) Then
                    Commands(i).Group = c$
                End If
            Next i
        
            ReList
        End If
    Else
            
        
        b$ = Commands(CurrIndex).Name
        a$ = InBox("Please enter a new name for " + Chr(34) + b$ + Chr(34) + ":", "Rename Script", b$)
        a$ = Trim(a$)
        If a$ = "" Then Exit Sub
        If InStr(1, a$, " ") Then MessBox "No spaces allowed!", vbCritical, "Error": Exit Sub
            
        For i = 1 To NumCommands
            If LCase(a$) = LCase(Commands(i).Name) Then
                MessBox "Can't have duplicate scripts!"
                Exit Sub
            End If
        Next i
            
        DisableUpdate = 1
        Commands(CurrIndex).Name = a$
        Commands(CurrIndex).Changed = True
        ReList
        'List1.ListIndex = SetIndex
        DoEvents
        DisableUpdate = 0
    
    End If

End Sub

Private Sub Command4_Click()
frmQuickKiller.Show
Unload Me

End Sub

Private Sub Command5_Click()

ComDlg1.FontName = Text1.Font
'ComDlg1.Color = Text1.ForeColor

ComDlg1.Flags = cdlCFEffects Or cdlCFBoth

ComDlg1.ShowFont

Text1.Font = ComDlg1.FontName
'Text1.ForeColor = ComDlg1.Color

'Settings.FontName = Text1.FontName
'Settings.FontBold = Text1.FontBold
'Settings.FontItalic = Text1.FontItalic
'Settings.FontSize = Text1.FontSize
'Settings.FontColor = Text1.ForeColor
'Settings.FontStrikethru = Text1.FontStrikethru
'Settings.FontUnderline = Text1.FontUnderline
End Sub

Private Sub Command6_Click()
'Form8.Show

a$ = InBox("Enter Search Criteria:", "Search Scripts")
        
b = CurrIndex + 1
If b = 0 Then b = 1

For i = b To NumCommands
    If InStr(1, LCase(Commands(i).Exec), a$) Then j = i: Exit For
Next i

If j = 0 Then
    MessBox "Nothing found!"
Else
    

    
    For i = 1 To TreeView1.Nodes.Count
        
        If Val(TreeView1.Nodes.Item(i).Tag) = j Then
                
            TreeView1.SelectedItem = TreeView1.Nodes.Item(i)
            TreeView1_Click
            
            Exit For
        End If
        
        
    Next i
    
    


End If





End Sub

Private Sub Command8_Click()
a = MessBox("Are you sure?", vbYesNo + vbQuestion, "Close without sending?")
Unload Me

End Sub

Private Sub Command9_Click()

a = Val(TreeView1.SelectedItem.Tag)
If a <= 0 Then Exit Sub

EditedButton = CurrIndex

frmButEditor.Show

End Sub

Private Sub Form_Load()
LoadCommandData

ReList
CurrIndex = 0
'Text1.FontName = Settings.FontName
'Text1.FontBold = Settings.FontBold
'Text1.FontItalic = Settings.FontItalic
'Text1.FontSize = Settings.FontSize
'Text1.FontStrikethru = Settings.FontStrikethru
'Text1.FontUnderline = Settings.FontUnderline
'Text1.ForeColor = Settings.FontColor
'Me.Width = MDIForm1.Width - 800
'Me.Height = MDIForm1.Height - 1600
'Me.Move 0, 0

On Error Resume Next
nm$ = Me.Name
winash = Val(GetSetting("Server Assistant Client", "Window", nm$ + "winash", -1))
If winash <> -1 Then Me.Show
winmd = Val(GetSetting("Server Assistant Client", "Window", nm$ + "winmd", 3))
If winmd <> 3 Then Me.WindowState = winmd
If Me.WindowState = 0 Then
    winh = GetSetting("Server Assistant Client", "Window", nm$ + "winh", -1)
    wint = GetSetting("Server Assistant Client", "Window", nm$ + "wint", -1)
    winl = GetSetting("Server Assistant Client", "Window", nm$ + "winl", -1)
    winw = GetSetting("Server Assistant Client", "Window", nm$ + "winw", -1)
    
    If winh <> -1 Then Me.Height = winh
    If wint <> -1 Then Me.Top = wint
    If winl <> -1 Then Me.Left = winl
    If winw <> -1 Then Me.Width = winw
End If


End Sub

Public Sub SendMenu(Index As Integer)

If Index = 0 Then
    
'    For I = 1 To NumCommands
'
'        If Commands(I).Changed = True Then
'
'            PackageOneScripts CInt(I)
'
'            Commands(I).Changed = False
'
'        End If
'    Next I

    PackageScripts True

ElseIf Index = 1 Then
    PackageScripts

ElseIf Index = 2 Then
    PackageScripts
    Unload Me
End If




End Sub

Private Sub Form_Resize()

If Me.WindowState = 1 Then Exit Sub

If Me.WindowState <> 2 Then

    If Me.Width < 8805 Then Me.Width = 8805
    If Me.Height < 6330 Then Me.Height = 6330
End If

Text1.Height = Me.Height - Text1.Top - 405 - StatusBar1.Height
Text1.Width = Me.Width - Text1.Left - 120



End Sub

Private Sub Form_Unload(Cancel As Integer)
Settings.CommandWidth = Me.Width
Settings.CommandHeight = Me.Height
SaveCommands


On Error Resume Next
nm$ = Me.Name
SaveSetting "Server Assistant Client", "Window", nm$ + "winmd", Me.WindowState
SaveSetting "Server Assistant Client", "Window", nm$ + "winh", Me.Height
SaveSetting "Server Assistant Client", "Window", nm$ + "wint", Me.Top
SaveSetting "Server Assistant Client", "Window", nm$ + "winl", Me.Left
SaveSetting "Server Assistant Client", "Window", nm$ + "winw", Me.Width



End Sub

Private Sub List1_Click()

    If CurrIndex <> 0 Then SaveExec
    a = List1.ListIndex
    CurrIndex = List1.ItemData(a)
    DisableUpdate = 1
    
    Text1.Text = Commands(CurrIndex).Exec
    Label4 = Ts(Commands(CurrIndex).NumButtons) + " buttons, " + Commands(CurrIndex).ScriptName
    Text2 = Trim(str(Commands(CurrIndex).NumParams))
    Check1 = Commands(CurrIndex).MustHave
        
    DoEvents
    DisableUpdate = 0
    
End Sub

Private Sub Text1_Change()
    If DisableUpdate = 0 Then SaveExec
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
       
    a3 = Text1.SelStart + 2
    
    If Text1.SelStart > 0 Then a1$ = Left(Text1.Text, Text1.SelStart)
    If Text1.SelStart < Len(Text1.Text) Then a2$ = Right(Text1.Text, Len(Text1.Text) - Text1.SelStart)
       
    
    RePaginate a3, a1$ + vbCrLf + a2$
    KeyAscii = 0
    
Else



End If


End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)

    On Error Resume Next


    a$ = Chr(KeyAscii)
'    If KeyAscii = 8 Then
'
'    Else
'        KeyBuffer = KeyBuffer + a$
'    End If
    
    
    b = Text1.SelStart
    c$ = Text1.Text
    If b = 0 Then Exit Sub
    
    e = InStrRev(c$, " ", b)
    e7 = InStrRev(c$, vbCrLf, b)
    If e7 > 0 And e7 > e Then e = e7 + 1
    
    
    e1 = InStrRev(c$, "(", b)
    e4 = b
    n = BracketCount(b, c$)
    
tygn:
    e3 = InStrRev(c$, ")", e4)
    
    nn = nn + 1
   
    
        
    If n > 0 Then
    
    
        If e3 > e1 And e3 > 0 Then
            e4 = e1
            e1 = InStrRev(c$, "(", e1 - 1)
            If nn < 100 Then GoTo tygn
        End If
    End If
    
    If e1 > 0 And (e1 > e Or n > 0) Then
        e = e1
        Mde = 1
        e2 = InStrRev(c$, " ", e)
        e7 = InStrRev(c$, vbCrLf, e)
        If e7 > 0 And e7 > e2 Then e2 = e7 + 1
    End If
    
    
    KeyBuffer = Mid(c$, e + 1, b - e)
       
    If Mde = 1 Then
        
        d$ = Mid(c$, e2 + 1, e - e2 - 1)
    
    End If
    
    e8 = InStrRev(d$, "(")
    If e8 > 0 Then
        d$ = Right(d$, Len(d$) - e8)
    End If
    

    
    
    'function name is d$
    'the params are in keyBuffer
    
    'search:
    n = UBound(CommandData, 2)
    
    For i = 1 To n
        If LCase(d$) = CommandData(1, i) Then
            j = i
            Exit For
        End If
    Next i
    
    If j > 0 Then
    
        StatusBar1.SimpleText = CommandData(3, j)
    
    Else
    
        StatusBar1.SimpleText = ""
    
        
    End If
    
    Debug.Print d$; " - " + KeyBuffer


End Sub

Private Sub Text2_Change()
    If DisableUpdate = 0 Then SaveExec
End Sub


Private Sub TreeView1_Click()
    
    
    If DisableUpdate = 1 Then Exit Sub

    If CurrIndex <> 0 Then SaveExec
    On Error GoTo errcur
    
    
    a = Val(TreeView1.SelectedItem.Tag)
    If a = 0 Then Exit Sub
    
    If a = -1 Then
    
        'clicked a group
        
        DisableUpdate = 1
        CurrIndex = 0
        Text1.Text = ""
        Label4 = "Group"
        Text2 = ""
        Check1 = 0
        Combo1.Text = ""
        DoEvents
        DisableUpdate = 0
    
    
    Else
        CurrIndex = a
        DisableUpdate = 1
        
        Text1.Text = Commands(CurrIndex).Exec
        Label4 = Ts(Commands(CurrIndex).NumButtons) + " buttons, " + Commands(CurrIndex).ScriptName
        Text2 = Trim(str(Commands(CurrIndex).NumParams))
        Check1 = Commands(CurrIndex).MustHave
        Combo1.Text = Commands(CurrIndex).Group
        DoEvents
        DisableUpdate = 0
        
        Me.Caption = Commands(CurrIndex).ScriptID
        
        
    End If
    
errcur:
       

End Sub

Private Sub TreeView1_KeyUp(KeyCode As Integer, Shift As Integer)

    If DisableUpdate = 1 Then Exit Sub

    If CurrIndex <> 0 Then SaveExec
    On Error GoTo errcur
    a = Val(TreeView1.SelectedItem.Tag)
    If a = 0 Then Exit Sub
    
    If a = -1 Then
    
        'clicked a group
        
        DisableUpdate = 1
        CurrIndex = 0
        Text1.Text = ""
        Label4 = "Group"
        Text2 = ""
        Check1 = 0
        Combo1.Text = ""
        DoEvents
        DisableUpdate = 0
    
    
    Else
        CurrIndex = a
        DisableUpdate = 1
        
        Text1.Text = Commands(CurrIndex).Exec
        Label4 = Ts(Commands(CurrIndex).NumButtons) + " buttons, " + Commands(CurrIndex).ScriptName
        Text2 = Trim(str(Commands(CurrIndex).NumParams))
        Check1 = Commands(CurrIndex).MustHave
        Combo1.Text = Commands(CurrIndex).Group
        DoEvents
        DisableUpdate = 0
        
        
    End If
    
errcur:
End Sub
