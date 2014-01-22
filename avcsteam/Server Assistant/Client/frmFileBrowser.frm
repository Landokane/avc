VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFileBrowser 
   Caption         =   "Browser"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10320
   Icon            =   "frmFileBrowser.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   10320
   Begin VB.Frame Frame1 
      Caption         =   "List 2"
      Height          =   5235
      Index           =   1
      Left            =   3840
      TabIndex        =   17
      Top             =   0
      Width           =   3615
      Begin VB.ComboBox Combo 
         Height          =   315
         Index           =   1
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   240
         Width           =   1995
      End
      Begin VB.Frame DirFrame 
         Caption         =   "Directories"
         Height          =   1215
         Index           =   1
         Left            =   2100
         TabIndex        =   28
         Top             =   1140
         Width           =   1455
         Begin VB.CommandButton CmdUp 
            Height          =   555
            Index           =   1
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   31
            Tag             =   "Up"
            ToolTipText     =   "Up 1 Dir"
            Top             =   240
            Width           =   555
         End
         Begin VB.CommandButton CmdCreate 
            Caption         =   "Create"
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   30
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton CmdChange 
            Caption         =   "ChgDir"
            Height          =   555
            Index           =   1
            Left            =   660
            TabIndex        =   29
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame DisFrame 
         Caption         =   "Display"
         Height          =   855
         Index           =   1
         Left            =   2100
         TabIndex        =   25
         Top             =   240
         Width           =   1455
         Begin VB.OptionButton OptLocal 
            Caption         =   "Local"
            Height          =   255
            Index           =   1
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton OptRemote 
            Caption         =   "Remote"
            Height          =   255
            Index           =   1
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   540
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame OpFrame 
         Caption         =   "Operations"
         Height          =   2775
         Index           =   1
         Left            =   2100
         TabIndex        =   18
         Top             =   2400
         Width           =   1455
         Begin VB.CommandButton CmdEditComplete 
            Caption         =   "Edit Complete!"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   35
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton CmdDelete 
            Caption         =   "Delete"
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   24
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton CmdCopy 
            Caption         =   "Copy"
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   23
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton CmdMove 
            Caption         =   "Move"
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   22
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton CmdRename 
            Caption         =   "Rename"
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   21
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton CmdEdit 
            Caption         =   "Edit"
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   20
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton CmdRefresh 
            Caption         =   "Refresh"
            Height          =   315
            Index           =   1
            Left            =   60
            TabIndex        =   19
            Top             =   2040
            Width           =   1335
         End
      End
      Begin MSComctlLib.ListView ListView 
         Height          =   4575
         Index           =   1
         Left            =   60
         TabIndex        =   33
         Top             =   600
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   8070
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Filename"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date/Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "List 1"
      Height          =   5235
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.Frame OpFrame 
         Caption         =   "Operations"
         Height          =   2775
         Index           =   0
         Left            =   2100
         TabIndex        =   9
         Top             =   2400
         Width           =   1455
         Begin VB.CommandButton CmdEditComplete 
            Caption         =   "Edit Complete!"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   34
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton CmdRefresh 
            Caption         =   "Refresh"
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   16
            Top             =   2040
            Width           =   1335
         End
         Begin VB.CommandButton CmdEdit 
            Caption         =   "Edit"
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   14
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton CmdRename 
            Caption         =   "Rename"
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   13
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton CmdMove 
            Caption         =   "Move"
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   12
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton CmdCopy 
            Caption         =   "Copy"
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   11
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton CmdDelete 
            Caption         =   "Delete"
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   10
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame DisFrame 
         Caption         =   "Display"
         Height          =   855
         Index           =   0
         Left            =   2100
         TabIndex        =   4
         Top             =   240
         Width           =   1455
         Begin VB.OptionButton OptRemote 
            Caption         =   "Remote"
            Height          =   255
            Index           =   0
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   540
            Width           =   1335
         End
         Begin VB.OptionButton OptLocal 
            Caption         =   "Local"
            Height          =   255
            Index           =   0
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame DirFrame 
         Caption         =   "Directories"
         Height          =   1215
         Index           =   0
         Left            =   2100
         TabIndex        =   3
         Top             =   1140
         Width           =   1455
         Begin VB.CommandButton CmdChange 
            Caption         =   "ChgDir"
            Height          =   555
            Index           =   0
            Left            =   660
            TabIndex        =   15
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton CmdCreate 
            Caption         =   "Create"
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   8
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton CmdUp 
            Height          =   555
            Index           =   0
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   7
            Tag             =   "Up"
            ToolTipText     =   "Up 1 Dir"
            Top             =   240
            Width           =   555
         End
      End
      Begin VB.ComboBox Combo 
         Height          =   315
         Index           =   0
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1995
      End
      Begin MSComctlLib.ListView ListView 
         Height          =   4575
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   600
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   8070
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Filename"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date/Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8220
      Top             =   6240
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
            Picture         =   "frmFileBrowser.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":05A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":09F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Height          =   5175
      Left            =   3660
      MousePointer    =   9  'Size W E
      TabIndex        =   36
      Top             =   60
      Width           =   135
   End
End
Attribute VB_Name = "frmFileBrowser"
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

Dim IsLocal(0 To 1) As Boolean
Dim RecentLocal As New Collection
Dim RecentRemote As New Collection
Dim LastIndex(0 To 1) As Integer
Dim Perc As Double


Public Sub RefreshList(Num As Integer)

'Refresh one of the 2 windows... depending on path and islocal status

If IsLocal(Num) Then 'this is a local window
    'add it to recent list
    p$ = DirFullPath(Num)

again1:
    For i = 1 To RecentLocal.Count
        If RecentLocal(i) = p$ Then RecentLocal.Remove i: GoTo again1  'remove this item
    Next i
        
    If RecentLocal.Count > 0 Then RecentLocal.Add p$, , 1
    If RecentLocal.Count = 0 Then RecentLocal.Add p$
    If RecentLocal.Count > 20 Then RecentLocal.Remove 21
    
    'fill combo
    Combo(Num).Clear
    For j = 1 To RecentLocal.Count
        Combo(Num).AddItem RecentLocal(j)
    Next j
    If Combo(Num).ListCount > 0 Then Combo(Num).ListIndex = 0
    
    'fill the happy little list window
    FillArray Num
    DisplayArray Num

Else
    
    p$ = DirFullPath(Num)

again2:
    For i = 1 To RecentRemote.Count
        If RecentRemote(i) = p$ Then RecentRemote.Remove i: GoTo again2  'remove this item
    Next i
        
    If RecentRemote.Count > 0 Then RecentRemote.Add p$, , 1
    If RecentRemote.Count = 0 Then RecentRemote.Add p$
    If RecentRemote.Count > 20 Then RecentRemote.Remove 21
    
    'fill combo
    Combo(Num).Clear
    For j = 1 To RecentRemote.Count
        Combo(Num).AddItem RecentRemote(j)
    Next j
    If Combo(Num).ListCount > 0 Then Combo(Num).ListIndex = 0
    
    DisplayArray Num
End If

End Sub

Public Sub DisplayArray(Num As Integer)

'add this info to the list box

k = ListView(Num).SortKey
ListView(Num).Sorted = False

On Error Resume Next
For i = 1 To NumDirs(Num)
    If Num = 0 Then If Len(DirList0(i).Size) > mxsize Then mxsize = Len(DirList0(i).Size)
    If Num = 1 Then If Len(DirList1(i).Size) > mxsize Then mxsize = Len(DirList1(i).Size)
Next i

With ListView(Num)
    With .ListItems
        .Clear
        
        'start adding
        
        For i = 1 To NumDirs(Num)
            .Add i
            With .Item(i)
                .Tag = i
                If Num = 0 Then
                    If DirList0(i).Type = 0 Then
                        .SmallIcon = 2
                        .Text = DirList0(i).Name
                        .SubItems(2) = Numberize(DirList0(i).Size, CInt(mxsize))
                        .SubItems(1) = Format(DirList0(i).DateTime, "dd/mm/yyyy hh:mm:ss")
                    Else
                        .SmallIcon = 1
                        .Text = " " + DirList0(i).Name
                        .SubItems(1) = " " + Format(DirList0(i).DateTime, "dd/mm/yyyy hh:mm:ss")
                    End If
                Else
                    If DirList1(i).Type = 0 Then
                        .SmallIcon = 2
                        .Text = DirList1(i).Name
                        .SubItems(2) = Numberize(DirList1(i).Size, CInt(mxsize))
                        .SubItems(1) = Format(DirList1(i).DateTime, "dd/mm/yyyy hh:mm:ss")
                    Else
                        .SmallIcon = 1
                        .Text = " " + DirList1(i).Name
                        .SubItems(1) = " " + Format(DirList1(i).DateTime, "dd/mm/yyyy hh:mm:ss")
                    End If
                End If
            End With
        Next i
    End With
End With


ListView(Num).SortKey = k
ListView(Num).Sorted = True


End Sub

Private Sub FillArray(Num As Integer)

'fills the directory array with the specified information

p$ = DirFullPath(Num)
s$ = p$ + "\*.*"

a$ = Dir(s$, vbArchive + vbDirectory + vbHidden + vbReadOnly + vbSystem)
On Error Resume Next

NumDirs(Num) = 0
Do While a$ <> ""
    
    If a$ <> "." And a$ <> ".." Then
        'add this one
        NumDirs(Num) = NumDirs(Num) + 1
        
        If Num = 0 Then ReDim Preserve DirList0(0 To NumDirs(Num))
        If Num = 1 Then ReDim Preserve DirList1(0 To NumDirs(Num))
        
        If Num = 0 Then
            DirList0(NumDirs(Num)).DateTime = FileDateTime(p$ + "\" + a$)
            DirList0(NumDirs(Num)).FullPath = p$ + "\" + a$
            DirList0(NumDirs(Num)).Name = a$
            
            If (GetAttr(p$ + "\" + a$) And vbDirectory) = vbDirectory Then
                DirList0(NumDirs(Num)).Type = 1
            Else
                DirList0(NumDirs(Num)).Type = 0
                DirList0(NumDirs(Num)).Size = Ts(FileLen(p$ + "\" + a$))
            End If
        Else
            DirList1(NumDirs(Num)).DateTime = FileDateTime(p$ + "\" + a$)
            DirList1(NumDirs(Num)).FullPath = p$ + "\" + a$
            DirList1(NumDirs(Num)).Name = a$
            
            If (GetAttr(p$ + "\" + a$) And vbDirectory) = vbDirectory Then
                DirList1(NumDirs(Num)).Type = 1
            Else
                DirList1(NumDirs(Num)).Type = 0
                DirList1(NumDirs(Num)).Size = Ts(FileLen(p$ + "\" + a$))
            End If
        End If
    End If
    a$ = Dir
Loop


End Sub

Public Sub RefreshDir(Num As Integer, p$)

'refresh a dir as needed

If Len(p$) > 0 Then
    If Right(p$, 1) = "\" Then p$ = Left(p$, Len(p$) - 1)
End If

LastRefresh = Num

If IsLocal(Num) Then
    If Dir(p$, vbDirectory) = "" Then MessBox "Local Directory not found!", vbCritical, "Error!": Exit Sub
    
    DirFullPath(Num) = p$
    RefreshList Num
Else
    'send refresh command to server
    SendPacket "F1", p$
    DirFullPath(Num) = p$

End If


End Sub
Private Sub EditFile(b$, Num)

'sends request to edit file

If IsLocal(Num) = False Then
    'set info
    FileBuffer = ""
    FileMode = 1
    FilePath = b$
    FileLocalPath = EditFileTemp
    SendPacket "F8", b$
Else
    ShellExecute MDIForm1.hwnd, "open", b$, vbNullString, vbNullString, SW_SHOW
End If

End Sub


Private Sub EditBSP(Fle$, b$, Num)

'sends request to edit file

If IsLocal(Num) = False Then
    If UCase(Right(Fle$, 3)) = "BSP" Then
        FileLocalPath = Left(Fle$, Len(Fle$) - 4)
        FilePath = b$
        SendPacket "BP", b$
    Else
        MessBox "BSP Editing only works for files of type BSP.", vbOKOnly, "BSP Edit"
    End If
Else
    MessBox "BSP Editing only works for remote files.", vbOKOnly, "BSP Edit"
End If

End Sub

Private Sub RenameFiles(b$, Num, d$)
'rename these files from either local or server

a$ = InBox("Rename file: " + vbCrLf + b$ + vbCrLf + "to:", "Rename File", b$)
If a$ = "" Then Exit Sub

If IsLocal(Num) = False Then
    'set info
    SendPacket "F7", d$ + "\" + b$ + Chr(250) + d$ + "\" + a$
Else
    Name d$ + "\" + b$ As d$ + "\" + a$
    a = Combo(Num).ListIndex
    If a = -1 Then Exit Sub

    RefreshDir CInt(Num), Combo(Index).List(a)
End If

End Sub

Private Sub CmdChange_Click(Index As Integer)
'change to remote dir

a = Combo(Index).ListIndex
If a <> -1 Then b$ = Combo(Index).List(a)

If IsLocal(Index) Then c$ = InBox("Enter local path:", "Change Local Directory", b$)
If IsLocal(Index) = False Then c$ = InBox("Enter remote path:", "Change Remote Directory", b$)

If c$ = "" Then Exit Sub


RefreshDir Index, c$

End Sub

Private Sub CmdCopy_Click(Index As Integer)
'COPIES FILES from EITHER:
'Server to Server - 1
'Server to Client - 2
'Client to Server - 3
'Client to Client (local copy) - 4

'CURRENTLY ONLY MODE 1 DONE!

'see whats selected
If Index = 0 Then othr = 1
If Index = 1 Then othr = 0


If IsLocal(Index) = False Then
    
    'Fill The Array
        NumDirs(Index) = 0
    If Index = 0 Then ReDim Preserve DirList0(0 To 1)
    If Index = 1 Then ReDim Preserve DirList1(0 To 1)
    
    For i = 1 To ListView(Index).ListItems.Count
        With ListView(Index).ListItems.Item(i)
            If .Selected = True Then 'selected!
                NumDirs(Index) = NumDirs(Index) + 1
                            
                'Add this record
                If Index = 0 Then
                    ReDim Preserve DirList0(0 To NumDirs(Index))
                    
                    DirList0(NumDirs(Index)).DateTime = Now
                    DirList0(NumDirs(Index)).FullPath = DirFullPath(Index) + "\" + Trim(.Text)
                    DirList0(NumDirs(Index)).Name = Trim(.Text)
                    DirList0(NumDirs(Index)).Size = Trim(.SubItems(2))
                    If .SmallIcon = 1 Then DirList0(NumDirs(Index)).Type = 1
                    If .SmallIcon = 2 Then DirList0(NumDirs(Index)).Type = 0
                Else
                    ReDim Preserve DirList1(0 To NumDirs(Index))
                    
                    DirList1(NumDirs(Index)).DateTime = Now
                    DirList1(NumDirs(Index)).FullPath = DirFullPath(Index) + "\" + Trim(.Text)
                    DirList1(NumDirs(Index)).Name = Trim(.Text)
                    DirList1(NumDirs(Index)).Size = Trim(.SubItems(2))
                    If .SmallIcon = 1 Then DirList1(NumDirs(Index)).Type = 1
                    If .SmallIcon = 2 Then DirList1(NumDirs(Index)).Type = 0
                End If
            End If
        End With
    Next i

    'Send appropriate commands
    
    If IsLocal(othr) = False Then 'Its Mode 1
        'Send COPY command
        SendPacket "F6", PackageDirList(Index, DirFullPath(othr))
    Else
        ' Copying from SERVER to CLIENT
        
        j = 0
        For i = 1 To ListView(Index).ListItems.Count
            With ListView(Index).ListItems.Item(i)
                If .Selected = True Then j = i: Exit For
            End With
        Next i
        
        If j > 0 Then
            'do shtuff
            If ListView(Index).ListItems.Item(j).SmallIcon = 2 Then 'file
                'The full LOCAL path the file...
                FileLocalPath = DirFullPath(othr) + "\" + Trim(ListView(Index).ListItems.Item(j).Text)
                'The full REMOTE path to the file...
                FilePath = DirFullPath(Index) + "\" + Trim(ListView(Index).ListItems.Item(j).Text)
                
                FileMode = 0
                SendPacket "F8", FilePath
            End If
        End If
        
    
    End If
Else
    If IsLocal(othr) = False Then
        'Window (index) copying to server
        
        For i = 1 To ListView(Index).ListItems.Count
            With ListView(Index).ListItems.Item(i)
                If .Selected = True Then j = i: Exit For
            End With
        Next i
        
        If j > 0 Then
            
            'do shtuff
            
            If ListView(Index).ListItems.Item(j).SmallIcon = 2 Then 'file
                'Copy this file over!
                
                'The full LOCAL path the file...
                b$ = DirFullPath(Index) + "\" + Trim(ListView(Index).ListItems.Item(j).Text)
                'The full REMOTE path to the file...
                c$ = DirFullPath(othr) + "\" + Trim(ListView(Index).ListItems.Item(j).Text)
                
                PackageFileSend b$, c$
                
            End If
        End If
            
    Else
        'Mode 4 - Client to Client


    End If
End If

End Sub


Private Sub CmdDelete_Click(Index As Integer)

'see whats selected

If IsLocal(Index) = False Then

    NumDirs(Index) = 0
    If Index = 0 Then ReDim Preserve DirList0(0 To 1)
    If Index = 1 Then ReDim Preserve DirList1(0 To 1)
    
    For i = 1 To ListView(Index).ListItems.Count
        With ListView(Index).ListItems.Item(i)
            If .Selected = True Then 'selected!
                NumDirs(Index) = NumDirs(Index) + 1
                            
                'Add this record
                If Index = 0 Then
                    ReDim Preserve DirList0(0 To NumDirs(Index))
                    
                    DirList0(NumDirs(Index)).DateTime = Now
                    DirList0(NumDirs(Index)).FullPath = DirFullPath(Index) + "\" + Trim(.Text)
                    DirList0(NumDirs(Index)).Name = Trim(.Text)
                    DirList0(NumDirs(Index)).Size = Trim(.SubItems(2))
                    If .SmallIcon = 1 Then DirList0(NumDirs(Index)).Type = 1
                    If .SmallIcon = 2 Then DirList0(NumDirs(Index)).Type = 0
                Else
                    ReDim Preserve DirList1(0 To NumDirs(Index))
                    
                    DirList1(NumDirs(Index)).DateTime = Now
                    DirList1(NumDirs(Index)).FullPath = DirFullPath(Index) + "\" + Trim(.Text)
                    DirList1(NumDirs(Index)).Name = Trim(.Text)
                    DirList1(NumDirs(Index)).Size = Trim(.SubItems(2))
                    If .SmallIcon = 1 Then DirList1(NumDirs(Index)).Type = 1
                    If .SmallIcon = 2 Then DirList1(NumDirs(Index)).Type = 0
                End If
            End If
        End With
    Next i
    
    'The array is now filled. Send the DELETE command
    
    If NumDirs(Index) > 0 Then
        bd = MessBox("Are you sure you want to delete these files?", vbYesNo + vbQuestion, "Delete Files?")
        If bd = vbNo Then Exit Sub
    End If
    SendPacket "F2", PackageDirList(Index)
Else
    
    
    Dim Result As Long
    Dim FileOp As SHFILEOPSTRUCT
    
    bd = MessBox("Are you sure you want to delete these files?", vbYesNo + vbQuestion, "Delete Files?")
    If bd = vbNo Then Exit Sub

    For i = 1 To ListView(Index).ListItems.Count
        With ListView(Index).ListItems.Item(i)
            If .Selected = True Then 'selected!
                If CheckForFile(DirFullPath(Index) + "\" + Trim(.Text)) Then
                
                    FileOp.wFunc = FO_DELETE
                    FileOp.pFrom = DirFullPath(Index) + "\" + Trim(.Text) + vbNullChar + vbNullChar
                    FileOp.pTo = DirFullPath(Index) + "\" + Trim(.Text) + vbNullChar + vbNullChar
                    FileOp.fFlags = FOF_NOCONFIRMATION Or FOF_SILENT Or FOF_NOCONFIRMMKDIR
                    Result = SHFileOperation(FileOp)
                    DoEvents
                    
                End If
            End If
        End With
    Next i
    
    RefreshDir Index, DirFullPath(Index)
End If

End Sub

Private Sub CmdEdit_Click(Index As Integer)

'Edit this file

For i = 1 To ListView(Index).ListItems.Count
    With ListView(Index).ListItems.Item(i)
        If .Selected = True Then j = i: Exit For
    End With
Next i

If j > 0 Then
    
    'do shtuff
    
    If ListView(Index).ListItems.Item(j).SmallIcon = 2 Then 'file
        'edit
        b$ = DirFullPath(Index) + "\" + Trim(ListView(Index).ListItems.Item(j).Text)
                
        Fle$ = Trim(ListView(Index).ListItems.Item(j).Text)
        
        If IsLocal(Index) = False Then
            If UCase(Right(Fle$, 3)) = "BSP" Then
                FileLocalPath = Left(Fle$, Len(Fle$) - 4)
                FilePath = b$
                SendPacket "BP", b$
            Else
                EditFile b$, Index
            End If
        Else
            EditFile b$, Index
        End If
    End If
End If

End Sub

Private Sub CmdEdit_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    
    MessBox "To edit BSP files in this version," + vbCrLf + "select the file and then left-click EDIT." + vbCrLf + vbCrLf + "Right-clicking is no longer needed."

End If

End Sub

Private Sub CmdEditComplete_Click(Index As Integer)

ms = MessBox("Are you done editing the file " + TheEditFile + "?" + vbCrLf + "If so, click YES to update it on the server," + vbCrLf + "or click NO not to update.", vbQuestion + vbYesNo, "File Edit")
If ms = vbYes Then PackageFileSend EditFileTemp, TheEditFile

CmdEditComplete(0).Enabled = False
CmdEditComplete(1).Enabled = False

End Sub

Private Sub CmdMove_Click(Index As Integer)
'MOVES FILES from EITHER:
'Server to Server - 1
'Server to Client - 2
'Client to Server - 3
'Client to Client (local move) - 4

'CURRENTLY ONLY MODE 1 DONE!

'see whats selected
If Index = 0 Then othr = 1
If Index = 1 Then othr = 0


If IsLocal(Index) = False Then
    
    'Fill The Array
        NumDirs(Index) = 0
    If Index = 0 Then ReDim Preserve DirList0(0 To 1)
    If Index = 1 Then ReDim Preserve DirList1(0 To 1)
    
    For i = 1 To ListView(Index).ListItems.Count
        With ListView(Index).ListItems.Item(i)
            If .Selected = True Then 'selected!
                NumDirs(Index) = NumDirs(Index) + 1
                            
                'Add this record
                If Index = 0 Then
                    ReDim Preserve DirList0(0 To NumDirs(Index))
                    
                    DirList0(NumDirs(Index)).DateTime = Now
                    DirList0(NumDirs(Index)).FullPath = DirFullPath(Index) + "\" + Trim(.Text)
                    DirList0(NumDirs(Index)).Name = Trim(.Text)
                    DirList0(NumDirs(Index)).Size = Trim(.SubItems(2))
                    If .SmallIcon = 1 Then DirList0(NumDirs(Index)).Type = 1
                    If .SmallIcon = 2 Then DirList0(NumDirs(Index)).Type = 0
                Else
                    ReDim Preserve DirList1(0 To NumDirs(Index))
                    
                    DirList1(NumDirs(Index)).DateTime = Now
                    DirList1(NumDirs(Index)).FullPath = DirFullPath(Index) + "\" + Trim(.Text)
                    DirList1(NumDirs(Index)).Name = Trim(.Text)
                    DirList1(NumDirs(Index)).Size = Trim(.SubItems(2))
                    If .SmallIcon = 1 Then DirList1(NumDirs(Index)).Type = 1
                    If .SmallIcon = 2 Then DirList1(NumDirs(Index)).Type = 0
                End If
            End If
        End With
    Next i

    'Send appropriate commands
    
    If IsLocal(othr) = False Then 'Its Mode 1
        'Send MOVE command
        SendPacket "F5", PackageDirList(Index, DirFullPath(othr))
    End If
Else
    If IsLocal(othr) = False Then
        'Window (index) copying to server
        
            
    Else
        'Mode 4 - Client to Client


    End If
End If



End Sub

Private Sub CmdRefresh_Click(Index As Integer)
'ask for a refresh of this item

a = Combo(Index).ListIndex
If a = -1 Then Exit Sub

RefreshDir Index, Combo(Index).List(a)


End Sub

Private Sub CmdRename_Click(Index As Integer)

'Edit this file

For i = 1 To ListView(Index).ListItems.Count
    With ListView(Index).ListItems.Item(i)
        If .Selected = True Then j = i: Exit For
    End With
Next i

If j > 0 Then
    
    'do shtuff
    
    If ListView(Index).ListItems.Item(j).SmallIcon = 2 Then 'file
        'edit
        b$ = Trim(ListView(Index).ListItems.Item(j).Text)
        RenameFiles b$, Index, DirFullPath(Index)
        
    End If
End If

End Sub

Private Sub CmdUp_Click(Index As Integer)

a = Combo(Index).ListIndex
If a = -1 Then Exit Sub

b$ = Combo(Index).List(a)

'go up a dir

f = InStrRev(b$, "\")

If f > 0 And f <> Len(b$) Then
    'extract path
    
    p$ = Left(b$, f - 1)
    
    RefreshDir Index, p$
End If

End Sub

Private Sub Combo_Click(Index As Integer)

'ask for a refresh of this item

a = Combo(Index).ListIndex
If a = -1 Then Exit Sub

If a = LastIndex(Index) Then Exit Sub
LastIndex(Index) = a


RefreshDir Index, Combo(Index).List(a)


End Sub

Private Sub Command1_Click()



End Sub

Private Sub Form_Load()
Perc = 0.5

IsLocal(0) = True
DirFullPath(0) = GetSetting("Server Assistant Client", "Files", "LocalPath", App.Path)

RefreshList 0

CmdUp(0).Picture = ImageList1.ListImages.Item(3).Picture
CmdUp(1).Picture = ImageList1.ListImages.Item(3).Picture

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

Private Sub Form_Resize()

'resize all elements

If Me.WindowState = vbMinimized Then Exit Sub

Dim Wid(0 To 1) As Long

If Me.Width < 6000 Then Me.Width = 6000
If Me.Height < 4000 Then Me.Height = 4000



'first, the main frames
w = Me.Width
h = Me.Height

w1 = Int(w * Perc) - 180
w12 = Int(w * (1 - Perc)) - 180

Wid(0) = w1
Wid(1) = w12

h1 = h - 520

Frame1(0).Width = w1
Frame1(1).Width = w12
Frame1(1).Left = Frame1(0).Left + w1 + 60

Label1.Left = Frame1(0).Left + w1

'resize elements inside frames

For i = 0 To 1

    Frame1(i).Height = h1

    w2 = DisFrame(i).Width
    h2 = Frame1(i).Height - 60 - ListView(i).Top
    
    Combo(i).Width = Abs(Wid(i) - w2 - 180)
    ListView(i).Width = Abs(Wid(i) - w2 - 180)
    
    ListView(i).Height = h2
    DisFrame(i).Left = Abs(Wid(i) - w2 - 60)
    DirFrame(i).Left = Abs(Wid(i) - w2 - 60)
    OpFrame(i).Left = Abs(Wid(i) - w2 - 60)

Next i

Label1.Height = Frame1(0).Height


'ProgressBar1.Top = 60 + Frame1(0).Top + Frame1(0).Height
'ProgressBar1.Width = w - 240 - ProgressBar1.Left - 60 - Command1.Width
'Command1.Left = ProgressBar1.Left + ProgressBar1.Width + 60


End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next
nm$ = Me.Name
SaveSetting "Server Assistant Client", "Window", nm$ + "winmd", Me.WindowState
SaveSetting "Server Assistant Client", "Window", nm$ + "winh", Me.Height
SaveSetting "Server Assistant Client", "Window", nm$ + "wint", Me.Top
SaveSetting "Server Assistant Client", "Window", nm$ + "winl", Me.Left
SaveSetting "Server Assistant Client", "Window", nm$ + "winw", Me.Width

SaveSetting "Server Assistant Client", "Files", "LocalPath", DirFullPath(0)

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    X1 = X + Label1.Left
    
    Perc = Abs(X1 / Me.Width)
    If Perc > 0.7 Then Perc = 0.7
    
    If Perc < 0.3 Then Perc = 0.3
    
    Form_Resize
End If

End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)



X1 = X + Label1.Left

Perc = Abs(X1 / Me.Width)
If Perc > 0.7 Then Perc = 0.7

If Perc < 0.3 Then Perc = 0.3

Form_Resize

End Sub

Private Sub ListView_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)

ListView(Index).Sorted = True
k = ListView(Index).SortKey

If k = (ColumnHeader.Index - 1) Then
    If ListView(Index).SortOrder = lvwDescending Then
        ListView(Index).SortOrder = lvwAscending
    Else
        ListView(Index).SortOrder = lvwDescending
    End If
End If

ListView(Index).SortKey = (ColumnHeader.Index - 1)


End Sub

Private Sub ListView_DblClick(Index As Integer)

'see whats selected

For i = 1 To ListView(Index).ListItems.Count
    With ListView(Index).ListItems.Item(i)
        If .Selected = True Then j = i: Exit For
    End With
Next i

If j > 0 Then
    
    'do shtuff
    
    If ListView(Index).ListItems.Item(j).SmallIcon = 2 Then 'file
        'edit
    Else    'dir
        'change dirs
        b$ = DirFullPath(Index) + "\" + Trim(ListView(Index).ListItems.Item(j).Text)
        RefreshDir Index, b$
    End If
End If

End Sub
Private Sub UpdateCombos()
    
    
For k = 0 To 1

    If IsLocal(k) Then
        'fill combo
        Combo(k).Clear
        For j = 1 To RecentLocal.Count
            Combo(k).AddItem RecentLocal(j)
        Next j
        If Combo(k).ListCount > 0 Then Combo(k).ListIndex = 0
    Else
        Combo(k).Clear
        For j = 1 To RecentRemote.Count
            Combo(k).AddItem RecentRemote(j)
        Next j
        If Combo(k).ListCount > 0 Then Combo(k).ListIndex = 0
    End If
Next k

End Sub


Private Sub ListView_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

Debug.Print Index, Source.Name, X, Y


End Sub

Private Sub OptLocal_Click(Index As Integer)

Dim blr As Boolean
blr = IsLocal(Index)
IsLocal(Index) = OptLocal(Index).Value
UpdateCombos

If IsLocal(Index) <> blr And Combo(Index).ListCount > 0 Then  ' a change
    DirFullPath(Index) = Combo(Index).List(0)
    RefreshDir Index, DirFullPath(Index)
End If

End Sub

Private Sub OptRemote_Click(Index As Integer)

Dim blr As Boolean
blr = IsLocal(Index)
IsLocal(Index) = OptLocal(Index).Value
UpdateCombos

If IsLocal(Index) <> blr And Combo(Index).ListCount > 0 Then  ' a change
    DirFullPath(Index) = Combo(Index).List(0)
    RefreshDir Index, DirFullPath(Index)
End If


End Sub
