VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConnectUsers 
   Caption         =   "Connected Users"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   Icon            =   "frmConnectUsers.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   7845
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   1620
      ScaleHeight     =   795
      ScaleWidth      =   855
      TabIndex        =   1
      Top             =   3660
      Visible         =   0   'False
      Width           =   915
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   3540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   6059
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   16777215
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3200
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Version"
         Object.Width           =   1295
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   1826
      EndProperty
   End
End
Attribute VB_Name = "frmConnectUsers"
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

Public OldWindowProc As Long
Private Sub Form_Load()
UpdateUsersList
ShowUsers = True
UpdateLogDetail


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

AddForm True, 67, 67, 0, 0, Me


n = ListView1.ColumnHeaders.Count

For i = 1 To n
    wp = GetSetting("Server Assistant Client", "UserList", ListView1.ColumnHeaders(i).Text + "pos", -1)
    If wp <> -1 Then ListView1.ColumnHeaders(i).Position = wp
    wp = GetSetting("Server Assistant Client", "UserList", ListView1.ColumnHeaders(i).Text + "wid", -1)
    If wp <> -1 Then ListView1.ColumnHeaders(i).Width = wp
Next i


End Sub

Private Sub Form_Resize()

If Me.WindowState = 1 Then Exit Sub

If Me.Width < 1000 Then Me.Width = 1000
If Me.Height < 1000 Then Me.Height = 1000

w = Me.Width
h = Me.Height

ListView1.Height = h - ListView1.Top - 120 - 320
ListView1.Width = w - ListView1.Left - 120

End Sub

Private Sub Form_Unload(Cancel As Integer)
ShowUsers = False
UpdateLogDetail



On Error Resume Next
nm$ = Me.Name
SaveSetting "Server Assistant Client", "Window", nm$ + "winmd", Me.WindowState
SaveSetting "Server Assistant Client", "Window", nm$ + "winh", Me.Height
SaveSetting "Server Assistant Client", "Window", nm$ + "wint", Me.Top
SaveSetting "Server Assistant Client", "Window", nm$ + "winl", Me.Left
SaveSetting "Server Assistant Client", "Window", nm$ + "winw", Me.Width

n = ListView1.ColumnHeaders.Count

For i = 1 To n
    SaveSetting "Server Assistant Client", "UserList", ListView1.ColumnHeaders(i).Text + "pos", Ts(ListView1.ColumnHeaders(i).Position)
    SaveSetting "Server Assistant Client", "UserList", ListView1.ColumnHeaders(i).Text + "wid", Ts(ListView1.ColumnHeaders(i).Width)
Next i


End Sub

Private Sub List1_Click()
'
'

End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)



End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


'istview1.ColumnHeaders(ListView1.SortKey + 1).for

ListView1.Sorted = True
k = ListView1.SortKey

If k = (ColumnHeader.Index - 1) Then
    If ListView1.SortOrder = lvwDescending Then
        ListView1.SortOrder = lvwAscending
    Else
        ListView1.SortOrder = lvwDescending
    End If
End If

ListView1.SortKey = (ColumnHeader.Index - 1)

End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
        
        
    On Error Resume Next
    e = Val(frmConnectUsers.ListView1.SelectedItem.Tag)
    If e = 0 Then Exit Sub
        
    
    If ConnectUsers(e).AwayMode Then
        MDIForm1.mnuGamesIn(0).Visible = True
    Else
        MDIForm1.mnuGamesIn(0).Visible = False
    End If
    
    
    MDIForm1.PopupMenu MDIForm1.mnuGames

End If
End Sub
