VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit User"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   ControlBox      =   0   'False
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   Begin VB.VScrollBar VScroll1 
      Height          =   2535
      LargeChange     =   500
      Left            =   4740
      SmallChange     =   90
      TabIndex        =   57
      Top             =   4380
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   60
      ScaleHeight     =   2475
      ScaleWidth      =   4575
      TabIndex        =   17
      Top             =   4380
      Width           =   4635
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   10755
         Left            =   0
         ScaleHeight     =   10755
         ScaleWidth      =   4515
         TabIndex        =   18
         Top             =   0
         Width           =   4515
         Begin VB.CheckBox Checks 
            Caption         =   "Can control HLDS"
            Height          =   255
            Index           =   36
            Left            =   60
            TabIndex        =   62
            Top             =   10200
            Width           =   3675
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Recieves ICQ when server stops responding"
            Height          =   255
            Index           =   35
            Left            =   60
            TabIndex        =   61
            Top             =   9960
            Width           =   3675
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Is CHAT ONLY admin."
            Height          =   255
            Index           =   34
            Left            =   60
            TabIndex        =   60
            Top             =   2760
            Width           =   3795
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to ent edit BSP files"
            Height          =   255
            Index           =   33
            Left            =   60
            TabIndex        =   59
            Top             =   9720
            Width           =   3375
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Can recieve ICQ messages from Players in-game"
            Height          =   255
            Index           =   32
            Left            =   60
            TabIndex        =   58
            Top             =   1020
            Width           =   3795
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to edit Bad Word List"
            Height          =   255
            Index           =   31
            Left            =   60
            TabIndex        =   56
            Top             =   9480
            Width           =   3375
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to see current server log"
            Height          =   255
            Index           =   8
            Left            =   60
            TabIndex        =   54
            Top             =   7560
            Width           =   3375
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to edit users"
            Height          =   255
            Index           =   11
            Left            =   60
            TabIndex        =   53
            Top             =   7800
            Width           =   3375
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to see client login list"
            Height          =   255
            Index           =   15
            Left            =   60
            TabIndex        =   52
            Top             =   7320
            Width           =   3375
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to edit Kick/Ban List"
            Height          =   255
            Index           =   18
            Left            =   60
            TabIndex        =   51
            Top             =   8040
            Width           =   3375
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to edit Clan Lists"
            Height          =   255
            Index           =   19
            Left            =   60
            TabIndex        =   50
            Top             =   8280
            Width           =   3375
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to edit Admin Speech"
            Height          =   255
            Index           =   20
            Left            =   60
            TabIndex        =   49
            Top             =   8520
            Width           =   3375
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to edit Real Players"
            Height          =   255
            Index           =   21
            Left            =   60
            TabIndex        =   48
            Top             =   8760
            Width           =   3375
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to edit Web Info"
            Height          =   255
            Index           =   22
            Left            =   60
            TabIndex        =   47
            Top             =   9000
            Width           =   3375
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to edit General Info"
            Height          =   255
            Index           =   24
            Left            =   60
            TabIndex        =   46
            Top             =   9240
            Width           =   3375
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to delete/move/copy files/dirs"
            Height          =   255
            Index           =   25
            Left            =   60
            TabIndex        =   44
            Top             =   6600
            Width           =   3315
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to view directories other than base dir"
            Height          =   255
            Index           =   26
            Left            =   60
            TabIndex        =   43
            Top             =   5880
            Width           =   3615
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to view directories"
            Height          =   255
            Index           =   27
            Left            =   60
            TabIndex        =   42
            Top             =   5640
            Width           =   3375
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to download files"
            Height          =   255
            Index           =   28
            Left            =   60
            TabIndex        =   41
            Top             =   6120
            Width           =   3735
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to rename files"
            Height          =   255
            Index           =   29
            Left            =   60
            TabIndex        =   40
            Top             =   6840
            Width           =   3375
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to upload files"
            Height          =   255
            Index           =   30
            Left            =   60
            TabIndex        =   39
            Top             =   6360
            Width           =   3795
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to run Button Scripts"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   33
            Top             =   3240
            Width           =   3795
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to log in"
            Height          =   255
            Index           =   1
            Left            =   60
            TabIndex        =   32
            Top             =   300
            Width           =   3795
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to chat with users on server"
            Height          =   255
            Index           =   2
            Left            =   60
            TabIndex        =   31
            Top             =   1800
            Width           =   3795
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to start mapvote"
            Height          =   255
            Index           =   3
            Left            =   60
            TabIndex        =   30
            Top             =   3480
            Width           =   3795
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to edit scripts"
            Height          =   255
            Index           =   4
            Left            =   60
            TabIndex        =   29
            Top             =   4680
            Width           =   3375
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to edt server info"
            Height          =   255
            Index           =   5
            Left            =   60
            TabIndex        =   28
            Top             =   4920
            Width           =   3375
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to view simple logs in game console"
            Height          =   255
            Index           =   6
            Left            =   60
            TabIndex        =   27
            Top             =   4200
            Width           =   3795
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to edit events (scheduler)"
            Height          =   255
            Index           =   7
            Left            =   60
            TabIndex        =   26
            Top             =   5160
            Width           =   3795
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to change their password"
            Height          =   255
            Index           =   9
            Left            =   60
            TabIndex        =   25
            Top             =   540
            Width           =   3795
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to monitor server"
            Height          =   255
            Index           =   10
            Left            =   60
            TabIndex        =   24
            Top             =   1560
            Width           =   3795
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to use RCON"
            Height          =   255
            Index           =   12
            Left            =   60
            TabIndex        =   23
            Top             =   3960
            Width           =   3795
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to see list of users on server"
            Height          =   255
            Index           =   13
            Left            =   60
            TabIndex        =   22
            Top             =   2040
            Width           =   3795
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to kick players on server"
            Height          =   255
            Index           =   14
            Left            =   60
            TabIndex        =   21
            Top             =   2280
            Width           =   3795
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to ban users on server"
            Height          =   255
            Index           =   16
            Left            =   60
            TabIndex        =   20
            Top             =   2520
            Width           =   3795
         End
         Begin VB.CheckBox Checks 
            Caption         =   "Allowed to leave messages for other users"
            Height          =   255
            Index           =   17
            Left            =   60
            TabIndex        =   19
            Top             =   780
            Width           =   3795
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Admin"
            Height          =   195
            Left            =   60
            TabIndex        =   55
            Top             =   7080
            Width           =   435
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "File Transfer"
            Height          =   195
            Left            =   60
            TabIndex        =   45
            Top             =   5460
            Width           =   870
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "General"
            Height          =   195
            Left            =   60
            TabIndex        =   38
            Top             =   60
            Width           =   555
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Monitor"
            Height          =   195
            Left            =   60
            TabIndex        =   37
            Top             =   1320
            Width           =   525
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Rcon"
            Height          =   195
            Left            =   60
            TabIndex        =   36
            Top             =   3780
            Width           =   390
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Scripts"
            Height          =   195
            Left            =   60
            TabIndex        =   35
            Top             =   3060
            Width           =   480
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Editing"
            Height          =   195
            Left            =   60
            TabIndex        =   34
            Top             =   4500
            Width           =   480
         End
      End
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1140
      TabIndex        =   15
      Top             =   1260
      Width           =   3855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1140
      TabIndex        =   13
      Top             =   1620
      Width           =   3855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   60
      TabIndex        =   12
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CheckBox Checks 
      Caption         =   "Unallowed"
      Height          =   195
      Index           =   23
      Left            =   3900
      TabIndex        =   11
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      Height          =   315
      Left            =   4380
      TabIndex        =   10
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Use"
      Height          =   315
      Left            =   3900
      TabIndex        =   9
      Top             =   60
      Width           =   435
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   660
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   60
      Width           =   3195
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   6960
      Width           =   3315
   End
   Begin VB.TextBox Text3 
      Height          =   2025
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2280
      Width           =   4935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1140
      TabIndex        =   3
      Top             =   900
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1140
      TabIndex        =   1
      Top             =   540
      Width           =   3855
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "ICQ Number"
      Height          =   195
      Left            =   60
      TabIndex        =   16
      Top             =   1320
      Width           =   870
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "FTP Root"
      Height          =   195
      Left            =   60
      TabIndex        =   14
      Top             =   1680
      Width           =   690
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   4980
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Presets"
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   120
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Allowed Commands:"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   2040
      Width           =   1725
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   960
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   420
   End
End
Attribute VB_Name = "Form11"
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

Dim b As Integer

Private Sub Checks_Click(Index As Integer)

If Index = 23 Then
    If Checks(Index) = 1 Then Label3 = "Unallowed Commands"
    If Checks(Index) = 0 Then Label3 = "Allowed Commands"
End If

End Sub

Private Sub Command1_Click()

Users(b).Name = Text1
Users(b).PassWord = Text2
Users(b).Allowed = Text3
Users(b).ICQ = Text5
Users(b).Email = Text4


're-import the bitfields

a = 0
For i = 0 To Checks.Count - 1
    c = 2 ^ i
    
    If Checks(i).Value = 1 Then a = a + c
Next i

Users(b).Flags = a

frmUserList.UpdateList

Unload Me


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()

a = Combo1.ListIndex
If a = -1 Then Exit Sub
a = a + 1

b = UserEditNum
Users(b).Allowed = Presets(a).Allowed
Users(b).Flags = Presets(a).Flags

Text3 = Users(b).Allowed

'now do bitfields

a = Users(b).Flags

For i = Checks.Count - 1 To 0 Step -1
    c = 2 ^ i
    
    If a >= c Then
        a = a - c
        Checks(i).Value = 1
    Else
        Checks(i).Value = 0
    End If
Next i


End Sub

Private Sub Command4_Click()

a$ = InBox("Name?")
If a$ = "" Then Exit Sub
b = UserEditNum

aa = 0
For i = 0 To Checks.Count - 1
    c = 2 ^ i
    
    If Checks(i).Value = 1 Then aa = aa + c
Next i


NumPresets = NumPresets + 1
Presets(NumPresets).Allowed = Text3
Presets(NumPresets).Flags = aa
Presets(NumPresets).Name = a$

UpdatePresets





End Sub

Private Sub UpdatePresets()

Combo1.Clear
For i = 1 To NumPresets
    Combo1.AddItem Presets(i).Name, i - 1
Next i

If Combo1.ListCount > 0 Then Combo1.ListIndex = 0


End Sub

Private Sub Command5_Click()
Unload Me

End Sub

Private Sub Form_Load()

VScroll1.Max = Picture2.Height - Picture1.Height

'get the info

b = UserEditNum

Text1 = Users(b).Name
Text2 = Users(b).PassWord
Text3 = Users(b).Allowed
Text5 = Users(b).ICQ
Text4 = Users(b).Email

'now do bitfields

a = Users(b).Flags

For i = Checks.Count - 1 To 0 Step -1
    c = 2 ^ i
    
    If a >= c Then
        a = a - c
        Checks(i).Value = 1
    End If
Next i

If Checks(23) = 1 Then Label3 = "Unallowed Commands"
If Checks(23) = 0 Then Label3 = "Allowed Commands"


UpdatePresets


End Sub

Private Sub VScroll1_Change()

Picture2.Top = -VScroll1.Value


End Sub

Private Sub VScroll1_Scroll()
VScroll1_Change
End Sub
