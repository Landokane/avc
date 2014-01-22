VERSION 5.00
Begin VB.Form frmLogSearch 
   AutoRedraw      =   -1  'True
   Caption         =   "Server Log"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8970
   Icon            =   "frmlogsearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4365
   ScaleWidth      =   8970
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   675
      Left            =   7860
      TabIndex        =   5
      Top             =   3660
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Frame Frame1 
      Caption         =   "Log Properties"
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   1860
      Width           =   8895
      Begin VB.CommandButton Command1 
         Caption         =   "View This Log"
         Height          =   315
         Left            =   7320
         TabIndex        =   4
         Top             =   600
         Width           =   1515
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   60
         TabIndex        =   2
         Top             =   240
         Width           =   8775
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   60
         TabIndex        =   3
         Top             =   600
         Width           =   7215
      End
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "frmLogSearch"
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



Private Sub Command1_Click()

If Label1 <> "" Then SendPacket "VL", Label1

'For I = 1 To NumLogFound
 '   If InStr(1, LCase(LogFound(I).LogLine), "icq") Then List1.ListIndex = I - 1
'Next I


End Sub

Private Sub Command2_Click()




Open App.Path + "\kirk.txt" For Append As #1
'L 06/23/2001 - 21:34:17: "cptkirk<86><109136><Blue>" say "hurry people lets send this roc kto the sext sex planet of sexus six"
For i = 1 To NumLogFound
    a$ = LogFound(i).LogLine

    e = InStr(1, a$, Chr(34))
    f = InStr(e + 1, a$, "<")
    
    nm$ = Mid(a$, e + 1, f - e - 1)
    
    e = InStr(1, a$, Chr(34) + " say")
    e = InStr(e + 1, a$, Chr(34))
    
    f = InStrRev(a$, Chr(34))
    
    ln$ = Mid(a$, e + 1, f - e - 1)
    
    
    Print #1, nm$ + ": " + ln$


Next i
Close #1
End Sub

Private Sub Form_Load()

'Fill the list box
On Error Resume Next

List1.Clear

For i = 1 To NumLogFound

    List1.AddItem LogFound(i).LogLine
    e = List1.NewIndex
    List1.ItemData(e) = i

Next i

Me.Caption = "Log Search Results - " + Ts(NumLogFound) + " results."



End Sub

Private Sub Form_Resize()

If Me.WindowState = 1 Then Exit Sub

If Me.WindowState = 0 Then
    If Me.Width < 2000 Then Me.Width = 2000
    If Me.Height < 2000 Then Me.Height = 2000
End If

w = Me.Width
h = Me.Height

List1.Width = w - 180 - List1.Left
Frame1.Width = w - 180 - Frame1.Left

List1.Height = h - List1.Top - 360 - Frame1.Height - 60
Frame1.Top = List1.Top + List1.Height + 60

Text1.Width = Frame1.Width - Text1.Left - 120
Label1.Width = Frame1.Width - Label1.Left - 180 - Command1.Width
Command1.Left = Label1.Left + Label1.Width + 60



End Sub

Private Sub List1_Click()
'display

a = List1.ListIndex
If a = -1 Then Exit Sub
e = List1.ItemData(a)

Text1 = LogFound(e).LogLine
Label1 = LogFound(e).LogFile

End Sub
