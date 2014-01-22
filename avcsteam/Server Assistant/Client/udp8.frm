VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Text Window"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   Icon            =   "udp8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5175
   Begin VB.TextBox Text2 
      Height          =   1815
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   2100
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Height          =   1755
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   300
      Width           =   5175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "This window will help you by turning ENTER into \n, "" into \q, and so on."
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5145
   End
End
Attribute VB_Name = "Form8"
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

Dim HasFocus As Integer
Dim Text2Top As Integer
Dim Text1H  As Integer
Dim Text1W  As Integer
Dim FormH As Integer
Dim FormW As Integer

Private Sub Form_Load()

Text2Top = Text2.Top
Text1H = Text1.Height
Text1W = Text1.Width
FormH = Me.Height
FormW = Me.Width

End Sub

Private Sub Form_Resize()

If Text2Top = 0 Then Exit Sub

If Me.Width < 1000 Then Me.Width = 1000
If Me.Height < 1000 Then Me.Height = 1000


Text1.Height = Int((Me.Height - Text1.Top - (Text2Top - Text1H - Text1.Top) - 395) / 2)
Text2.Top = Text1.Top + Text1.Height + (Text2Top - Text1H - Text1.Top)
Text2.Height = Text1.Height

Text1.Width = Me.Width - Text1.Left - (FormW - Text1.Left - Text1W)
Text2.Width = Text1.Width




End Sub

Private Sub Text1_Change()

If HasFocus = 1 Then
    a$ = Text1
    a$ = ReplaceString(a$, "\", Chr(255))
    a$ = ReplaceString(a$, vbCrLf, "\n")
    a$ = ReplaceString(a$, Chr(34), "\q")
    a$ = ReplaceString(a$, Chr(255), "\\")
    Text2 = a$
End If

End Sub

Private Sub Text1_GotFocus()
HasFocus = 1
End Sub

Private Sub Text2_Change()

If HasFocus = 2 Then
    a$ = Text2
    a$ = ReplaceString(a$, "\\", Chr(255))
    a$ = ReplaceString(a$, "\n", vbCrLf)
    a$ = ReplaceString(a$, "\q", Chr(34))
    a$ = ReplaceString(a$, Chr(255), "\")
    Text1 = a$
End If

End Sub

Private Sub Text2_GotFocus()
HasFocus = 2
End Sub
