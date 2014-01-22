VERSION 5.00
Begin VB.Form frmInputBox 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   840
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   480
      TabIndex        =   1
      Top             =   1260
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   2400
      TabIndex        =   0
      Top             =   1260
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "something"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   4560
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmInputBox"
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

Public Prompt As String
Public Title As String
Public ReturnValue As String
Public Default As String
Public Finito As Integer

Public Sub Display()

'Initialize
Label1 = Prompt

'Icons

Me.Caption = Title
If Title = "" Then Me.Caption = App.Title

Me.Left = Int(MDIForm1.Width / 2) - Int(Me.Width / 2)
Me.Top = Int(MDIForm1.Height / 2) - Int(Me.Height / 2)

Text1 = Default


'Icon
Me.Visible = True
Me.Show

Text1.SetFocus
Text1.SelLength = Len(Text1)

End Sub

Private Sub Command1_Click()

ReturnValue = Text1
Finito = 1

End Sub

Private Sub Command2_Click()
ReturnValue = ""
Finito = 1

End Sub

Private Sub ImageUse_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)

If Finito = 0 Then Finito = 1: Cancel = 1

End Sub

Private Sub Timer1_Timer()

If TimeToShow > 0 Then
    If Timer - StartTime > TimeToShow Then Unload Me
End If

End Sub

