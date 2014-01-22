VERSION 5.00
Begin VB.Form frmMessageBox 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   Icon            =   "frmMessageBox.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4350
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3480
      Top             =   480
   End
   Begin VB.PictureBox Image3 
      Height          =   555
      Left            =   2760
      Picture         =   "frmMessageBox.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Image2 
      Height          =   555
      Left            =   2160
      Picture         =   "frmMessageBox.frx":0884
      ScaleHeight     =   495
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Image1 
      Height          =   555
      Left            =   1560
      Picture         =   "frmMessageBox.frx":0CC6
      ScaleHeight     =   495
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   435
      Left            =   1980
      TabIndex        =   1
      Top             =   1380
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   1380
      Width           =   1815
   End
   Begin VB.Image ImageUse 
      Height          =   480
      Left            =   300
      Picture         =   "frmMessageBox.frx":1108
      Top             =   420
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frmMessageBox"
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

Public ShowMode As Boolean
Public Prompt As String
Public Buttons As Long
Public Title As String
Public ReturnValue As Long
Public TimeToShow As Long
Dim StartTime As Long


Public Sub Display()

StartTime = Timer

'Initialize
Label1 = Prompt

'Icons

If (Buttons And vbQuestion) = vbQuestion Then
    ImageUse.Picture = Image2.Picture
    img = 1
End If
If (Buttons And vbCritical) = vbCritical Then
    ImageUse.Picture = Image1.Picture
    img = 1
End If
If (Buttons And vbInformation) = vbInformation Then
    ImageUse.Picture = Image3.Picture
    img = 1
End If

If img = 1 Then ImageUse.Visible = True


Me.Width = Label1.Width + Label1.Left + 120
If img = 1 Then
    Me.Width = Me.Width + ImageUse.Left + ImageUse.Width + 60
    Label1.Left = ImageUse.Left + ImageUse.Width + 60
End If

If Me.Width < 5040 Then
    Me.Width = 5040
    If img = 1 Then
        Label1.Left = Int((Me.Width - ImageUse.Left + ImageUse.Width + 60) / 2) - Int(Label1.Width / 2) ' + ImageUse.Left + ImageUse.Width + 60
    Else
        Label1.Left = Int(Me.Width / 2) - Int(Label1.Width / 2)
    End If
End If

Me.Height = Label1.Height + Label1.Top + Command1.Height + 360
If Me.Height < 1695 Then
    Me.Height = 1695
    Label1.Top = Int((Me.Height - Command1.Height - 400) / 2) - Int(Label1.Height / 2)
End If
If img = 1 Then ImageUse.Top = Int((Me.Height - Command1.Height - 400) / 2) - Int(ImageUse.Height / 2)

Command1.Top = Me.Height - Command1.Height - 400
Command2.Top = Command1.Top

Me.Caption = Title
If Title = "" Then Me.Caption = App.Title

Me.Left = Int(MDIForm1.Width / 2) - Int(Me.Width / 2)
Me.Top = Int(MDIForm1.Height / 2) - Int(Me.Height / 2)

'Handle Buttons
If (Buttons And vbOKOnly) = vbOKOnly Then 'OK ONLY
    Command1.Caption = "OK"
    Command1.Default = True
    Command1.Left = Int(Me.Width / 2) - Int(Command1.Width / 2)
    'Command1.SetFocus
    Command2.Visible = False
    Command1.Tag = vbOK
    
End If
If (Buttons And vbOKCancel) = vbOKCancel Then 'OK and CANCEL
    
    Command1.Caption = "OK"
    Command1.Default = True
    'Command1.SetFocus
    Command1.Tag = vbOK
    
    Command2.Caption = "Cancel"
    Command2.Cancel = True
    
    Command1.Left = Int(Me.Width / 2) - Int((Command1.Width + Command2.Width + 120) / 2)
    Command2.Left = Command1.Left + Command1.Width + 120
    Command2.Visible = True
    Command2.Tag = vbCancel
End If
If (Buttons And vbYesNo) = vbYesNo Then 'YES and NO
    
    Command1.Caption = "Yes"
    Command1.Default = True
    'Command1.SetFocus
    Command1.Tag = Ts(vbYes)
    
    Command2.Caption = "No"
    Command2.Cancel = True
    
    Command1.Left = Int(Me.Width / 2) - Int((Command1.Width + Command2.Width + 120) / 2)
    Command2.Left = Command1.Left + Command1.Width + 120
    Command2.Visible = True
    Command2.Tag = vbNo
End If

'Icon

Me.Visible = True

Me.Show

Beep

End Sub

Private Sub Command1_Click()

ReturnValue = Val(Command1.Tag)
If ShowMode Then Unload Me

End Sub

Private Sub Command2_Click()
ReturnValue = Val(Command2.Tag)
If ShowMode Then Unload Me
End Sub

Private Sub Timer1_Timer()

If TimeToShow > 0 Then
    If Timer - StartTime > TimeToShow Then Unload Me
End If

End Sub
