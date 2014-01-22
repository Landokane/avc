VERSION 5.00
Begin VB.Form frmServerLog 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Server Log"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5865
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   4095
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4395
   End
End
Attribute VB_Name = "frmServerLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Resize()

If Me.WindowState = 1 Then Exit Sub

If Me.Width < 1000 Then Me.Width = 1000
If Me.Height < 1000 Then Me.Height = 1000

w = Me.Width
h = Me.Height

Text1.Height = h - Text1.Top - 360
Text1.Width = w - Text1.Left - 120

End Sub
