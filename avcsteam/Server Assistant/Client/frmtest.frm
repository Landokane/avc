VERSION 5.00
Begin VB.Form frmTest 
   BackColor       =   &H00000000&
   Caption         =   "Form2"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8730
   LinkTopic       =   "Form2"
   ScaleHeight     =   5310
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   4800
      Width           =   555
   End
   Begin VB.Label lblMain 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   8175
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TextLines(1 To 1000) As String

Public Sub LoadLabels()

h = Me.Height

h = h - 600 - lblMain(0).Top
h = Int(h / lblMain(0).Height)

For I = 1 To h
    Load lblMain(I)
    
    lblMain(I).Left = lblMain(0).Left
    lblMain(I).Top = lblMain(0).Top + (lblMain(0).Height * I)
    lblMain(I).Visible = True
    lblMain(I) = "Line " + Ts(I)

Next I


End Sub

Public Sub UpdateLabels()




End Sub

Private Sub Command1_Click()

LoadLabels

End Sub

