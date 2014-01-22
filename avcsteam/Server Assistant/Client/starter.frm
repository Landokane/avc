VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   Icon            =   "starter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   2475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   60
      Top             =   60
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EXE As String
Dim Timer As Integer

Private Sub Form_Load()

'EXE = Command$
'EXE = Trim(EXE)

'If EXE = "" Then MsgBox "This program is for internal use only.": End

EXE = App.Path + "\server.exe"

If Dir(EXE) = "" Then End

End Sub

Private Sub Timer1_Timer()
Timer = Timer + 1

On Error Resume Next

If Timer = 10 Then
    
    'kill server.exe
    Randomize
    a$ = Trim(Str((Int(Rnd * 1000) + 1)))
    
    'Name App.Path + "\server.exe" As App.Path + "\server_" + a$ + "_OLD.exe"
End If

If Timer = 12 Then
    
    'new one
    'FileCopy EXE, App.Path + "\server.exe"
    'Kill EXE
End If

If Timer = 5 Then
    
    'start new one
    Shell App.Path + "\server.exe", vbMinimizedNoFocus
End If

If Timer = 7 Then End

End Sub
