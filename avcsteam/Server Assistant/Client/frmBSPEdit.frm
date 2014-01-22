VERSION 5.00
Begin VB.Form frmBSPEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "BSP Edit"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2025
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   2025
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Height          =   1095
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1935
   End
End
Attribute VB_Name = "frmBSPEdit"
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

Dim FilePa As String
Dim FileLocPa As String

Private Sub Command1_Click()
'Send the file back
If CheckForFile(FileLocPa) = False Then
    MessBox "File not Found!"
    Unload Me
    Exit Sub
End If

h = FreeFile
ret$ = ""
Open FileLocPa For Binary As h
    Do While Not EOF(h)
        ret$ = ret$ + Input(65000, #h)
    Loop
Close h

Kill FileLocPa

a$ = a$ + Chr(251)
a$ = a$ + FilePa + Chr(250)
a$ = a$ + ret$ + Chr(250)
a$ = a$ + Chr(251)

SendPacket "B1", a$

Unload Me

End Sub

Private Sub Form_Load()
FilePa = FilePath
FileLocPa = FileLocalPath + ".ent"

Command1.Caption = "CLICK ME" + vbCrLf + "When you are done editing" + vbCrLf + FileLocalPath
End Sub
