VERSION 5.00
Begin VB.Form frmEvents 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Current Events"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmEvents.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   3540
      Width           =   2595
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add New Event..."
      Height          =   375
      Left            =   5340
      TabIndex        =   6
      Top             =   3540
      Width           =   1875
   End
   Begin VB.Frame Frame1 
      Caption         =   "Event Details"
      Height          =   1035
      Left            =   0
      TabIndex        =   1
      Top             =   2460
      Width           =   7215
      Begin VB.CommandButton Command2 
         Caption         =   "Delete this Event"
         Height          =   375
         Left            =   5280
         TabIndex        =   5
         Top             =   600
         Width           =   1875
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Edit this Event"
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   600
         Width           =   1755
      End
      Begin VB.Label lblNextRun 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   6075
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Next run at:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   825
      End
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "frmEvents"
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

Public Sub UpdateList()

List1.Clear

For i = 1 To NumEvents
    List1.AddItem Events(i).Name
    List1.ItemData(List1.NewIndex) = i
Next i

End Sub

Private Sub Command1_Click()

'edit event
a = List1.ListIndex
If a = -1 Then Exit Sub
e = List1.ItemData(a)

frmAddEvent.EditNum = e
frmAddEvent.EditMode = True
frmAddEvent.Show

End Sub

Private Sub Command2_Click()

'delete event
a = List1.ListIndex
If a = -1 Then Exit Sub
e = List1.ItemData(a)

m = MessBox("Are you sure you want to delete this event?", vbYesNo, "Delete Event")

If m = vbYes Then
    'send packet to server, telling it to delete this event
    SendPacket "DE", Events(e).Name
End If

End Sub

Private Sub Command3_Click()

'add a new event
frmAddEvent.EditMode = False
frmAddEvent.EditNum = 0
frmAddEvent.Show



End Sub

Private Sub Command4_Click()
Unload Me

End Sub

Private Sub Form_Load()
UpdateList

End Sub

Private Sub List1_Click()

'show data

a = List1.ListIndex
If a = -1 Then Exit Sub

e = List1.ItemData(a)
lblNextRun = Format(Events(e).FirstCheck, "dddd, mmm d yyyy, hh:mm:ss")

End Sub
