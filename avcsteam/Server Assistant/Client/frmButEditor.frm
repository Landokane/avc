VERSION 5.00
Begin VB.Form frmButEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Button Editor"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "frmButEditor.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5280
   Begin VB.CommandButton Command4 
      Caption         =   "Close"
      Height          =   255
      Left            =   60
      TabIndex        =   18
      Top             =   4200
      Width           =   5175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Control Edit"
      Height          =   2115
      Left            =   60
      TabIndex        =   6
      Top             =   2040
      Width           =   5175
      Begin VB.OptionButton Option4 
         Caption         =   "Multi-Line Text Box"
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   540
         Width           =   1695
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Players Menu"
         Height          =   255
         Left            =   3720
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Top             =   1740
         Visible         =   0   'False
         Width           =   3675
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   1440
         Width           =   3675
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Check Box"
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   240
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Text Box"
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   1140
         Visible         =   0   'False
         Width           =   3675
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   840
         Width           =   3675
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Unchecked Value"
         Height          =   195
         Left            =   60
         TabIndex        =   16
         Top             =   1800
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Default Value"
         Height          =   195
         Left            =   60
         TabIndex        =   14
         Top             =   1500
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         Height          =   195
         Left            =   60
         TabIndex        =   11
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Text on Control"
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Top             =   1200
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Control Name"
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   900
         Width           =   960
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      Height          =   255
      Left            =   780
      TabIndex        =   5
      Top             =   1740
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   1740
      Width           =   675
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1020
      TabIndex        =   1
      Top             =   60
      Width           =   4215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Controls"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   420
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Script Name"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   870
   End
End
Attribute VB_Name = "frmButEditor"
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

Dim CurrIndex As Integer
Dim ButIndex As Integer

Private Sub Command1_Click()

For i = 1 To Commands(ButIndex).NumButtons
    If Commands(ButIndex).Buttons(i).Type = 3 Then
        MessBox "Can't add controls when one is set to PLAYERS MENU type!"
        Exit Sub
    End If
Next i
'add a control

Commands(ButIndex).NumButtons = Commands(ButIndex).NumButtons + 1
a = Commands(ButIndex).NumButtons

ReDim Preserve Commands(ButIndex).Buttons(0 To a)

Commands(ButIndex).Buttons(a).ButtonName = "Unnamed"
Commands(ButIndex).Buttons(a).Type = 1

DrawList

End Sub

Private Sub Command2_Click()

a = List1.ListIndex
If a = -1 Then Exit Sub
e = CurrIndex

'remove

Commands(ButIndex).NumButtons = Commands(ButIndex).NumButtons - 1

For i = e To Commands(ButIndex).NumButtons
    
    Commands(ButIndex).Buttons(i).ButtonName = Commands(ButIndex).Buttons(i + 1).ButtonName
    Commands(ButIndex).Buttons(i).ButtonText = Commands(ButIndex).Buttons(i + 1).ButtonText
    Commands(ButIndex).Buttons(i).OptionOff = Commands(ButIndex).Buttons(i + 1).OptionOff
    Commands(ButIndex).Buttons(i).OptionOn = Commands(ButIndex).Buttons(i + 1).OptionOn
    Commands(ButIndex).Buttons(i).Type = Commands(ButIndex).Buttons(i + 1).Type

Next i

ReDim Preserve Commands(ButIndex).Buttons(0 To Commands(ButIndex).NumButtons)

CurrIndex = 0

DrawList

End Sub

Private Sub Command4_Click()
Commands(ButIndex).ScriptName = Text1
Commands(ButIndex).Changed = True

Unload Me




End Sub

Private Sub Form_Load()

ButIndex = EditedButton
Me.Caption = "Button Editor - Script " + Chr(34) + Commands(ButIndex).Name + Chr(34)

Text1 = Commands(ButIndex).ScriptName

DrawList

End Sub

Private Sub DrawList()

List1.Clear

For i = 1 To Commands(ButIndex).NumButtons
    List1.AddItem Ts(i) + " - " + Commands(ButIndex).Buttons(i).ButtonName
Next i

End Sub

Private Sub DisplayControl()

If CurrIndex = 0 Then Exit Sub

Text2 = Commands(ButIndex).Buttons(CurrIndex).ButtonName
Text3 = Commands(ButIndex).Buttons(CurrIndex).ButtonText
Text4 = Commands(ButIndex).Buttons(CurrIndex).OptionOn
Text5 = Commands(ButIndex).Buttons(CurrIndex).OptionOff

If Commands(ButIndex).Buttons(CurrIndex).Type = 1 Then Option1 = True
If Commands(ButIndex).Buttons(CurrIndex).Type = 2 Then Option2 = True
If Commands(ButIndex).Buttons(CurrIndex).Type = 3 Then Option3 = True
If Commands(ButIndex).Buttons(CurrIndex).Type = 4 Then Option4 = True

Option1_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)
Commands(ButIndex).ScriptName = Text1
Commands(ButIndex).Changed = True
End Sub

Private Sub List1_Click()

CurrIndex = List1.ListIndex + 1
DisplayControl

End Sub

Private Sub Option1_Click()
OptionClick

End Sub

Private Sub Option2_Click()
OptionClick

End Sub

Private Sub Option3_Click()

'check to make sure it's ok to set this.

If Commands(ButIndex).NumButtons > 1 Then
    MessBox "Can't set to PLAYERS MENU with more than one control!"
    Option1.Value = True
End If

OptionClick

End Sub

Sub OptionClick()



If Option1.Value = True Then
    Label6 = "Default Value"
    Label7.Visible = False
    Text5.Visible = False
    Text3.Visible = False
    Label4.Visible = False
    Label3 = "Control Name"
    Label4 = "Text on Control"
    Text4.Visible = True
    Label7.Visible = True
    Text5.Visible = True
    Label6.Visible = True
    
    If CurrIndex > 0 Then Commands(ButIndex).Buttons(CurrIndex).Type = 1
End If
If Option2.Value = True Then
    Label6 = "Checked Value"
    Label7.Visible = True
    Label3 = "Control Name"
    Text5.Visible = True
    Text3.Visible = True
    Label4.Visible = True
    Label4 = "Text on Control"
    Text4.Visible = True
    Label7.Visible = True
    Text5.Visible = True
    Label6.Visible = True
    
    If CurrIndex > 0 Then Commands(ButIndex).Buttons(CurrIndex).Type = 2
End If
If Option3.Value = True Then
    
    Label3 = "Menu Name"
    Label4 = "Question (optional)"
    Label4.Visible = True
    Label6.Visible = False
    Text3.Visible = True
    Text4.Visible = False
    Text5.Visible = False
    Label7.Visible = False
    
    If CurrIndex > 0 Then Commands(ButIndex).Buttons(CurrIndex).Type = 3
End If


If Option4.Value = True Then
    Label6 = "Default Value"
    Label7.Visible = False
    Text5.Visible = False
    Text3.Visible = False
    Label4.Visible = False
    Label3 = "Control Name"
    Label4 = "Text on Control"
    Text4.Visible = True
    Label7.Visible = True
    Text5.Visible = True
    Label6.Visible = True
    
    If CurrIndex > 0 Then Commands(ButIndex).Buttons(CurrIndex).Type = 4
End If

End Sub

Private Sub Option4_Click()
OptionClick
End Sub

Private Sub Text2_Change()

If CurrIndex > 0 Then Commands(ButIndex).Buttons(CurrIndex).ButtonName = Text2

End Sub

Private Sub Text3_Change()

If CurrIndex > 0 Then Commands(ButIndex).Buttons(CurrIndex).ButtonText = Text3

End Sub

Private Sub Text4_Change()

If CurrIndex > 0 Then Commands(ButIndex).Buttons(CurrIndex).OptionOn = Text4

End Sub

Private Sub Text5_Change()

If CurrIndex > 0 Then Commands(ButIndex).Buttons(CurrIndex).OptionOff = Text5

End Sub
