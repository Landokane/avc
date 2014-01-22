VERSION 5.00
Begin VB.Form frmControlFill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Script Properties"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "frmControlFill.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6690
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4380
      TabIndex        =   4
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START"
      Height          =   495
      Left            =   60
      TabIndex        =   3
      Top             =   4200
      Width           =   4275
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3915
      LargeChange     =   300
      Left            =   6420
      Max             =   2000
      SmallChange     =   150
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   3915
      Left            =   60
      ScaleHeight     =   3855
      ScaleWidth      =   6255
      TabIndex        =   0
      Top             =   240
      Width           =   6315
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   0
         ScaleHeight     =   1935
         ScaleWidth      =   6270
         TabIndex        =   5
         Top             =   0
         Width           =   6270
         Begin VB.TextBox TextMul 
            Height          =   285
            Index           =   0
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   9
            Text            =   "frmControlFill.frx":0442
            Top             =   840
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.CheckBox Check 
            Caption         =   "Check"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   8
            Top             =   540
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.TextBox Text 
            Height          =   285
            Index           =   0
            Left            =   60
            TabIndex        =   7
            Text            =   "Text"
            Top             =   300
            Visible         =   0   'False
            Width           =   6135
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Name"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Visible         =   0   'False
            Width           =   420
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Fill in the needed fields. When you are done, click START."
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   4275
   End
End
Attribute VB_Name = "frmControlFill"
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

Public ButIndex As Integer

Public Sub Draw()

'makes the right controls

Dim CurrHeight As Long
CurrHeight = 30


For i = 1 To Commands(ButIndex).NumButtons

    'first make the name
    a = i
    Load Label(a)
    
    Label(a).Top = CurrHeight
    Label(a) = Commands(ButIndex).Buttons(i).ButtonName
    Label(a).Visible = True
            
    CurrHeight = CurrHeight + Label(a).Height + 60
    
    'now the text box/check
    
    If Commands(ButIndex).Buttons(i).Type = 1 Then 'text box
        
        a = i
        Load Text(a)
        
        Text(a).Top = CurrHeight
        Text(a) = Commands(ButIndex).Buttons(i).OptionOn
        Text(a).Visible = True
                
        CurrHeight = CurrHeight + Text(a).Height + 60

    ElseIf Commands(ButIndex).Buttons(i).Type = 2 Then
           
        a = i
        Load Check(a)
        
        Check(a).Top = CurrHeight
        Check(a).Caption = Commands(ButIndex).Buttons(i).ButtonText
        Check(a).Visible = True
                
        CurrHeight = CurrHeight + Check(a).Height + 60
    
    ElseIf Commands(ButIndex).Buttons(i).Type = 4 Then 'multi-line text box
        
        a = i
        Load TextMul(a)
        
        TextMul(a).Top = CurrHeight
        
        TextMul(a).Height = TextMul(a).Height * 4
        TextMul(a).Text = Commands(ButIndex).Buttons(i).OptionOn
        TextMul(a).Visible = True
                
        CurrHeight = CurrHeight + TextMul(a).Height + 60
    
    End If

    CurrHeight = CurrHeight + 180

Next i

Picture2.Height = CurrHeight
VScroll1.Max = CurrHeight + 180 - Picture1.Height

If Commands(ButIndex).NumButtons = 0 Then Command1_Click: Unload Me

End Sub

Private Sub Command1_Click()

'Run
Dim Params() As String
ReDim Params(0 To Commands(ButIndex).NumButtons)

a$ = ""
a$ = a$ + Chr(251)
a$ = a$ + Commands(ButIndex).Name + Chr(250)
For i = 1 To Commands(ButIndex).NumButtons

    
    If Commands(ButIndex).Buttons(i).Type = 1 Then
        a$ = a$ + Text(i) + Chr(250)
    ElseIf Commands(ButIndex).Buttons(i).Type = 2 Then
        If Check(i).Value = 1 Then a$ = a$ + Commands(ButIndex).Buttons(i).OptionOn + Chr(250)
        If Check(i).Value = 0 Then a$ = a$ + Commands(ButIndex).Buttons(i).OptionOff + Chr(250)
    ElseIf Commands(ButIndex).Buttons(i).Type = 4 Then
        a$ = a$ + ReplaceString(TextMul(i), vbCrLf, Chr(10)) + Chr(250)
    End If
    
Next i
a$ = a$ + Chr(251)

SendPacket "SS", a$

Unload Me

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub VScroll1_Change()

Picture2.Top = -VScroll1.Value

End Sub

Private Sub VScroll1_Scroll()

Picture2.Top = -VScroll1.Value

End Sub
