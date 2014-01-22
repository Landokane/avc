VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Interface Colors"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "frmColors.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   7455
   Begin VB.CommandButton Command3 
      Caption         =   "Set Defaults"
      Height          =   375
      Left            =   3360
      TabIndex        =   17
      Top             =   2400
      Width           =   1155
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   2400
      Width           =   2895
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00000000&
      Height          =   2055
      Left            =   3360
      ScaleHeight     =   1995
      ScaleWidth      =   4035
      TabIndex        =   10
      Top             =   300
      Width           =   4095
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<TELL Fred>"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   8
         Left            =   60
         TabIndex        =   28
         Top             =   1740
         Width           =   930
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jackie Whats up"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   16
         Left            =   1080
         TabIndex        =   27
         Top             =   1740
         Width           =   1200
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server Assistant"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   15
         Left            =   1080
         TabIndex        =   26
         Top             =   1500
         Width           =   1140
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<MESSAGE>"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   7
         Left            =   60
         TabIndex        =   25
         Top             =   1500
         Width           =   960
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Did you hear about this new tool?"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   11
         Left            =   360
         TabIndex        =   24
         Top             =   540
         Width           =   2370
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hey whats up?"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   9
         Left            =   480
         TabIndex        =   23
         Top             =   60
         Width           =   1065
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nothing much..."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   10
         Left            =   480
         TabIndex        =   22
         Top             =   300
         Width           =   1125
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server going down."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   13
         Left            =   1080
         TabIndex        =   21
         Top             =   1020
         Width           =   1380
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Map time: 5 minutes remaining"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   14
         Left            =   960
         TabIndex        =   20
         Top             =   1260
         Width           =   2130
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Yeah its called Server Assistant or something..."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   12
         Left            =   540
         TabIndex        =   19
         Top             =   780
         Width           =   3300
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill:"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   18
         Top             =   540
         Width           =   240
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frank:"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   15
         Top             =   780
         Width           =   450
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<SERVER>"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   6
         Left            =   60
         TabIndex        =   14
         Top             =   1260
         Width           =   840
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<ADMIN Bill>"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   5
         Left            =   60
         TabIndex        =   13
         Top             =   1020
         Width           =   945
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fred:"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   12
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Joey:"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   11
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   435
      Left            =   2820
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   8
      Top             =   1500
      Width           =   435
   End
   Begin VB.PictureBox Picture3 
      Height          =   435
      Left            =   2820
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   7
      Top             =   960
      Width           =   435
   End
   Begin VB.PictureBox Picture2 
      Height          =   435
      Left            =   2820
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   6
      Top             =   420
      Width           =   435
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   435
      Left            =   60
      TabIndex        =   3
      Top             =   420
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   767
      _Version        =   393216
      Max             =   255
      TickFrequency   =   17
   End
   Begin VB.PictureBox Picture1 
      Height          =   675
      Left            =   60
      ScaleHeight     =   615
      ScaleWidth      =   3135
      TabIndex        =   2
      Top             =   2040
      Width           =   3195
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   660
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   2595
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   435
      Left            =   60
      TabIndex        =   4
      Top             =   960
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   767
      _Version        =   393216
      Max             =   255
      TickFrequency   =   17
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   435
      Left            =   60
      TabIndex        =   5
      Top             =   1500
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   767
      _Version        =   393216
      Max             =   255
      TickFrequency   =   17
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sample"
      Height          =   195
      Left            =   3360
      TabIndex        =   9
      Top             =   60
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Color"
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   660
   End
End
Attribute VB_Name = "frmColors"
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


Sub UpdateColor()

a = Combo1.ListIndex

r = Slider1.Value
g = Slider2.Value
b = Slider3.Value

Picture1.BackColor = RGB(r, g, b)

Picture2.BackColor = RGB(r, 0, 0)
Picture3.BackColor = RGB(0, g, 0)
Picture4.BackColor = RGB(0, 0, b)

If a > -1 Then
    a = a + 1
    
    If (a > 1 And a < 9) Or a = 10 Then
        
        If a <> 10 Then Sample(a - 1).ForeColor = RGB(r, g, b)
        If a = 10 Then Sample(a - 2).ForeColor = RGB(r, g, b)
    ElseIf a = 1 Then
        For i = 9 To 16
            Sample(i).ForeColor = RGB(r, g, b)
        Next i
    ElseIf a = 9 Then
        Picture5.BackColor = RGB(RichColors(9).r, RichColors(9).g, RichColors(9).b)
    End If
    RichColors(a).r = r
    RichColors(a).g = g
    RichColors(a).b = b
End If

End Sub

Sub ShowColor()

a = Combo1.ListIndex

If a > -1 Then
    a = a + 1

    r = RichColors(a).r
    g = RichColors(a).g
    b = RichColors(a).b
    
    Picture1.BackColor = RGB(r, g, b)
    
    Picture2.BackColor = RGB(r, 0, 0)
    Picture3.BackColor = RGB(0, g, 0)
    Picture4.BackColor = RGB(0, 0, b)
    
    Slider1.Value = r
    Slider2.Value = g
    Slider3.Value = b

End If

End Sub

Sub UpdateSample()

For i = 1 To 7

    r = RichColors(i + 1).r
    g = RichColors(i + 1).g
    b = RichColors(i + 1).b
    Sample(i).ForeColor = RGB(r, g, b)
Next i

i = 9
    
r = RichColors(i + 1).r
g = RichColors(i + 1).g
b = RichColors(i + 1).b
Sample(8).ForeColor = RGB(r, g, b)


For i = 9 To 16

    r = RichColors(1).r
    g = RichColors(1).g
    b = RichColors(1).b
    Sample(i).ForeColor = RGB(r, g, b)
Next i

Picture5.BackColor = RGB(RichColors(9).r, RichColors(9).g, RichColors(9).b)

End Sub

Private Sub Combo1_Click()
ShowColor

End Sub

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command2_Click()

UpdateColors


Unload Me
End Sub

Private Sub Command3_Click()

a = MessBox("Are you sure you want to reload defaults?", vbYesNo + vbQuestion, "Reload Default Colors")

If a = vbNo Then Exit Sub

'normal
RichColors(1).r = 255
RichColors(1).g = 255
RichColors(1).b = 255

'blue
RichColors(2).r = 116
RichColors(2).g = 117
RichColors(2).b = 255

'red
RichColors(3).r = 255
RichColors(3).g = 94
RichColors(3).b = 105

'yellow
RichColors(4).r = 255
RichColors(4).g = 255
RichColors(4).b = 0

'green
RichColors(5).r = 0
RichColors(5).g = 255
RichColors(5).b = 0

'admin
RichColors(6).r = 255
RichColors(6).g = 0
RichColors(6).b = 255

'server
RichColors(7).r = 145
RichColors(7).g = 145
RichColors(7).b = 145

'message
RichColors(8).r = 200
RichColors(8).g = 177
RichColors(8).b = 100

RichColors(9).r = 0
RichColors(9).g = 0
RichColors(9).b = 0

RichColors(10).r = 214
RichColors(10).g = 99
RichColors(10).b = 119

UpdateSample
End Sub

Private Sub Form_Load()
UpdateSample

Combo1.AddItem "Normal Text"
Combo1.AddItem "Blue Speech"
Combo1.AddItem "Red Speech"
Combo1.AddItem "Yellow Speech"
Combo1.AddItem "Green Speech"
Combo1.AddItem "ADMIN Speech"
Combo1.AddItem "Server Speech"
Combo1.AddItem "Messages"
Combo1.AddItem "Background"
Combo1.AddItem "TELL Speech"
Combo1.ListIndex = 0

End Sub

Private Sub Slider1_Click()
UpdateColor
End Sub

Private Sub Slider1_Scroll()
UpdateColor
End Sub

Private Sub Slider2_Click()
UpdateColor
End Sub

Private Sub Slider2_Scroll()
UpdateColor
End Sub

Private Sub Slider3_Click()
UpdateColor
End Sub

Private Sub Slider3_Scroll()
UpdateColor
End Sub

