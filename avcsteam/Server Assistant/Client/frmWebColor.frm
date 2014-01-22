VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWebColor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web Colors"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "frmWebColor.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7455
   Begin VB.CommandButton Command3 
      Caption         =   "Set Defaults"
      Height          =   375
      Left            =   2100
      TabIndex        =   34
      Top             =   2700
      Width           =   1155
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   435
      Left            =   1320
      TabIndex        =   33
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   60
      TabIndex        =   32
      Top             =   5040
      Width           =   1215
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00000000&
      Height          =   5175
      Left            =   3360
      ScaleHeight     =   5115
      ScaleWidth      =   4035
      TabIndex        =   10
      Top             =   300
      Width           =   4095
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(TEAM) Fred: Hey whats up?"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   21
         Left            =   60
         TabIndex        =   31
         Top             =   4860
         Width           =   2055
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fred: Hey whats up?"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   20
         Left            =   60
         TabIndex        =   30
         Top             =   4620
         Width           =   1470
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(TEAM) Joey: Hey whats up?"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   19
         Left            =   60
         TabIndex        =   29
         Top             =   4380
         Width           =   2070
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Joey: Hey whats up?"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   18
         Left            =   60
         TabIndex        =   28
         Top             =   4140
         Width           =   1485
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<SERVER> Map time: 5 minutes remaining"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   17
         Left            =   60
         TabIndex        =   27
         Top             =   3900
         Width           =   3015
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SuperFly destroyed Bill's Dispenser"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   16
         Left            =   60
         TabIndex        =   26
         Top             =   3660
         Width           =   2445
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fred built a sentry"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   15
         Left            =   60
         TabIndex        =   25
         Top             =   3420
         Width           =   1245
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fred changed team to RED"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   14
         Left            =   60
         TabIndex        =   24
         Top             =   3180
         Width           =   1950
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fred changed class to Engineer"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   13
         Left            =   60
         TabIndex        =   23
         Top             =   2940
         Width           =   2250
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill changed name to Mitsoki"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   12
         Left            =   60
         TabIndex        =   22
         Top             =   2700
         Width           =   2025
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The goal ""Timer"" was activated"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   11
         Left            =   60
         TabIndex        =   21
         Top             =   2460
         Width           =   2250
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill activated the goal ""Blue Flag"""
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   10
         Left            =   60
         TabIndex        =   20
         Top             =   2220
         Width           =   2370
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SuperFly has left the game"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   9
         Left            =   60
         TabIndex        =   19
         Top             =   1980
         Width           =   1875
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Joey has joined the game"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   8
         Left            =   60
         TabIndex        =   18
         Top             =   1740
         Width           =   1800
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Joey killed Bob with shotgun"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   7
         Left            =   60
         TabIndex        =   17
         Top             =   1500
         Width           =   2010
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<ADMIN> Server going down."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   6
         Left            =   60
         TabIndex        =   16
         Top             =   1260
         Width           =   2130
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<ADMIN Bill> Server going down."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   5
         Left            =   60
         TabIndex        =   15
         Top             =   1020
         Width           =   2370
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(TEAM) Fred: Hey whats up?"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   14
         Top             =   780
         Width           =   2055
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fred: Hey whats up?"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   13
         Top             =   540
         Width           =   1470
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(TEAM) Joey: Hey whats up?"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   12
         Top             =   300
         Width           =   2070
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Joey: Hey whats up?"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   11
         Top             =   60
         Width           =   1485
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
      Height          =   615
      Left            =   60
      ScaleHeight     =   555
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
Attribute VB_Name = "frmWebColor"
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
    Sample(a).ForeColor = RGB(r, g, b)
    
    Web.Colors(a).r = r
    Web.Colors(a).g = g
    Web.Colors(a).b = b
    
    
End If

End Sub

Sub ShowColor()

a = Combo1.ListIndex

If a > -1 Then
    a = a + 1

    r = Web.Colors(a).r
    g = Web.Colors(a).g
    b = Web.Colors(a).b
    
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

For i = 1 To 21

    r = Web.Colors(i).r
    g = Web.Colors(i).g
    b = Web.Colors(i).b

    Sample(i).ForeColor = RGB(r, g, b)
Next i

End Sub

Private Sub Combo1_Click()
ShowColor

End Sub

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command2_Click()
PackageWebColors

Unload Me
End Sub

Private Sub Command3_Click()

a = MessBox("Are you sure you want to reload defaults?", vbYesNo + vbQuestion, "Reload Default Colors")

If a = vbNo Then Exit Sub


'Open App.Path + "\webcol.txt" For Append As #1
'
'    For I = 1 To 21
'
'        R = Web.Colors(I).R
'        G = Web.Colors(I).G
'        b = Web.Colors(I).b
'
'        Print #1, "Web.Colors(" + Ts(I) + ").R = " + Ts(R)
'        Print #1, "Web.Colors(" + Ts(I) + ").G = " + Ts(G)
'        Print #1, "Web.Colors(" + Ts(I) + ").B = " + Ts(b)
'        Print #1, ""
'
'    Next I
'
'
'Close #1

Web.Colors(1).r = 116
Web.Colors(1).g = 117
Web.Colors(1).b = 255

Web.Colors(2).r = 0
Web.Colors(2).g = 153
Web.Colors(2).b = 255

Web.Colors(3).r = 255
Web.Colors(3).g = 94
Web.Colors(3).b = 105

Web.Colors(4).r = 255
Web.Colors(4).g = 105
Web.Colors(4).b = 143

Web.Colors(5).r = 213
Web.Colors(5).g = 180
Web.Colors(5).b = 255

Web.Colors(6).r = 180
Web.Colors(6).g = 119
Web.Colors(6).b = 211

Web.Colors(7).r = 0
Web.Colors(7).g = 97
Web.Colors(7).b = 151

Web.Colors(8).r = 131
Web.Colors(8).g = 121
Web.Colors(8).b = 0

Web.Colors(9).r = 131
Web.Colors(9).g = 117
Web.Colors(9).b = 0

Web.Colors(10).r = 204
Web.Colors(10).g = 160
Web.Colors(10).b = 0

Web.Colors(11).r = 182
Web.Colors(11).g = 145
Web.Colors(11).b = 0

Web.Colors(12).r = 139
Web.Colors(12).g = 88
Web.Colors(12).b = 158

Web.Colors(13).r = 138
Web.Colors(13).g = 71
Web.Colors(13).b = 128

Web.Colors(14).r = 126
Web.Colors(14).g = 129
Web.Colors(14).b = 129

Web.Colors(15).r = 65
Web.Colors(15).g = 189
Web.Colors(15).b = 224

Web.Colors(16).r = 255
Web.Colors(16).g = 146
Web.Colors(16).b = 0

Web.Colors(17).r = 255
Web.Colors(17).g = 255
Web.Colors(17).b = 255

Web.Colors(18).r = 255
Web.Colors(18).g = 255
Web.Colors(18).b = 0

Web.Colors(19).r = 223
Web.Colors(19).g = 216
Web.Colors(19).b = 0

Web.Colors(20).r = 0
Web.Colors(20).g = 255
Web.Colors(20).b = 0

Web.Colors(21).r = 105
Web.Colors(21).g = 255
Web.Colors(21).b = 194
UpdateSample
End Sub

Private Sub Form_Load()
UpdateSample

Combo1.AddItem "Blue Speech"
Combo1.AddItem "Blue TEAM Speech"
Combo1.AddItem "Red Speech"
Combo1.AddItem "Red TEAM Speech"
Combo1.AddItem "Named ADMIN Speech"
Combo1.AddItem "Unnamed ADMIN Speech"
Combo1.AddItem "Kills"
Combo1.AddItem "Joins"
Combo1.AddItem "Leaves"
Combo1.AddItem "Goals (named)"
Combo1.AddItem "Goals (unnamed)"
Combo1.AddItem "Name Changes"
Combo1.AddItem "Class Changes"
Combo1.AddItem "Team Changes"
Combo1.AddItem "Building Builds"
Combo1.AddItem "Building Destroys"
Combo1.AddItem "Server Speech"
Combo1.AddItem "Yellow Speech"
Combo1.AddItem "Yellow TEAM Speech"
Combo1.AddItem "Green Speech"
Combo1.AddItem "Green TEAM Speech"

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

