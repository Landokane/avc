VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMap 
   Caption         =   "Map"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10155
   Icon            =   "frmmap.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   10155
   Begin VB.CheckBox Check3 
      Height          =   135
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command5"
      Height          =   615
      Left            =   8820
      TabIndex        =   14
      Top             =   5040
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   3495
      Left            =   8340
      Max             =   360
      Min             =   -360
      TabIndex        =   13
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   3495
      Left            =   8040
      Max             =   360
      Min             =   -360
      TabIndex        =   12
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3495
      Left            =   7740
      Max             =   360
      Min             =   -360
      TabIndex        =   11
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic3D 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7695
      Left            =   9480
      ScaleHeight     =   513
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   501
      TabIndex        =   10
      Top             =   6900
      Visible         =   0   'False
      Width           =   7515
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   615
      Left            =   8820
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3180
      Left            =   0
      ScaleHeight     =   212
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   212
      TabIndex        =   0
      Top             =   0
      Width           =   3180
      Begin VB.Image Image1 
         DragMode        =   1  'Automatic
         Height          =   240
         Index           =   0
         Left            =   120
         Picture         =   "frmmap.frx":08CA
         Top             =   120
         Width           =   240
      End
      Begin VB.Image imgTele 
         Height          =   240
         Index           =   0
         Left            =   420
         Picture         =   "frmmap.frx":0E54
         Top             =   120
         Width           =   240
      End
      Begin VB.Image RotImg 
         Height          =   480
         Left            =   780
         Picture         =   "frmmap.frx":13DE
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   9000
      Top             =   3840
   End
   Begin MSComctlLib.ImageList RotImgSet 
      Left            =   7500
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":1CA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":2584
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":2E60
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":373C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":4018
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":45B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":4B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":50EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Map Settings"
      Height          =   3735
      Left            =   7740
      TabIndex        =   1
      Top             =   0
      Width           =   2355
      Begin VB.CheckBox Check4 
         Caption         =   "Team Coloured Map"
         Height          =   255
         Left            =   60
         TabIndex        =   16
         Top             =   3420
         Value           =   1  'Checked
         Width           =   2235
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Hide Arrows"
         Height          =   315
         Left            =   60
         TabIndex        =   8
         Top             =   1260
         Width           =   2235
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Hide Players"
         Height          =   255
         Left            =   60
         TabIndex        =   7
         Top             =   3180
         Width           =   2235
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Hide Locations"
         Height          =   255
         Left            =   60
         TabIndex        =   6
         Top             =   2940
         Width           =   2235
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Refresh Locations"
         Height          =   615
         Left            =   1200
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Redraw"
         Height          =   315
         Left            =   60
         TabIndex        =   3
         Top             =   240
         Width           =   2235
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Refresh Map"
         Height          =   615
         Left            =   60
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1275
         Left            =   60
         TabIndex        =   4
         Top             =   1620
         Width           =   2235
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8880
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":5688
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":5C24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":61C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":675C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":6CF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":7294
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":7830
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":7CF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":81B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":8678
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":8B38
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":8FFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmap.frx":94BC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMap"
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

Dim NumImage As Integer
Dim ImageLoad(0 To 40) As Boolean
Dim NumTeleImage As Integer
Dim ImageTeleLoad(0 To 400) As Boolean
Dim ShowTele As Boolean
Dim ShowPlay As Boolean
Dim RotImgNum As Integer
Public RotPlayerNum As Integer
Dim SelTele As Integer
Public OldWindowProc As Long



Private Sub Check1_Click()

If Check1 = 0 Then ShowTele = True
If Check1 = 1 Then ShowTele = False
Update2

End Sub

Private Sub Check2_Click()
If Check2 = 0 Then ShowPlay = True
If Check2 = 1 Then ShowPlay = False
Update2

End Sub

Private Sub Check3_Click()

Form_Resize
If Check3.Value = 1 Then Frame1.Visible = False
If Check3.Value = 0 Then Frame1.Visible = True



End Sub

Private Sub Check4_Click()
Draw

End Sub

Private Sub Command1_Click()
Draw


End Sub

Private Sub Command2_Click()

SendPacket "MD", ""


End Sub

Private Sub Command3_Click()

SendPacket "TE", ""

End Sub

Private Sub Command4_Click()
RotPlayerNum = 0
RotImg.Visible = False

End Sub

Private Sub Command5_Click()
pic3D.Visible = True

Draw3D

End Sub


Private Function CalcAng(InPt As typPlayerPos, Phi, Rho, Theta) As typPlayerPos
'Takes X, Y, Z as input, then calculates the X,Y,Z output after angles have been implemented
    X = InPt.X
    Y = InPt.Y
    Z = InPt.Z
        
    X1 = X - center
    Y1 = Y - center
    z1 = Z
    
    'find existing angle
    If Y1 <> 0 Then m = z1 / Y1
    a2 = Atn(m) * (180 / 3.14159)
    
    If Y1 >= 0 Then a2 = a2 + 180
    
    a2 = a2 + Rho
    Do Until a2 <= 360: a2 = a2 - 360: Loop
    a1 = a2 * (3.14159 / 180)
    
    y3 = Cos(a1) * Sqr(z1 ^ 2 + Y1 ^ 2)  ' Cos * Distance from center
    Z2 = Sin(a1) * Sqr(z1 ^ 2 + Y1 ^ 2)  ' Sin * Distance from center
        
    'find existing angle
    If X1 <> 0 Then m = y3 / X1
    a2 = Atn(m) * (180 / 3.14159)
    
    If X1 >= 0 Then a2 = a2 + 180
    
    a2 = a2 + Phi
    Do Until a2 <= 360: a2 = a2 - 360: Loop
    a1 = a2 * (3.14159 / 180)
    
    X2 = Cos(a1) * Sqr(X1 ^ 2 + y3 ^ 2)  ' Cos * Distance from center
    Y2 = Sin(a1) * Sqr(X1 ^ 2 + y3 ^ 2)  ' Sin * Distance from center
    
    CalcAng.X = X2
    CalcAng.Y = Y2
    CalcAng.Z = Z2
    
    
End Function

Sub Draw3D()
pic3D.Cls

Dim OutPt As typPlayerPos
Dim OldPt As typPlayerPos


For X1 = 0 To 64
    For Y1 = 0 To 64
        ' MapArray(x1, y1) <> 0 Then
            
            cc = (MapArray(X1, Y1) + 4096)
            cc = cc - low
            If (high - low) <> 0 Then cc = cc / (high - low)
            
            cc3 = Int((cc * (rngtop - rngbot)) + rngbot)
            If cc3 < 0 Then cc3 = 0
            If cc3 > 255 Then cc3 = 255
            
            c = RGB(cc3, cc3, cc3)
            X2 = X1 * mp
            Y2 = (Picture1.Height / Screen.TwipsPerPixelY) - (Y1 * mp)
            x3 = (X1 * mp) + mp
            y3 = (Picture1.Height / Screen.TwipsPerPixelY) - ((Y1 * mp) + mp)
            Z = MapArray(X1, Y1) / 128
            
            OutPt.X = (X1 - 32) * 4
            OutPt.Y = (Y1 - 32) * 4
            OutPt.Z = Z * 4
            OldPt.X = OutPt.X
            OldPt.Y = OutPt.Y
            OldPt.Z = OutPt.Z
            
            
            OutPt = CalcAng(OutPt, VScroll1, VScroll2, VScroll3)
            'If coefx = 0 And coefy = 0 And coefz = 0 Then
                
                'coefx = OldPt.x / OutPt.x
                'coefy = OldPt.y / OutPt.y
                'if OutPt.Z <> 0c oefz = OldPt.Z / OutPt.Z
            'Else
             '   OutPt.x = OutPt.x * coefx
             '   OutPt.y = OutPt.y * coefy
             '   OutPt.Z = OutPt.Z * coefz
            'End If
            
            OutPt = Calc3D(OutPt)
            
            pic3D.PSet (OutPt.X, OutPt.Y), RGB(255, 255, 255)
            'pic3D.Line (0, 0)-(OutPt.X, OutPt.Y), RGB(255, 255, 0)
            
            
        'End If
    Next Y1
    DoEvents
Next X1








End Sub


Private Function CalcRot(InPt As typPlayerPos, Phi As Long, Rho As Long, Theta As Long) As typPlayerPos


Phi2 = Phi * (3.141592654 / 180)
rho2 = Rho * (3.141592654 / 180)
theta2 = Theta * (3.141592654 / 180)


CalcRot.X = -InPt.X * Cos(Phi2) * Cos(rho2) - InPt.Y * Cos(Phi2) * Sin(rho2) + InPt.Z * Sin(Phi2)
CalcRot.Y = InPt.X * Sin(theta2) * Sin(Phi2) * Cos(rho2) + Cos(theta2) * Sin(rho2) - InPt.Y * Sin(theta2) * Sin(Phi2) * Sin(rho2) + Cos(theta2) * Cos(rho2) - InPt.Z * Sin(theta2) * Cos(Phi2)
CalcRot.Z = -InPt.X * Cos(theta2) * Sin(Phi2) * Cos(rho2) + Sin(theta2) * Sin(rho2) + InPt.Y * Cos(theta2) * Sin(Phi2) * Sin(rho2) + Sin(theta2) * Cos(rho2) + InPt.Z * Cos(theta2) * Cos(Phi2)

End Function

Private Function Calc3D(InPt As typPlayerPos) As typPlayerPos

screenwidth = pic3D.Width / Screen.TwipsPerPixelX
ScreenHeight = pic3D.Height / Screen.TwipsPerPixelY

m = 256
X2 = InPt.X
Z2 = InPt.Z
Y2 = InPt.Y

If m - Y2 <> 0 Then Sx = (m * X2 / (m - Y2)) + screenwidth
If m - Y2 <> 0 Then Sy = ScreenHeight - (m * Z2 / (m - Y2))
Sy = Sy - 200
Sx = Sx - 200

Calc3D.X = Sx
Calc3D.Y = Sy

End Function


Private Sub Command6_Click()
pic3D.Visible = False

End Sub

Private Sub Form_Load()


SendPacket "TE", ""
SendPacket "MD", ""
ShowTele = True
ShowPlay = True

Check4 = Val(GetSetting("Server Assistant Client", "Map", "TeamColours", 1))

'Draw
Update2
ShowMap = True

For i = 0 To 40
    ImageLoad(i) = False
Next i

On Error Resume Next
nm$ = Me.Name
winash = Val(GetSetting("Server Assistant Client", "Window", nm$ + "winash", -1))
If winash <> -1 Then Me.Show
winmd = Val(GetSetting("Server Assistant Client", "Window", nm$ + "winmd", 3))
If winmd <> 3 Then Me.WindowState = winmd
If Me.WindowState = 0 Then
    winh = GetSetting("Server Assistant Client", "Window", nm$ + "winh", -1)
    wint = GetSetting("Server Assistant Client", "Window", nm$ + "wint", -1)
    winl = GetSetting("Server Assistant Client", "Window", nm$ + "winl", -1)
    winw = GetSetting("Server Assistant Client", "Window", nm$ + "winw", -1)
    
    If winh <> -1 Then Me.Height = winh
    If wint <> -1 Then Me.Top = wint
    If winl <> -1 Then Me.Left = winl
    If winw <> -1 Then Me.Width = winw
End If


Check3 = GetSetting("Server Assistant Client", "Window", "mapsmall", 0)

AddForm True, 267, 133, 0, 0, Me

End Sub
Public Sub Functions(Index As Integer)

    
If SelTele > 0 Then
    b$ = Tele(SelTele).Name
    
    If Index = 0 Then
        'delete
        SendPacket "RC", "telekill " + b$
        SendPacket "TE", ""

    End If
    If Index = 1 Then
        'rename
        
        c$ = InBox("Rename location " + b$ + " to:", "Rename Location", b$)
        If c$ = "" Then Exit Sub
                
        SendPacket "RC", "telename " + b$ + " " + c$
        SendPacket "TE", ""
    
    End If
End If

End Sub

Public Sub Update2()

'Unload Me
On Error Resume Next

If NumImage > NumPlayers Then

    For i = NumPlayers + 1 To NumImage
        
        If ImageLoad(i) = True Then Unload Image1(i)
        ImageLoad(i) = False
    Next i

End If

If NumTeleImage > NumTele Then

    For i = NumTele + 1 To NumTeleImage
        If ImageTeleLoad(i) = True Then Unload imgTele(i)
        ImageTeleLoad(i) = False
    Next i
End If

Image1(0).Visible = False

w = (Picture1.Width / Screen.TwipsPerPixelX)
w2 = Int(w / 2)

mp = Int(8192 / w)

For i = 1 To NumPlayers
    
    c = RGB(RichColors(Players(i).Team + 1).r, RichColors(Players(i).Team + 1).g, RichColors(Players(i).Team + 1).b)
                
    X = Players(i).pos.X
    Y = Players(i).pos.Y
    Z = Players(i).pos.Z
    
    
    If ImageLoad(i) = False Then
        
        'Unload Image1(i)
        Load Image1(i)
        ImageLoad(i) = True
        If NumImage < i Then NumImage = i
    End If
    
    
    
    If Players(i).UserID = RotPlayerNum Then
    
        If Picture1.Width / Screen.TwipsPerPixelX < 300 Then
            RotImg.Picture = RotImgSet.ListImages(5).Picture
        Else
            RotImg.Picture = RotImgSet.ListImages(1).Picture
        End If
        RotImg.Top = (Picture1.Height / Screen.TwipsPerPixelY) - Int((Y / mp) + w2) - Int(RotImg.Height / 2)
        RotImg.Left = Int((X / mp) + w2) - Int(RotImg.Width / 2)
        RotImg.Visible = True
        Image1(i).ZOrder
    End If
    
    t = Players(i).Team
    If t = 0 Then t = 5
    
    
    
    If Picture1.Width / Screen.TwipsPerPixelX < 300 Then t = t + 6
    
    Image1(i).Picture = ImageList1.ListImages(t).Picture
    
    Image1(i).Top = (Picture1.Height / Screen.TwipsPerPixelY) - Int((Y / mp) + w2) - Int(Image1(i).Height / 2)
    Image1(i).Left = Int((X / mp) + w2) - Int(Image1(i).Width / 2)
    Image1(i).Tag = Ts(Players(i).UserID)
    Image1(i).Visible = ShowPlay

    If Players(i).Team = 6 Then Image1(i).Visible = False

    If X = 0 And Y = 0 And Z = 0 Then Image1(i).Visible = False

Next i



imgTele(0).Visible = False

For i = 1 To NumTele
    
    X = Tele(i).X
    Y = Tele(i).Y
    Z = Tele(i).Z
    
    If ImageTeleLoad(i) = False Then
    
        Load imgTele(i)
        ImageTeleLoad(i) = True
        
        
        If NumTeleImage < i Then NumTeleImage = i
    End If
    
    If Picture1.Width / Screen.TwipsPerPixelX < 300 Then
        imgTele(i).Picture = ImageList1.ListImages(12).Picture
    Else
        imgTele(i).Picture = ImageList1.ListImages(13).Picture
    End If
    imgTele(i).Top = (Picture1.Height / Screen.TwipsPerPixelY) - Int((Y / mp) + w2) - Int(imgTele(i).Height / 2)
    imgTele(i).Left = Int((X / mp) + w2) - Int(imgTele(i).Width / 2)
    imgTele(i).Tag = Ts(i)
    imgTele(i).Visible = ShowTele
    imgTele(i).ZOrder 1

Next i





End Sub

Function GetMapArrayTeam(z1) As Integer

t = 0
If z1 >= 4097 And z1 <= 12288 Then t1 = 1
If z1 >= 12289 And z1 <= 20480 Then t1 = 2
If z1 >= 20481 Then t1 = 5
If z1 <= -4097 And z1 >= -12288 Then t1 = 3
If z1 <= -12289 Then t1 = 4

GetMapArrayTeam = t1

End Function

Public Sub Draw()

Picture1.Cls
Sc = (8192 / Picture1.Width) * ZoomLevel

'Draw the array

w = (Picture1.Width / Screen.TwipsPerPixelX)
mp = w / 64

'Get highest

high = -4096
low = 4096


For X1 = 1 To 63
    For Y1 = 1 To 63
        
        z1 = MapArray(X1, Y1)
        tt1 = GetMapArrayTeam(z1)
        
        If tt1 = 1 Then z1 = z1 - 8192
        If tt1 = 2 Then z1 = z1 - 16384
        If tt1 = 3 Then z1 = z1 + 8192
        If tt1 = 4 Then z1 = z1 + 16384
        If tt1 = 5 Then z1 = z1 - 24576
            
        cs2 = z1
        
        If cs2 > high And cs2 <= 4096 And cs2 <> 0 Then high = cs2
        If cs2 < low And cs2 >= -4096 And cs2 <> 0 Then low = cs2


    Next Y1
Next X1

'range
rngtop = 180
rngbot = 100

low = low + 4096
high = high + 4096

If high - low = 0 Then high = high + 1


For X1 = 0 To 64
    For Y1 = 0 To 64
        If MapArray(X1, Y1) <> 0 Then
            
            z1 = MapArray(X1, Y1)
            tt1 = GetMapArrayTeam(z1)
            
            If tt1 = 1 Then z1 = z1 - 8192
            If tt1 = 2 Then z1 = z1 - 16384
            If tt1 = 3 Then z1 = z1 + 8192
            If tt1 = 4 Then z1 = z1 + 16384
            If tt1 = 5 Then z1 = z1 - 24576

'Map Data Format:
' -4096   to   4096 -> more than one team / old format
'  4097   to  12288 -> blue team   (norm: -8192)
' 12289   to  20480 -> red team    (norm: -16384)
' -4097   to -12288 -> yellow team (norm: +8192)
'-12289   to -20480 -> green team  (norm: +16384)


            cc = (z1 + 4096)
            cc = cc - low
            
            cc = cc / (high - low)
            
            cc3 = Int((cc * (rngtop - rngbot)) + rngbot)
            
            
            If cc3 < 0 Then cc3 = 0
            If cc3 > 255 Then cc3 = 255
            
            c = RGB(cc3, cc3, cc3)
            
            'If tt1 = 0 And Check4 Then c = RGB(cc3 + 50, 25, cc3 + 50)
            If Check4 And tt1 > 0 Then
                
                If tt1 = 1 Then c = RGB(0, 0, cc3)
                If tt1 = 2 Then c = RGB(cc3, 0, 0)
                If tt1 = 3 Then c = RGB(cc3, cc3, 0)
                If tt1 = 4 Then c = RGB(0, cc3, 0)
                
            
            End If
            
            Picture1.Line (X1 * mp, (Picture1.Height / Screen.TwipsPerPixelY) - (Y1 * mp))-((X1 * mp) + mp, (Picture1.Height / Screen.TwipsPerPixelY) - ((Y1 * mp) + mp)), c, BF
        End If
    Next Y1
    DoEvents
Next X1

End Sub

Private Sub Form_Resize()
'Exit Sub

w = Me.Width
h = Me.Height

If Me.WindowState = 0 Then
    If Check3.Value = 0 And w < 4000 Then Me.Width = 4000: w = 4000
    If Check3.Value = 1 And w < 2000 Then Me.Width = 2000: w = 2000
    If h < 2000 Then Me.Height = 2000: h = 2000
End If

If Me.WindowState <> 0 Then Exit Sub
 
If Check3.Value = 0 Then w = w - Frame1.Width - 120
h = h - 520

a = w
If h < w Then
    a = h
End If

Picture1.Width = a
Picture1.Height = a

Frame1.Left = Me.Width - Frame1.Width - 60

Draw
Update2



End Sub

Private Sub Form_Unload(Cancel As Integer)
ShowMap = False

On Error Resume Next
nm$ = Me.Name
SaveSetting "Server Assistant Client", "Window", nm$ + "winmd", Me.WindowState
SaveSetting "Server Assistant Client", "Window", nm$ + "winh", Me.Height
SaveSetting "Server Assistant Client", "Window", nm$ + "wint", Me.Top
SaveSetting "Server Assistant Client", "Window", nm$ + "winl", Me.Left
SaveSetting "Server Assistant Client", "Window", nm$ + "winw", Me.Width
SaveSetting "Server Assistant Client", "Window", "mapsmall", Check3
SaveSetting "Server Assistant Client", "Map", "TeamColours", Check4

End Sub

Private Sub Image1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

If Source.Index <> Index Then
    
    c1$ = Trim(Image1(Index).Tag)
    b = FindPlayer(c1$)
    If b = 0 Then Exit Sub
    
    c$ = Trim(Source.Tag)
    a = FindPlayer(c$)
    If a = 0 Then Exit Sub

    X1 = Players(b).pos.X
    Y1 = Players(b).pos.Y
    z1 = Players(b).pos.Z
    
    'If j > 0 Then
    SendPacket "RC", "teleportto " + c$ + " " + Ts(X1) + " " + Ts(Y1) + " " + Ts(z1)
    'End If
    
    Players(a).pos.X = X1
    Players(a).pos.Y = Y1
    Players(a).pos.Z = z1
    
'    Source.Top = Image1(Index).Top
'    Source.Left = Image1(Index).Left
    
    Update2
    


End If



End Sub

Private Sub Image1_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)


If Source.Index <> Index Then
    
    c$ = Trim(Image1(Index).Tag)
    b = FindPlayer(c$)
    If b = 0 Then Exit Sub
    
    c$ = Trim(Source.Tag)
    a = FindPlayer(c$)
    If a = 0 Then Exit Sub
    
    Label1 = "Teleport: " + vbCrLf + Players(a).Name + vbCrLf + "to same loc as player: " + vbCrLf + Players(b).Name + vbCrLf + "at: " + Ts(Players(b).pos.X) + ", " + Ts(Players(b).pos.Y) + ", " + Ts(Players(b).pos.Z)
    If Check3.Value = 1 Then Me.Caption = Players(b).Name

End If


End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

a$ = Val(Image1(Index).Tag)
b = FindPlayer(a$)

If b = 0 Then Exit Sub

Label1 = Players(b).Name + vbCrLf + Ts(Players(b).pos.X) + ", " + Ts(Players(b).pos.Y) + ", " + Ts(Players(b).pos.Z)
If Check3.Value = 1 Then Me.Caption = Players(b).Name

'Label1.AutoSize = True


End Sub

Private Sub imgTele_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

i = Source.Index
u$ = Source.Tag

For i = 1 To NumPlayers
    If Val(u$) = Players(i).UserID Then j = i: Exit For
Next i

a = Val(imgTele(Index).Tag)

X1 = Tele(a).X
Y1 = Tele(a).Y
z1 = Tele(a).Z

If j > 0 Then
    SendPacket "RC", "teleport " + u$ + " " + Tele(a).Name
End If

Players(j).pos.X = X1
Players(j).pos.Y = Y1
Players(j).pos.Z = z1

'Source.Top = imgTele(Index).Top
'    Source.Left = imgTele(Index).Left

Update2

End Sub

Private Sub imgTele_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)

c$ = Trim(Source.Tag)
b = FindPlayer(c$)
If b = 0 Then Exit Sub
a = Val(imgTele(Index).Tag)
Label1 = "Teleport: " + vbCrLf + Players(b).Name + vbCrLf + "to location: " + vbCrLf + Tele(a).Name + vbCrLf + "at: " + Ts(Tele(a).X) + ", " + Ts(Tele(a).Y) + ", " + Ts(Tele(a).Z)
If Check3.Value = 1 Then Me.Caption = Tele(a).Name

End Sub

Private Sub imgTele_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

a = Val(imgTele(Index).Tag)
Label1 = "Location:" + vbCrLf + Tele(a).Name + vbCrLf + Ts(Tele(a).X) + ", " + Ts(Tele(a).Y) + ", " + Ts(Tele(a).Z)
If Check3.Value = 1 Then Me.Caption = Tele(a).Name

End Sub

Private Sub imgTele_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

SelTele = Val(imgTele(Index).Tag)
MDIForm1.PopupMenu MDIForm1.mnuTeleporters

End Sub

Private Sub Picture1_DblClick()

If Check3.Value = 1 Then
    Check3.Value = 0
Else
    Check3.Value = 1
End If

End Sub

Private Sub Picture1_DragDrop(Source As Control, X As Single, Y As Single)

w = (Picture1.Width / Screen.TwipsPerPixelX)
mp = Int(8192 / w)
mp2 = w / 64


i = Source.Index
u$ = Source.Tag

For i = 1 To NumPlayers
    If Val(u$) = Players(i).UserID Then j = i: Exit For
Next i

X1 = (X - Int(w / 2)) * mp
Y1 = (((Picture1.Height / Screen.TwipsPerPixelY) - Y) - Int(w / 2)) * mp
z1 = MapArray(Int(X / mp2), Int(((Picture1.Height / Screen.TwipsPerPixelY) / mp2) - (Y / mp2)))
z1 = ConvertPoint(z1)

If j > 0 Then
    SendPacket "RC", "teleportto " + u$ + " " + Ts(X1) + " " + Ts(Y1) + " " + Ts(z1)
End If


Players(j).pos.X = X1
Players(j).pos.Y = Y1
Players(j).pos.Z = z1

'Source.Top = (Y - Int(Source.Height / 2))
'Source.Left = X - Int(Source.Width / 2)

Update2

End Sub

Private Sub Picture1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)


If X < 0 Or Y < 0 Then Exit Sub

w = (Picture1.Width / Screen.TwipsPerPixelX)
mp = Int(8192 / w)
mp2 = w / 64

X1 = (X - Int(w / 2)) * mp
Y1 = (((Picture1.Height / Screen.TwipsPerPixelY) - Y) - Int(w / 2)) * mp
z1 = MapArray(Int(X / mp2), Int(((Picture1.Height / Screen.TwipsPerPixelY) / mp2) - (Y / mp2)))
z1 = ConvertPoint(z1)

Label1 = Ts(X1) + ", " + Ts(Y1) + ", " + Ts(z1)
If Check3.Value = 1 Then Me.Caption = "Map"

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Shift = 3 Then
w = (Picture1.Width / Screen.TwipsPerPixelX)
mp = Int(8192 / w)
mp2 = w / 64

    X2 = Int(X / mp2)
    Y2 = Int(((Picture1.Height / Screen.TwipsPerPixelY) / mp2) - (Y / mp2))
    
    If X2 >= 0 And X2 <= 64 And Y2 >= 0 And Y2 <= 64 Then
    
        If MapArray(X2, Y2) <> 0 Then
        
            SendPacket "RC", "setgrid " + Ts(X2) + " " + Ts(Y2) + " 0"
            w = (Picture1.Width / Screen.TwipsPerPixelX)
            mp = w / 64
            Picture1.Line (X2 * mp, (Picture1.Height / Screen.TwipsPerPixelY) - (Y2 * mp))-((X2 * mp) + mp, (Picture1.Height / Screen.TwipsPerPixelY) - ((Y2 * mp) + mp)), 0, BF
            MapArray(X2, Y2) = 0
        End If
    End If
End If

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Label1 = ""
If Check3.Value = 1 Then Me.Caption = "Map"

If X < 0 Or Y < 0 Then Exit Sub

If X > Picture1.Width / Screen.TwipsPerPixelX Or Y > Picture1.Height / Screen.TwipsPerPixelY Then Exit Sub


w = (Picture1.Width / Screen.TwipsPerPixelX)
mp = Int(8192 / w)
mp2 = w / 64

X1 = (X - Int(w / 2)) * mp
Y1 = (((Picture1.Height / Screen.TwipsPerPixelY) - Y) - Int(w / 2)) * mp
z1 = MapArray(Int(X / mp2), Int(((Picture1.Height / Screen.TwipsPerPixelY) / mp2) - (Y / mp2)))

X2 = Int(X / mp2)
Y2 = Int(((Picture1.Height / Screen.TwipsPerPixelY) / mp2) - (Y / mp2))

tt1 = GetMapArrayTeam(z1)
z1 = ConvertPoint(z1)

Label1 = Ts(X1) + ", " + Ts(Y1) + ", " + Ts(z1) + vbCrLf + NameTeam(tt1)




If Shift = 3 And Button = 1 Then
w = (Picture1.Width / Screen.TwipsPerPixelX)
mp = Int(8192 / w)
mp2 = w / 64

    X2 = Int(X / mp2)
    Y2 = Int(((Picture1.Height / Screen.TwipsPerPixelY) / mp2) - (Y / mp2))
        SendPacket "RC", "setgrid " + Ts(X2) + " " + Ts(Y2) + " 0"

    w = (Picture1.Width / Screen.TwipsPerPixelX)
    mp = w / 64
    Picture1.Line (X2 * mp, (Picture1.Height / Screen.TwipsPerPixelY) - (Y2 * mp))-((X2 * mp) + mp, (Picture1.Height / Screen.TwipsPerPixelY) - ((Y2 * mp) + mp)), 0, BF


End If

End Sub

Function NameTeam(tt1) As String

If tt1 = 0 Then NameTeam = "Unknown Team"
If tt1 = 1 Then NameTeam = "Blue Team"
If tt1 = 2 Then NameTeam = "Red Team"
If tt1 = 3 Then NameTeam = "Yellow Team"
If tt1 = 4 Then NameTeam = "Green Team"
If tt1 = 5 Then NameTeam = "Generic Team"




End Function

Function ConvertPoint(z1) As Integer

tt1 = GetMapArrayTeam(z1)

If tt1 = 1 Then z1 = z1 - 8192
If tt1 = 2 Then z1 = z1 - 16384
If tt1 = 3 Then z1 = z1 + 8192
If tt1 = 4 Then z1 = z1 + 16384
If tt1 = 5 Then z1 = z1 - 24576

ConvertPoint = z1

End Function


Private Sub RotImgSet_Click()

End Sub

Private Sub RotImg_DragDrop(Source As Control, X As Single, Y As Single)

Picture1_DragDrop Source, (X / Screen.TwipsPerPixelX) + RotImg.Left, (Y / Screen.TwipsPerPixelY) + RotImg.Top

End Sub

Private Sub RotImg_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

Picture1_DragOver Source, (X / Screen.TwipsPerPixelX) + RotImg.Left, (Y / Screen.TwipsPerPixelY) + RotImg.Top, State

End Sub

Private Sub RotImg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Picture1_MouseMove Button, Shift, (X / Screen.TwipsPerPixelX) + RotImg.Left, (Y / Screen.TwipsPerPixelY) + RotImg.Top

End Sub

Private Sub Timer1_Timer()


If RotImg.Visible = True Then
    
    RotImgNum = RotImgNum + 1
    
    nm = RotImgNum
    If Picture1.Width / Screen.TwipsPerPixelX < 300 Then nm = nm + 4
    RotImg.Picture = RotImgSet.ListImages(nm).Picture
    If RotImgNum = 4 Then RotImgNum = 0
    RotImg.ZOrder 1
    
End If

End Sub

