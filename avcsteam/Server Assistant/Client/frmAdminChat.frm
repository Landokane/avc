VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAdminChat 
   Caption         =   "Admin Chat"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13290
   Icon            =   "frmAdminChat.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   644
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   886
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   7500
      Picture         =   "frmAdminChat.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5520
      Width           =   300
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   8760
      Top             =   1080
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10200
      Top             =   1500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   66
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":0E54
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":1AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":26FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":3350
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":3FA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":4BF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":584C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":64A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":70F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":7D48
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":899C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":95F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":A244
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":AE98
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":BAEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":C740
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":D394
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":DFE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":EC3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":F890
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":104E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":11138
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":11D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":129E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":13634
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":14288
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":14EDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":15B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":16784
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":173D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":1802C
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":18C80
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":198D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":1A528
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":1B17C
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":1BDD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":1CA24
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":1D678
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":1E2CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":1EF20
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":1FB74
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":207C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":2141C
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":22070
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":22CC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":23918
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":2456C
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":251C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":25E14
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":26A68
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":276BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":2830E
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":28F62
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":29BB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":2A80A
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":2B45E
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":2C0B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":2CD06
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":2D95A
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":2E5AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":2F202
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":2FE56
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":30AAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":316FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":32352
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminChat.frx":32FA6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   8220
      Top             =   1080
   End
   Begin VB.Frame Frame1 
      Caption         =   "Name Colour"
      Height          =   2775
      Left            =   420
      TabIndex        =   2
      Top             =   660
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CommandButton Command2 
         Caption         =   "Click Here to select a custom picture. It MUST be 32x32, BMP file."
         Height          =   735
         Left            =   60
         TabIndex        =   17
         Top             =   1920
         Width           =   3195
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Set"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   1500
         Width           =   615
      End
      Begin VB.PictureBox Picture2 
         Height          =   1095
         Left            =   2220
         ScaleHeight     =   1035
         ScaleWidth      =   975
         TabIndex        =   9
         Top             =   360
         Width           =   1035
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   255
         Left            =   300
         TabIndex        =   3
         Top             =   720
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   450
         _Version        =   393216
         Max             =   255
         TickFrequency   =   16
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   300
         TabIndex        =   5
         Top             =   300
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   450
         _Version        =   393216
         Max             =   255
         TickFrequency   =   16
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   255
         Left            =   300
         TabIndex        =   7
         Top             =   1140
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   450
         _Version        =   393216
         Max             =   255
         TickFrequency   =   16
      End
      Begin VB.Label Label3 
         Caption         =   "B"
         Height          =   195
         Left            =   60
         TabIndex        =   8
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "R"
         Height          =   195
         Left            =   60
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "G"
         Height          =   195
         Left            =   60
         TabIndex        =   4
         Top             =   780
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   555
      Index           =   0
      Left            =   8460
      Picture         =   "frmAdminChat.frx":33BFA
      ScaleHeight     =   495
      ScaleWidth      =   855
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00000000&
      Height          =   5295
      Left            =   0
      ScaleHeight     =   349
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   501
      TabIndex        =   12
      Top             =   0
      Width           =   7575
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   5295
         Left            =   0
         ScaleHeight     =   353
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   497
         TabIndex        =   13
         Top             =   -60
         Width           =   7455
         Begin RichTextLib.RichTextBox Text 
            Height          =   1335
            Index           =   0
            Left            =   2700
            TabIndex        =   18
            Top             =   120
            Visible         =   0   'False
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   2355
            _Version        =   393217
            BackColor       =   0
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            MousePointer    =   6
            Appearance      =   0
            TextRTF         =   $"frmAdminChat.frx":3433C
            MouseIcon       =   "frmAdminChat.frx":34423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Image Pics 
            Height          =   480
            Index           =   0
            Left            =   960
            Top             =   60
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Names 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Avatar-X:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Index           =   0
            Left            =   1500
            TabIndex        =   15
            Top             =   180
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label ChatTime 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "12:45:10"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   270
            Index           =   0
            Left            =   60
            TabIndex        =   14
            Top             =   180
            Visible         =   0   'False
            Width           =   930
            WordWrap        =   -1  'True
         End
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5295
      LargeChange     =   100
      Left            =   7740
      SmallChange     =   10
      TabIndex        =   11
      Top             =   60
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   7980
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   5580
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmAdminChat.frx":3473D
      Top             =   5520
      Width           =   7455
   End
End
Attribute VB_Name = "frmAdminChat"
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

Dim ChatCol As Long
Dim LastPlay As Double
Public OldWindowProc As Long

Private Type typURLList
    URLAddr() As String
    URLStartPos() As Long
    URLEndPos() As Long
    NumURL As Integer
End Type
Private Const EM_CHARFROMPOS& = &HD7
Dim URLList(50) As typURLList

Dim lastSelIndex As Long
Dim lastSelStart As Long
Dim lastSelLen As Long
Dim lastTextX As Long
Dim lastTextY As Long

Public Sub AddChat(Msg$, Nme$, col, tm$)

On Error Resume Next

If col < 0 Then col = 0

If Round(Timer - LastPlay, 2) < 0 Then LastPlay = 0
If Round(Timer - LastPlay, 2) > 1 And MyAwayMode = 0 Then
    If InStr(1, LCase(Msg), "bing!") And MDIForm1.mnuSettingsIn(7).Checked Then PlayWaveRes 103: LastPlay = Timer
End If
' Adds a line of chat.

Dim TotalNumber As Integer

TotalNumber = Names.Count - 1

maxattime = 40

If TotalNumber < maxattime Then
    ' Add a new one
        
    WeUse = TotalNumber + 1
    
    Load ChatTime(WeUse)
    Load Names(WeUse)
    Load Text(WeUse)
    Load Pics(WeUse)
    
    URLList(WeUse).NumURL = 0
    
Else
    ' Shift em up
    
    For i = 2 To TotalNumber
    
        ChatTime(i - 1).Caption = ChatTime(i).Caption
        Names(i - 1).Caption = Names(i).Caption
        Names(i - 1).Tag = Names(i).Tag
        Names(i - 1).ForeColor = Names(i).ForeColor
        Text(i - 1).TextRTF = Text(i).TextRTF
        Text(i - 1).Tag = Text(i).Tag
                
        
        'Text(i - 1).MousePointer = Text(i).MousePointer
        Pics(i - 1).Picture = Pics(i).Picture
        
        Text(i - 1).Width = Text(i).Width
        Text(i - 1).Left = Text(i).Left
        
        URLList(i - 1) = URLList(i)
        
        OrganizeStuff i - 1
    
    Next i
    
    If lastSelIndex > 0 Then lastSelIndex = lastSelIndex - 1
    
    WeUse = TotalNumber
    URLList(WeUse).NumURL = 0
    Text(WeUse).TextRTF = ""
    Text(WeUse).Text = ""
    Text(WeUse).SelUnderline = False
End If

' Fill it in with the infomation

' position em first.

'is it a /me?
isme = 0
If Len(Msg$) > 4 Then
    If Left(LCase(Msg$), 3) = "/me" Then
        isme = 1
    End If
End If

'is it a HYPERLINK?
Text(WeUse).Tag = ""
'Text(WeUse).ForeColor = RGB(255, 255, 255)
'Text(WeUse).FontUnderline = False
Text(WeUse).MousePointer = 0

textcolor = RGB(255, 255, 255)


'ALRIGHT! It's time to go thru the text letter by letter, wrap it, and decide if we have a URL.
'First, determine how much width we have.

If isme = 1 Then
    Msg$ = Right(Msg$, Len(Msg$) - 4)
    Names(WeUse).Caption = "* " + Nme$
    textcolor = col
Else
    Names(WeUse).Caption = Nme$ + ":"
End If


Text(0).Left = Names(WeUse).Left + Names(WeUse).Width + 5
availwidth = Picture3.Width - Text(WeUse).Left - 10
Text(0).Width = availwidth
'Text(0).Refresh

'Scan thru the message, one letter at time.

msg2$ = Msg$


Dim Wrd As String
Wrd = ""
wordstart = 0

msg2$ = msg2$ + " "

Names(WeUse).Tag = msg2$

e = 0
f = 0

For i = 1 To 100
   
    e = InStr(f + 1, msg2$, "[url=", vbTextCompare)
    
    If e > 0 Then
    
        ' Find the position of the [/url]
        
        g = InStr(e + 1, msg2$, "]", vbTextCompare)
        f = InStr(g + 1, msg2$, "[/url]", vbTextCompare)
        
        If f > g And g > e Then
            'Extract the URL
            url = Mid(msg2$, e + 5, g - e - 5)
            
            'Extract the text
            txt1 = Mid(msg2$, g + 1, f - g - 1)
            
            If InStr(1, url, "@") And InStr(1, url, ".") Then
                url = "mailto:" & url
            End If
            
            
            'Grab the stuff from before ...
            
            leftside$ = Left(msg2$, e - 1)
            rightside$ = Right(msg2$, Len(msg2$) - f - 5)
            
            ' Assemble the new thingy
            
            msg2$ = leftside$ & txt1 & rightside$
            
            wordstart = Len(leftside$) + 1
                        
            ' Add to URL list.
            
            URLList(WeUse).NumURL = URLList(WeUse).NumURL + 1
            n = URLList(WeUse).NumURL
            ReDim Preserve URLList(WeUse).URLAddr(0 To n)
            ReDim Preserve URLList(WeUse).URLStartPos(0 To n)
            ReDim Preserve URLList(WeUse).URLEndPos(0 To n)
            
            URLList(WeUse).URLAddr(n) = url
            URLList(WeUse).URLStartPos(n) = wordstart - 1
            URLList(WeUse).URLEndPos(n) = wordstart - 1 + Len(txt1)
                
        End If
        
        f = Len(leftside$) + Len(txt1)
    End If
    If f = 0 Then f = e
    If e = 0 Then Exit For
Next

Text(0).Text = msg2$

numlines = SendMessage(Text(0).hwnd, EM_GETLINECOUNT, 0, 0)
Text(WeUse).Height = numlines * 16 + 4

Text(WeUse).Text = msg2$

' Set the color of the text.

With Text(WeUse)
    .SelStart = 0
    .SelLength = Len(msg2$)
    .SelColor = textcolor
    .SelUnderline = False
    .SelLength = 0
    .SelStart = 0
End With

' Look for URLs
For i = 1 To URLList(WeUse).NumURL
        With Text(WeUse)
            .SelStart = URLList(WeUse).URLStartPos(i)
            .SelLength = URLList(WeUse).URLEndPos(i) - URLList(WeUse).URLStartPos(i)
            .SelColor = RGB(65, 65, 255)
            .SelUnderline = True
            .SelStart = 0
            .SelLength = 0
        End With
Next i


For i = 1 To Len(msg2$)

    ' www.hat.com, www.someotherurl.com
    
        
    b$ = Mid(msg2$, i, 1)
    
    If b$ = " " Or b$ = "," Or b$ = "<" Or b$ = ">" Or b$ = vbCr Or b$ = vbLf Or b$ = "(" Or b$ = ")" Or b$ = "," Then
        ' See what this word is.
            
        If Len(Wrd) > 7 Then
            
            Wrd2 = LCase(Wrd)
            
            isurl = 0
            If Left(Wrd2, 3) = "www" And InStr(1, Wrd2, ".") > 0 Then
                ' Is URL.
                url = "http://" & Wrd
                isurl = 1
            ElseIf InStr(1, Wrd2, "://") > 0 Then
                url = Wrd
                isurl = 1
            ElseIf InStr(1, Wrd2, "@") > 0 And InStr(1, Wrd2, ".") > 0 Then
                url = "mailto:" & Wrd
                isurl = 1
            End If
            
            If isurl = 1 Then
                ' Select and colour the word.
                                
                If Right(url, 1) = "." Then
                    url = Left(url, Len(url) - 1)
                    Wrd = Left(Wrd, Len(Wrd) - 1)
                End If
                                
            
                ' Add to URL list.
                
                doit22 = 1
                ' Make sure not already there
                For k = 1 To URLList(WeUse).NumURL
                    If URLList(WeUse).URLStartPos(k) = wordstart - 1 Then
                        doit22 = 0: Exit For
                    End If
                Next k
                
                If doit22 Then
                    URLList(WeUse).NumURL = URLList(WeUse).NumURL + 1
                    n = URLList(WeUse).NumURL
                    ReDim Preserve URLList(WeUse).URLAddr(0 To n)
                    ReDim Preserve URLList(WeUse).URLStartPos(0 To n)
                    ReDim Preserve URLList(WeUse).URLEndPos(0 To n)
                    
                    URLList(WeUse).URLAddr(n) = url
                    URLList(WeUse).URLStartPos(n) = wordstart - 1
                    URLList(WeUse).URLEndPos(n) = wordstart - 1 + Len(Wrd)
                
                    With Text(WeUse)
                        .SelStart = wordstart - 1
                        .SelLength = Len(Wrd)
                        .SelColor = RGB(65, 65, 255)
                        .SelUnderline = True
                        .SelStart = 0
                        .SelLength = 0
                    End With
                End If
            End If
        End If
            
            
        Wrd = ""
    Else
        If Wrd = "" Then wordstart = i
        Wrd = Wrd & b$
    End If
    
Next i




ChatTime(WeUse).Caption = tm$
Names(WeUse).ForeColor = col

' See if we have this admin's BMP on file.
use = 0
For i = 1 To Picture5.Count - 1
    If Picture5(i).Tag = Nme$ Then
        ' Yes this admin's BMP is loaded.
        use = i
        Exit For
    End If
Next i

If use = 0 Then
    
    'Load the BMP.
    n = Picture5.Count - 1
    
    n = n + 1
    
    For i = 1 To NumAdminBMP
        If AdminBMP(i).AdminName = Nme$ Then k = i: Exit For
    Next i
    
    If k > 0 Then
        If CheckForFile(App.Path + "\apics\" + AdminBMP(k).BMPFile) Then
            Load Picture5(n)
            On Error Resume Next
            
            Picture5(n).Picture = LoadPicture(App.Path + "\apics\" + AdminBMP(k).BMPFile)
            Picture5(n).Tag = Nme$
            use = n
        End If
    End If
    
    
End If


If use <> 0 Then
    Pics(WeUse).Picture = Picture5(use).Picture
Else
    Pics(WeUse).Picture = Picture5(0).Picture
End If

addamt = 12
OrganizeStuff WeUse


Text(WeUse).Visible = True
ChatTime(WeUse).Visible = True
Names(WeUse).Visible = True
Pics(WeUse).Visible = True


h = Text(WeUse).Top + Text(WeUse).Height + addamt
h2 = Pics(WeUse).Top + Pics(WeUse).Height + addamt
If h2 > h Then h = h2

Picture3.Height = h

VScroll1.Max = Picture3.Height - Picture4.Height
VScroll1.Value = VScroll1.Max
Picture3.Top = -VScroll1.Max

'Form_Resize

End Sub

Sub OrganizeStuff(WeUse)



addamt = 12

h = Text(WeUse - 1).Top + Text(WeUse - 1).Height + addamt
h2 = Pics(WeUse - 1).Top + Pics(WeUse - 1).Height + addamt
If h2 > h Then h = h2


ChatTime(WeUse).Top = h
Pics(WeUse).Top = h - (ChatTime(0).Top - Pics(0).Top)
Names(WeUse).Top = h
Text(WeUse).Top = h - 2

ChatTime(WeUse).Left = ChatTime(0).Left
Pics(WeUse).Left = Pics(0).Left
Names(WeUse).Left = Names(0).Left


Text(WeUse).Left = Names(WeUse).Left + Names(WeUse).Width + 5

w3 = Picture3.Width - Text(WeUse).Left - 10
If w3 < 0 Then Exit Sub

Text(WeUse).Width = w3
'Text(WeUse).Refresh

numlines = SendMessage(Text(WeUse).hwnd, EM_GETLINECOUNT, 0, 0)
Text(WeUse).Height = numlines * 16 + 4


End Sub



Sub ShowCol(Optional nosetslide As Boolean)

If nosetslide Then
    ChatCol = RGB(Slider1, Slider2, Slider3)
End If

Picture1.BackColor = ChatCol
Picture2.BackColor = ChatCol

c = ChatCol

r = c Mod 256
g = (c \ 256) Mod 256
b = c \ 256 \ 256

If Not nosetslide Then Slider1.Value = r
If Not nosetslide Then Slider2.Value = g
If Not nosetslide Then Slider3.Value = b

End Sub

Private Sub ChatTime_Click(Index As Integer)
Text2.SetFocus
End Sub

Private Sub ChatTime_DblClick(Index As Integer)
AskEasy
End Sub

Private Sub Command1_Click()

If Frame1.Visible = True Then
    Frame1.Visible = False
    Picture4.Visible = True
Else
    Frame1.Visible = True
    Picture4.Visible = False
End If

End Sub

Private Sub Command2_Click()

Form1.Dlg1.DialogTitle = "Select 32x32 BMP file"
Form1.Dlg1.Filter = "BMP Files|*.bmp"

Form1.Dlg1.InitDir = App.Path
Form1.Dlg1.MaxFileSize = 12000
Form1.Dlg1.ShowOpen

a$ = Form1.Dlg1.FileName

If a$ = "" Then Exit Sub

'package this file

PackageAdminBMP a$



End Sub


Public Sub clearChat()
    'Clear the admin chat.
    
TotalNumber = Names.Count - 1

    For i = 1 To TotalNumber
        Unload ChatTime(i)
        Unload Names(i)
        Unload Text(i)
        Unload Pics(i)
    Next i
    TotalNumber = 0
    
End Sub

Private Sub Command3_Click()
frmAddURL.Show

End Sub

Private Sub Form_Load()
ShowChat = True

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

Form_Resize

ChatCol = GetSetting("Server Assistant Client", "Window", "adminchatcol", RGB(200, 200, 200))
ShowCol

before = GetSetting("Server Assistant Client", "Window", "ACFirst", 0)

If before = 0 Then
    Picture3.Visible = True
    Me.Show
    DoEvents
    
    MessBox "This is the first time you are opening the Admin Chat window." + vbCrLf + "You can set a custom chat icon and name colour by clicking the" + vbCrLf + "small box marked out with the arrow, in the bottom right corner of the window."
    SaveSetting "Server Assistant Client", "Window", "ACFirst", "1"
    
End If

AddForm True, 140, 140, 0, 0, Me

End Sub

Private Sub Form_Resize()

w = Me.ScaleWidth
h = Me.ScaleHeight
'
If Me.WindowState = 1 Then Exit Sub
'
If h < 140 Then h = 140: Me.Height = 140 * Screen.TwipsPerPixelY


Picture4.Width = w - VScroll1.Width - 2
Picture3.Width = Picture4.Width

'Image1.Left = w - Image1.Width

'Text1.Width = w - Text1.Left - 120
Text2.Width = Picture4.Width - Picture1.Width - 4 + VScroll1.Width + 2 - 4 - Command3.Width


Command3.Left = Text2.Width + Text2.Left + 4
Picture1.Left = Text2.Width + Text2.Left + 4 + Command3.Width + 4

Picture3.Top = Picture1.Top - Picture3.Height
'Command3.Left = Picture1.Left - Picture3.Width


VScroll1.Left = Picture4.Left + Picture4.Width + 2
Picture4.Height = h - Picture4.Top - 4 - Text2.Height '- 25
VScroll1.Height = Picture4.Height

Text2.Top = Picture4.Top + Picture4.Height + 4
TotalNumber = Names.Count - 1

For i = 1 To TotalNumber
    OrganizeStuff i
Next i

addamt = 12

h = Text(TotalNumber).Top + Text(TotalNumber).Height + addamt
h2 = Pics(TotalNumber).Top + Pics(TotalNumber).Height + addamt
If h2 > h Then h = h2

Picture3.Height = h

VScroll1.Max = Picture3.Height - Picture4.Height
VScroll1.Value = VScroll1.Max
Picture3.Top = -VScroll1.Max

Picture1.Top = Text2.Top + 2
Command3.Top = Text2.Top



End Sub

Private Sub Form_Unload(Cancel As Integer)
ShowChat = False


On Error Resume Next
nm$ = Me.Name
SaveSetting "Server Assistant Client", "Window", nm$ + "winmd", Me.WindowState
SaveSetting "Server Assistant Client", "Window", nm$ + "winh", Me.Height
SaveSetting "Server Assistant Client", "Window", nm$ + "wint", Me.Top
SaveSetting "Server Assistant Client", "Window", nm$ + "winl", Me.Left
SaveSetting "Server Assistant Client", "Window", nm$ + "winw", Me.Width
SaveSetting "Server Assistant Client", "Window", "adminchatcol", ChatCol


End Sub

Private Sub Names_Click(Index As Integer)
Text2.SetFocus
End Sub

Private Sub Names_DblClick(Index As Integer)
AskEasy
End Sub

Private Sub Pics_DblClick(Index As Integer)
AskEasy
End Sub

Private Sub Picture1_Click()

If Frame1.Visible = True Then
    Frame1.Visible = False
    Picture4.Visible = True
Else
    Frame1.Visible = True
    Picture4.Visible = False
End If


End Sub

Private Sub Picture3_Click()

Text2.SetFocus


End Sub

Private Sub Picture3_DblClick()

AskEasy
End Sub

Sub AskEasy()
'n = MessBox("Would you like this text in easy-edit text format?", vbYesNo + vbQuestion, "Editable Text")

'If n = vbYes Then
    
    TotalNumber = Names.Count - 1
    
    For i = 1 To TotalNumber
        a$ = a$ + "[" + ChatTime(i) + "]   " + Names(i) + " " + Names(i).Tag + vbCrLf
    Next i

    frmDebug.Show
    frmDebug.Text1 = a$

'End If


End Sub

Private Sub Slider1_Change()
ShowCol True
End Sub

Private Sub Slider1_Scroll()

ShowCol True

End Sub

Private Sub Slider2_Change()
ShowCol True
End Sub

Private Sub Slider2_Scroll()
ShowCol True
End Sub

Private Sub Slider3_Change()
ShowCol True
End Sub

Private Sub Slider3_Scroll()
ShowCol True
End Sub

Private Sub Text_DblClick(Index As Integer)
AskEasy
End Sub

Private Sub Text_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'Me.Caption = "Mouse over index " & Index

lastTextX = X
lastTextY = Y

If URLList(Index).NumURL > 0 Then

    n = URLList(Index).NumURL
    
    ' Search the URL's
    
    Dim pos As Long
    Dim pt As POINTAPI
    pt.X = X / Screen.TwipsPerPixelX
    pt.Y = Y / Screen.TwipsPerPixelY

    ' Get the character number
    pos = SendMessage(Text(Index).hwnd, EM_CHARFROMPOS, 0&, pt)
    
   ' Me.Caption = pos & ", x: " & x & ", y: " & y
    
    over = 0
    For i = 1 To n
    
        If pos >= URLList(Index).URLStartPos(i) And pos <= URLList(Index).URLEndPos(i) Then
            
            'Me.Caption = "OVER A URL: " & URLList(Index).URLAddr(i)
            
            ' HIGHLIGHT THIS URL.
            
            'e.Caption = "TRYING OUT " & Index & "<>" & lastSelIndex & "  TO " & i
            
            If lastSelIndex <> Index Or (lastSelIndex = Index And lastSelStart <> URLList(Index).URLStartPos(i)) Then
                
                
                With Text(Index)
                    q = .SelStart
                    r = .SelLength
                    
                    If lastSelIndex > 0 Then
                        
                        Text(lastSelIndex).TextRTF = Text(lastSelIndex).Tag
                        lastSelIndex = 0
                        lastSelStart = 0
                    End If
                    
                    .Tag = .TextRTF
                    .SelStart = URLList(Index).URLStartPos(i)
                    .SelLength = URLList(Index).URLEndPos(i) - URLList(Index).URLStartPos(i)
                    .SelColor = RGB(255, 65, 65)
                    .SelUnderline = True
                    
                    .SelStart = q
                    .SelLength = r
                        
                    lastSelIndex = Index
                    lastSelStart = URLList(Index).URLStartPos(i)
                    lastSelLen = URLList(Index).URLEndPos(i) - URLList(Index).URLStartPos(i)
                    
                End With
            End If
            
            over = 1
            
            Exit For
        End If
    Next i

    If over = 1 Then
        
        Text(WeUse).MousePointer = 99
    Else
        
        Text(WeUse).MousePointer = rtfArrowHourglass
        If lastSelIndex > 0 Then
            
            Text(lastSelIndex).TextRTF = Text(lastSelIndex).Tag
            lastSelIndex = 0
            lastSelStart = 0
        End If
    End If

Else

    Text(WeUse).MousePointer = 1
    If lastSelIndex > 0 Then
        
        Text(lastSelIndex).TextRTF = Text(lastSelIndex).Tag
        lastSelIndex = 0
        lastSelStart = 0
    End If

End If




End Sub

Private Sub Text_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


If Button <> 1 Then Exit Sub
  
X = lastTextX
Y = lastTextY

If URLList(Index).NumURL > 0 Then

    n = URLList(Index).NumURL
    
    ' Search the URL's
    
    Dim pos As Long
    Dim pt As POINTAPI
    pt.X = X / Screen.TwipsPerPixelX
    pt.Y = Y / Screen.TwipsPerPixelY

    ' Get the character number
    pos = SendMessage(Text(Index).hwnd, EM_CHARFROMPOS, 0&, pt)
    
   ' Me.Caption = pos & ", x: " & x & ", y: " & y
    
    over = 0
    For i = 1 To n
    
        If pos >= URLList(Index).URLStartPos(i) And pos <= URLList(Index).URLEndPos(i) Then
            
             'ShellExecute MDIForm1.hwnd, "open", URLList(Index).URLAddr(i), vbNullString, vbNullString, SW_SHOW
             ShellExecute MDIForm1.hwnd, "open", "iexplore.exe", URLList(Index).URLAddr(i), vbNullString, SW_SHOW

            Exit For
        End If
    Next i

End If


End Sub

Private Sub Text2_GotFocus()
If Text2 = "Chat Here!" Then Text2 = ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   If Trim(Text2) <> "" Then
        
        If MyAwayMode > 0 Then
            MessBox "Cant chat in Away Mode! Turn it off in the settings menu."
        Else
            
            
            c$ = Text2
            
                        
            'send chat
            
            a$ = Chr(251)
            a$ = a$ + c$ + Chr(250)
            a$ = a$ + Ts(ChatCol) + Chr(250)
            a$ = a$ + Chr(251)
            
            SendPacket "AC", a$
        End If
    End If
    Text2 = ""
    KeyAscii = 0
End If

End Sub

Private Sub Timer1_Timer()

Picture4.Picture = Nothing

Timer1.Enabled = False

End Sub

Private Sub Timer2_Timer()
    Exit Sub
    
'    Static CurrFrame As Integer
    
    If CurrFrame >= 12 Then CurrFrame = 1
    TotalNumber = Names.Count - 1

    For i = 1 To TotalNumber

        CurrFrame = Val(Pics(i).Tag) + 1
        If CurrFrame >= 12 Then CurrFrame = 1

        If InStr(Text(i).Text, "ph34r me!") > 0 Then

            addamt = 55
            If Names(i).Caption = "Avatar-X:" Then addamt = 0
            If Names(i).Caption = "Freaky:" Then addamt = 11
            If Names(i).Caption = "JeffRaven:" Then addamt = 44

            Pics(i).Picture = ImageList1.ListImages(CurrFrame + addamt).Picture
            Pics(i).Tag = Ts(CurrFrame)
        End If
    Next i
   ' Picture6.Visible = False
    
End Sub

Private Sub VScroll1_Change()

Picture3.Top = -VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()

Picture3.Top = -VScroll1.Value

End Sub
