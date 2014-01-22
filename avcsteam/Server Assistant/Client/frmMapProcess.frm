VERSION 5.00
Begin VB.Form frmMapProcess 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Statistics"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   Icon            =   "frmMapProcess.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   6465
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   4140
      TabIndex        =   5
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   5160
      TabIndex        =   3
      Top             =   5280
      Width           =   975
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5175
      LargeChange     =   50
      Left            =   6180
      SmallChange     =   10
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      Height          =   5175
      Left            =   60
      ScaleHeight     =   341
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   0
      Top             =   60
      Width           =   6075
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   4995
         Left            =   0
         ScaleHeight     =   333
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   397
         TabIndex        =   1
         Top             =   0
         Width           =   5955
         Begin VB.Label Label1 
            Caption         =   "Total Maps Played:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   300
            TabIndex        =   4
            Top             =   120
            Width           =   4875
         End
      End
   End
End
Attribute VB_Name = "frmMapProcess"
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

Public Sub DrawIt()

For i = 1 To NumMapProcess
    total = total + MapProcess(i).TimesPlayed
Next i

For i = 1 To NumMapProcess
    prc = 100 * (MapProcess(i).TimesPlayed / total)
    MapProcess(i).Perc = prc
Next i


'sort maps

Do

    done = 0
    For i = 1 To NumMapProcess - 1
        
        If MapProcess(i + 1).Perc > MapProcess(i).Perc Then
            Swap MapProcess(i + 1).Perc, MapProcess(i).Perc
            Swap MapProcess(i + 1).MapName, MapProcess(i).MapName
            Swap MapProcess(i + 1).LastTimePlayed, MapProcess(i).LastTimePlayed
            Swap MapProcess(i + 1).TimesPlayed, MapProcess(i).TimesPlayed
            done = 1
        End If
    Next i

Loop Until done = 0

Picture1.Cls





Picture1.Height = (NumMapProcess * 50) + 50

For i = 1 To NumMapProcess
    
    Y = i * 50
    X = 150
    
    prc = Int(MapProcess(i).Perc)
    Picture1.Line (X - 1, Y - 1)-(X + 101, Y + 11), RGB(0, 0, 255), BF
    Picture1.Line (X, Y)-(X + prc, Y + 10), RGB(255, 255, 0), BF

    'text
    
    Picture1.PSet (X + 110, Y), Picture1.BackColor
    Picture1.Print Ts(Round(MapProcess(i).Perc, 2)) + "%"
    
    Picture1.PSet (X + 160, Y), Picture1.BackColor
    Picture1.Print "Total: " + Ts(MapProcess(i).TimesPlayed)
    
    Picture1.PSet (10, Y), Picture1.BackColor
    Picture1.Print MapProcess(i).MapName
    
    Picture1.PSet (10, Y + 15), Picture1.BackColor
    Picture1.Print "Last time played: " + Format(MapProcess(i).LastTimePlayed, "ddd, mmm d, yyyy hh:mm:ss AMPM")

Next i

Picture1.Height = Y + 50

mx = Picture1.Height - (Picture2.Height / Screen.TwipsPerPixelY)

If mx > 0 Then
    VScroll1.Visible = True
    VScroll1.Max = mx
End If

Label1 = "Total Maps Played: " + Ts(total)

End Sub

Private Sub Command1_Click()
Unload Me



End Sub

Private Sub Command2_Click()
SendPacket "MP", ""

End Sub

Private Sub Form_Load()

DrawIt
End Sub

Private Sub VScroll1_Change()

Picture1.Top = -VScroll1.Value


End Sub

Private Sub VScroll1_Scroll()
Picture1.Top = -VScroll1.Value
End Sub
