VERSION 5.00
Begin VB.Form frmGeneral 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "General Settings"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   Icon            =   "frmGeneral.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   Begin VB.VScrollBar VScroll1 
      Height          =   5115
      LargeChange     =   500
      Left            =   4500
      SmallChange     =   100
      TabIndex        =   48
      Top             =   60
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      Height          =   5115
      Left            =   60
      ScaleHeight     =   5055
      ScaleWidth      =   4335
      TabIndex        =   2
      Top             =   60
      Width           =   4395
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   14835
         Left            =   0
         ScaleHeight     =   14835
         ScaleWidth      =   4755
         TabIndex        =   3
         Top             =   0
         Width           =   4755
         Begin VB.Frame Frame11 
            Caption         =   "Connecting"
            Height          =   735
            Left            =   60
            TabIndex        =   85
            Top             =   13740
            Width           =   4215
            Begin VB.CheckBox Check16 
               Caption         =   "Disallow more than one person with the same CD-Key at the same time"
               Height          =   375
               Left            =   120
               TabIndex        =   86
               Top             =   240
               Width           =   3555
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "File Transfers"
            Height          =   855
            Left            =   60
            TabIndex        =   79
            Top             =   12840
            Width           =   4215
            Begin VB.TextBox Text17 
               Height          =   285
               Left            =   2820
               TabIndex        =   82
               Top             =   480
               Width           =   795
            End
            Begin VB.TextBox Text16 
               Height          =   285
               Left            =   2820
               TabIndex        =   80
               Top             =   180
               Width           =   495
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               Caption         =   "Newest Client Version:"
               Height          =   195
               Left            =   60
               TabIndex        =   83
               Top             =   540
               Width           =   1590
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               Caption         =   "Limit download speed: (kilobytes/sec)"
               Height          =   195
               Left            =   60
               TabIndex        =   81
               Top             =   240
               Width           =   2655
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Messages"
            Height          =   615
            Left            =   60
            TabIndex        =   52
            Top             =   12180
            Width           =   4215
            Begin VB.CheckBox Check11 
               Caption         =   "Delete old messages after 2 weeks"
               Height          =   195
               Left            =   60
               TabIndex        =   53
               Top             =   240
               Width           =   2835
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "General Votes"
            Height          =   555
            Left            =   60
            TabIndex        =   50
            Top             =   4740
            Width           =   4215
            Begin VB.CheckBox Check10 
               Caption         =   "Use Menu System"
               Height          =   255
               Left            =   60
               TabIndex        =   51
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "RealPlayers"
            Height          =   3135
            Left            =   60
            TabIndex        =   46
            Top             =   9000
            Width           =   4215
            Begin VB.TextBox Flag4 
               Height          =   285
               Left            =   2640
               TabIndex        =   76
               Top             =   2760
               Width           =   1515
            End
            Begin VB.TextBox Flag2 
               Height          =   285
               Left            =   540
               TabIndex        =   74
               Top             =   2760
               Width           =   1515
            End
            Begin VB.TextBox Flag3 
               Height          =   285
               Left            =   2640
               TabIndex        =   72
               Top             =   2400
               Width           =   1515
            End
            Begin VB.TextBox Flag1 
               Height          =   285
               Left            =   540
               TabIndex        =   70
               Top             =   2400
               Width           =   1515
            End
            Begin VB.CheckBox Check14 
               Caption         =   "Don't add if player already has 3 UniqueID's"
               Height          =   255
               Left            =   180
               TabIndex        =   67
               Top             =   1920
               Width           =   3495
            End
            Begin VB.CheckBox Check12 
               Caption         =   "Add uniqueID's for players with same name."
               Height          =   255
               Left            =   60
               TabIndex        =   61
               Top             =   1680
               Width           =   4035
            End
            Begin VB.Frame Frame9 
               Caption         =   "Auto-Adding"
               Height          =   975
               Left            =   60
               TabIndex        =   54
               Top             =   660
               Width           =   4095
               Begin VB.TextBox Text12 
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   56
                  Top             =   600
                  Width           =   495
               End
               Begin VB.TextBox Text13 
                  Height          =   285
                  Left            =   2580
                  TabIndex        =   55
                  Top             =   240
                  Width           =   495
               End
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  Caption         =   "connects"
                  Height          =   195
                  Left            =   3120
                  TabIndex        =   60
                  Top             =   300
                  Width           =   660
               End
               Begin VB.Label Label16 
                  AutoSize        =   -1  'True
                  Caption         =   "Convert to normal realplayers after"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   59
                  Top             =   300
                  Width           =   2400
               End
               Begin VB.Label Label17 
                  AutoSize        =   -1  'True
                  Caption         =   "days"
                  Height          =   195
                  Left            =   3420
                  TabIndex        =   58
                  Top             =   660
                  Width           =   330
               End
               Begin VB.Label Label18 
                  AutoSize        =   -1  'True
                  Caption         =   "Remove from DB if not seen again after"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   57
                  Top             =   660
                  Width           =   2790
               End
            End
            Begin VB.CheckBox Check8 
               Caption         =   "Automatically add unknown players as temporary realplayers"
               Height          =   375
               Left            =   60
               TabIndex        =   47
               Top             =   240
               Width           =   3195
            End
            Begin VB.Label Label26 
               Caption         =   "Custom ToggleFlag Names:"
               Height          =   195
               Left            =   60
               TabIndex        =   78
               Top             =   2160
               Width           =   2235
            End
            Begin VB.Label Label25 
               Caption         =   "Flag 4"
               Height          =   255
               Left            =   2160
               TabIndex        =   77
               Top             =   2820
               Width           =   435
            End
            Begin VB.Label Label24 
               Caption         =   "Flag 2"
               Height          =   255
               Left            =   60
               TabIndex        =   75
               Top             =   2820
               Width           =   435
            End
            Begin VB.Label Label23 
               Caption         =   "Flag 3"
               Height          =   255
               Left            =   2160
               TabIndex        =   73
               Top             =   2460
               Width           =   435
            End
            Begin VB.Label Label22 
               Caption         =   "Flag 1"
               Height          =   255
               Left            =   60
               TabIndex        =   71
               Top             =   2460
               Width           =   435
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Spam Control"
            Height          =   1095
            Left            =   60
            TabIndex        =   36
            Top             =   5340
            Width           =   4215
            Begin VB.TextBox Text1 
               Height          =   285
               Left            =   2640
               TabIndex        =   40
               Text            =   "15"
               Top             =   180
               Width           =   495
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Left            =   3660
               TabIndex        =   39
               Text            =   "6"
               Top             =   180
               Width           =   495
            End
            Begin VB.TextBox Text10 
               Height          =   285
               Left            =   1860
               TabIndex        =   38
               Text            =   "3"
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox Text11 
               Height          =   285
               Left            =   840
               TabIndex        =   37
               Text            =   "20"
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Maximum number of messages per "
               Height          =   195
               Left            =   60
               TabIndex        =   45
               Top             =   240
               Width           =   2475
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "sec is"
               Height          =   195
               Left            =   3180
               TabIndex        =   44
               Top             =   240
               Width           =   405
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "sec is"
               Height          =   195
               Left            =   1380
               TabIndex        =   43
               Top             =   780
               Width           =   405
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Maximum number of times the same bind can be"
               Height          =   195
               Left            =   60
               TabIndex        =   42
               Top             =   540
               Width           =   3855
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "pressed in"
               Height          =   195
               Left            =   60
               TabIndex        =   41
               Top             =   780
               Width           =   720
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Logging"
            Height          =   975
            Left            =   60
            TabIndex        =   33
            Top             =   7980
            Width           =   4215
            Begin VB.CheckBox Check15 
               Caption         =   "(SillyZone) Send Logs to SillyZone Website"
               Height          =   195
               Left            =   60
               TabIndex        =   84
               Top             =   720
               Width           =   3795
            End
            Begin VB.CheckBox Check4 
               Caption         =   "Log Server Events to a file"
               Height          =   195
               Left            =   60
               TabIndex        =   35
               Top             =   240
               Width           =   2535
            End
            Begin VB.CheckBox Check5 
               Caption         =   "Sort Game Server Logs into folders automatically"
               Height          =   195
               Left            =   60
               TabIndex        =   34
               Top             =   480
               Width           =   3795
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Speech (AutoAdmin)"
            Height          =   1455
            Left            =   60
            TabIndex        =   26
            Top             =   6480
            Width           =   4215
            Begin VB.TextBox Text15 
               Height          =   285
               Left            =   1620
               TabIndex        =   64
               Top             =   1020
               Width           =   2535
            End
            Begin VB.CheckBox Check6 
               Caption         =   "Automatically add new speech"
               Height          =   255
               Left            =   60
               TabIndex        =   30
               Top             =   240
               Width           =   2535
            End
            Begin VB.CheckBox Check7 
               Caption         =   "Disable Speech feature"
               Height          =   255
               Left            =   60
               TabIndex        =   29
               Top             =   480
               Width           =   2355
            End
            Begin VB.TextBox Text7 
               Height          =   285
               Left            =   3480
               TabIndex        =   28
               Text            =   "6"
               Top             =   660
               Width           =   495
            End
            Begin VB.TextBox Text8 
               Height          =   285
               Left            =   2460
               TabIndex        =   27
               Text            =   "15"
               Top             =   660
               Width           =   495
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Name of AutoAdmin"
               Height          =   195
               Left            =   60
               TabIndex        =   65
               Top             =   1080
               Width           =   1410
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "sec is"
               Height          =   195
               Left            =   3000
               TabIndex        =   32
               Top             =   720
               Width           =   405
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Max num of admin questions per "
               Height          =   195
               Left            =   60
               TabIndex        =   31
               Top             =   720
               Width           =   2325
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Kick Voting"
            Height          =   2235
            Left            =   60
            TabIndex        =   16
            Top             =   2460
            Width           =   4215
            Begin VB.CheckBox Check9 
               Caption         =   "Use Menu System"
               Height          =   195
               Left            =   60
               TabIndex        =   49
               Top             =   1920
               Width           =   2535
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Enable Kick Votes"
               Height          =   195
               Left            =   60
               TabIndex        =   21
               Top             =   240
               Width           =   2535
            End
            Begin VB.TextBox Text3 
               Height          =   285
               Left            =   3300
               TabIndex        =   20
               Text            =   "2"
               Top             =   480
               Width           =   555
            End
            Begin VB.TextBox Text4 
               Height          =   285
               Left            =   3300
               TabIndex        =   19
               Text            =   "2"
               Top             =   840
               Width           =   555
            End
            Begin VB.TextBox Text5 
               Height          =   285
               Left            =   3300
               TabIndex        =   18
               Text            =   "10"
               Top             =   1200
               Width           =   555
            End
            Begin VB.TextBox Text6 
               Height          =   285
               Left            =   3300
               TabIndex        =   17
               Text            =   "75"
               Top             =   1560
               Width           =   555
            End
            Begin VB.Label Label21 
               Caption         =   "%"
               Height          =   255
               Left            =   3900
               TabIndex        =   69
               Top             =   1620
               Width           =   255
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Max kick votes per player per map"
               Height          =   195
               Left            =   60
               TabIndex        =   25
               Top             =   540
               Width           =   2430
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Number of kicks by kickvote before auto-ban"
               Height          =   195
               Left            =   60
               TabIndex        =   24
               Top             =   900
               Width           =   3195
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Auto-Ban Time (minutes)"
               Height          =   195
               Left            =   60
               TabIndex        =   23
               Top             =   1260
               Width           =   1725
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Required YES percentage before kick"
               Height          =   195
               Left            =   60
               TabIndex        =   22
               Top             =   1620
               Width           =   2700
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Map Votes"
            Height          =   2415
            Left            =   60
            TabIndex        =   4
            Top             =   0
            Width           =   4215
            Begin VB.TextBox Text14 
               Height          =   285
               Left            =   2460
               TabIndex        =   62
               Top             =   2040
               Width           =   1635
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Enable automatic Map Votes"
               Height          =   195
               Left            =   60
               TabIndex        =   12
               Top             =   240
               Width           =   2535
            End
            Begin VB.CheckBox Check3 
               Caption         =   "Maps cannot be voted for again until 3 maps later"
               Height          =   195
               Left            =   60
               TabIndex        =   11
               Top             =   480
               Width           =   3915
            End
            Begin VB.TextBox Text9 
               Height          =   285
               Left            =   1620
               TabIndex        =   10
               Text            =   "4"
               Top             =   720
               Width           =   375
            End
            Begin VB.OptionButton Option1 
               Caption         =   "End of map"
               Height          =   195
               Left            =   3000
               TabIndex        =   9
               Top             =   1020
               Value           =   -1  'True
               Width           =   1155
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Start of map"
               Height          =   195
               Left            =   3000
               TabIndex        =   8
               Top             =   780
               Width           =   1155
            End
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   735
               Left            =   1440
               ScaleHeight     =   735
               ScaleWidth      =   2715
               TabIndex        =   5
               Top             =   1260
               Width           =   2715
               Begin VB.OptionButton Option5 
                  Caption         =   "Utility DLL"
                  Height          =   195
                  Left            =   0
                  TabIndex        =   68
                  Top             =   480
                  Width           =   1395
               End
               Begin VB.CheckBox Check13 
                  Caption         =   "Immediate"
                  Height          =   195
                  Left            =   1260
                  TabIndex        =   66
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton Option3 
                  Caption         =   "Changelevel"
                  Height          =   195
                  Left            =   0
                  TabIndex        =   7
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1395
               End
               Begin VB.OptionButton Option4 
                  Caption         =   "Map Cycle File"
                  Height          =   195
                  Left            =   0
                  TabIndex        =   6
                  Top             =   240
                  Width           =   1395
               End
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "DEBUG: Map Change Command: "
               Height          =   195
               Left            =   60
               TabIndex        =   63
               Top             =   2100
               Width           =   2415
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Auto-Start Map Votes"
               Height          =   195
               Left            =   60
               TabIndex        =   15
               Top             =   780
               Width           =   1515
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "minutes from"
               Height          =   195
               Left            =   2040
               TabIndex        =   14
               Top             =   780
               Width           =   885
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Map change mode:"
               Height          =   195
               Left            =   60
               TabIndex        =   13
               Top             =   1260
               Width           =   1380
            End
         End
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2820
      TabIndex        =   1
      Top             =   5220
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   5220
      Width           =   2715
   End
End
Attribute VB_Name = "frmGeneral"
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

Private Sub Command1_Click()

If Check1 = 1 Then General.NoAutoVotes = False
If Check1 = 0 Then General.NoAutoVotes = True

If Check2 = 1 Then General.NoKickVotes = False
If Check2 = 0 Then General.NoKickVotes = True

If Check3 = 1 Then General.LastMapsDisabled = False
If Check3 = 0 Then General.LastMapsDisabled = True

If Check4 = 1 Then General.LoggingDisabled = False
If Check4 = 0 Then General.LoggingDisabled = True

a = 0
If Check5 Then a = a + 1
If Check6 Then a = a + 2
If Check7 Then a = a + 4
If Check9 Then a = a + 8
If Check10 Then a = a + 16
If Check11 Then a = a + 32
If Check12 Then a = a + 64
If Check14 Then a = a + 128
If Check16 Then a = a + 256


General.Flags = a


General.MaxTime = Val(Text1)
General.MaxMsg = Val(Text2)
General.MaxKickVotes = Val(Text3)
General.MaxKicks = Val(Text4)
General.BanTime = Val(Text5)
General.VotePercent = Val(Text6)
General.MaxSpeech = Val(Text7)
General.MaxSpeechTime = Val(Text8)
General.MapVoteStartTime = Val(Text9) * 60
General.SameSpamNum = Ts(Val(Text10))
General.SameSpamTime = Ts(Val(Text11))
General.AutoAddRealDays = Ts(Val(Text12))
General.AutoAddRealTimes = Ts(Val(Text13))
General.AutoAddReal = Check8
General.MapChangeMode = Text14
General.AutoAdminName = Text15
General.MaxFileSend = Val(Text16)
General.NewestClient = Text17

General.CustomFlag1 = Flag1
General.CustomFlag2 = Flag2
General.CustomFlag3 = Flag3
General.CustomFlag4 = Flag4

If Option1.Value Then General.MapVoteStartTimeMode = 0
If Option2.Value Then General.MapVoteStartTimeMode = 1

If Option3.Value Then General.MapVoteMode = "0"
If Option4.Value Then General.MapVoteMode = "1"
If Check13 Then General.MapVoteMode = "2"
If Option5.Value Then General.MapVoteMode = "3"

General.SendToDisco = Check15


PackageGeneral
Unload Me


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()

VScroll1.Max = Picture3.Height - Picture2.Height

If General.NoAutoVotes = True Then Check1 = 0
If General.NoAutoVotes = False Then Check1 = 1

If General.NoKickVotes = True Then Check2 = 0
If General.NoKickVotes = False Then Check2 = 1

If General.LastMapsDisabled = True Then Check3 = 0
If General.LastMapsDisabled = False Then Check3 = 1

If General.LoggingDisabled = True Then Check4 = 0
If General.LoggingDisabled = False Then Check4 = 1

Text1 = Ts(General.MaxTime)
Text2 = Ts(General.MaxMsg)
Text3 = Ts(General.MaxKickVotes)
Text4 = Ts(General.MaxKicks)
Text5 = Ts(General.BanTime)
Text6 = Ts(General.VotePercent)
Text7 = Ts(General.MaxSpeech)
Text8 = Ts(General.MaxSpeechTime)
Text9 = Ts(General.MapVoteStartTime \ 60)
Text10 = General.SameSpamNum
Text11 = General.SameSpamTime
Text12 = General.AutoAddRealDays
Text13 = General.AutoAddRealTimes
Text14 = General.MapChangeMode
Text15 = General.AutoAdminName
Text16 = Ts(General.MaxFileSend)
Text17 = General.NewestClient

Check8 = General.AutoAddReal

Flag1 = General.CustomFlag1
Flag2 = General.CustomFlag2
Flag3 = General.CustomFlag3
Flag4 = General.CustomFlag4


If Text14 = "" Then Text14 = "changelevel"

If General.MapVoteStartTimeMode = 0 Then Option1.Value = True
If General.MapVoteStartTimeMode = 1 Then Option2.Value = True

If General.MapVoteMode <> "1" Then Option3.Value = True
If General.MapVoteMode = "1" Then Option4.Value = True
If General.MapVoteMode = "2" Then Check13 = 1
If General.MapVoteMode = "3" Then Option5.Value = True

If CheckBit2(General.Flags, 0) Then Check5 = 1
If CheckBit2(General.Flags, 1) Then Check6 = 1
If CheckBit2(General.Flags, 2) Then Check7 = 1
If CheckBit2(General.Flags, 3) Then Check9 = 1
If CheckBit2(General.Flags, 4) Then Check10 = 1
If CheckBit2(General.Flags, 5) Then Check11 = 1
If CheckBit2(General.Flags, 6) Then Check12 = 1
If CheckBit2(General.Flags, 7) Then Check14 = 1
If CheckBit2(General.Flags, 8) Then Check16 = 1

Check15 = General.SendToDisco

End Sub

Private Sub Text9_Change()
If Val(Text9) < 0 Then Text9 = "4"

End Sub
Private Sub VScroll1_Change()

Picture3.Top = -VScroll1.Value


End Sub

Private Sub VScroll1_Scroll()
VScroll1_Change
End Sub
