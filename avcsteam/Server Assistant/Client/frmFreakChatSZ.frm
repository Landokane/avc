VERSION 5.00
Begin VB.Form frmSZChat 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Private Chat 2 - SillyZone Special Edition"
   ClientHeight    =   6480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   Begin VB.Timer flashOk 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3840
      Top             =   3840
   End
   Begin VB.Timer flashCancel 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5400
      Top             =   3840
   End
   Begin VB.PictureBox picMsgBox 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   2160
      ScaleHeight     =   2295
      ScaleWidth      =   3855
      TabIndex        =   27
      Top             =   1920
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox msgboxDefault 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   480
         TabIndex        =   30
         Text            =   "MsgBoxDefault"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Line Line35 
         BorderColor     =   &H0000FF00&
         X1              =   120
         X2              =   3720
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label msgboxCancel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   2040
         TabIndex        =   32
         Top             =   1720
         Width           =   1335
      End
      Begin VB.Shape shpMsgboxCancel 
         BorderColor     =   &H0000FF00&
         Height          =   375
         Left            =   2040
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label msgboxOK 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   480
         TabIndex        =   31
         Top             =   1720
         Width           =   1335
      End
      Begin VB.Shape shpMsgboxOK 
         BorderColor     =   &H0000FF00&
         Height          =   375
         Left            =   480
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Line lineMsg2 
         BorderColor     =   &H0000FF00&
         X1              =   240
         X2              =   240
         Y1              =   1200
         Y2              =   1440
      End
      Begin VB.Line lineMsg5 
         BorderColor     =   &H0000FF00&
         X1              =   3480
         X2              =   3480
         Y1              =   1200
         Y2              =   1440
      End
      Begin VB.Line lineMsg6 
         BorderColor     =   &H0000FF00&
         X1              =   3360
         X2              =   3480
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line lineMsg4 
         BorderColor     =   &H0000FF00&
         X1              =   3360
         X2              =   3480
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line lineMsg3 
         BorderColor     =   &H0000FF00&
         X1              =   240
         X2              =   360
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line lineMsg1 
         BorderColor     =   &H0000FF00&
         X1              =   360
         X2              =   240
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label msgboxQuestion 
         BackStyle       =   0  'Transparent
         Caption         =   "MsgBoxQuestion"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   3375
      End
      Begin VB.Line Line28 
         BorderColor     =   &H0000FF00&
         X1              =   120
         X2              =   3720
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line27 
         BorderColor     =   &H0000FF00&
         X1              =   120
         X2              =   3720
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label msgboxTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "MsgBoxTitle"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   120
         Width           =   3375
      End
      Begin VB.Line Line26 
         BorderColor     =   &H0000FF00&
         X1              =   3720
         X2              =   3720
         Y1              =   120
         Y2              =   2160
      End
      Begin VB.Line Line25 
         BorderColor     =   &H0000FF00&
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   2160
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   3
         Height          =   2295
         Left            =   0
         Top             =   0
         Width           =   3855
      End
   End
   Begin VB.Timer tmrPingCheck 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   9240
      Top             =   5040
   End
   Begin VB.Timer tmrWhoColour 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1320
      Top             =   600
   End
   Begin VB.PictureBox picExtras 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   8040
      ScaleHeight     =   4215
      ScaleWidth      =   1935
      TabIndex        =   33
      Top             =   720
      Visible         =   0   'False
      Width           =   1935
      Begin VB.Timer flashTime 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1440
         Top             =   2400
      End
      Begin VB.Label lblKick 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Kick"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   240
         TabIndex        =   42
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Shape shpKick 
         BorderColor     =   &H0000FF00&
         Height          =   495
         Left            =   240
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Shape shpTime 
         BorderColor     =   &H0000FF00&
         Height          =   495
         Left            =   240
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Shape shpLogToFile 
         BorderColor     =   &H0000FF00&
         Height          =   495
         Left            =   240
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label lblLogToFile 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Log Chat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lblState 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "State Msg"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Shape shpState 
         BorderColor     =   &H0000FF00&
         Height          =   495
         Left            =   240
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Line Line22 
         BorderColor     =   &H0000FF00&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   4200
      End
      Begin VB.Line Line23 
         BorderColor     =   &H0000FF00&
         X1              =   120
         X2              =   0
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line24 
         BorderColor     =   &H0000FF00&
         X1              =   120
         X2              =   0
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Shape shpLevel 
         BorderColor     =   &H0000FF00&
         Height          =   495
         Left            =   240
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label lblLevel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   120
         Width           =   1575
      End
      Begin VB.Shape shpGoodMessage 
         BorderColor     =   &H0000FF00&
         Height          =   495
         Left            =   240
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblGoodMessage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Message 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   720
         Width           =   1575
      End
      Begin VB.Shape shpQuestion 
         BorderColor     =   &H0000FF00&
         Height          =   495
         Left            =   240
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblQuestion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Question"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.Timer flashExtra 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7560
      Top             =   2640
   End
   Begin VB.Timer flashLeave 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7920
      Top             =   5880
   End
   Begin VB.Timer flashAway 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7560
      Top             =   4560
   End
   Begin VB.Timer flashAction 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7560
      Top             =   2040
   End
   Begin VB.Timer flashMessage 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7560
      Top             =   1440
   End
   Begin VB.Timer flashBeep 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7560
      Top             =   840
   End
   Begin VB.Timer flashSend 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4920
      Top             =   5160
   End
   Begin VB.Timer tmrAway 
      Interval        =   1000
      Left            =   5340
      Top             =   5160
   End
   Begin VB.Timer tmrDuration 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4560
      Top             =   5160
   End
   Begin VB.Timer StatusMove 
      Interval        =   100
      Left            =   5760
      Top             =   5160
   End
   Begin VB.TextBox txtDisplay 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   3375
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   5895
   End
   Begin VB.TextBox txtSay 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   4560
      Width           =   4215
   End
   Begin VB.Line Line21 
      BorderColor     =   &H0000FF00&
      X1              =   120
      X2              =   120
      Y1              =   5640
      Y2              =   5040
   End
   Begin VB.Shape Shape5 
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   5040
      Width           =   135
   End
   Begin VB.Line Line29 
      BorderColor     =   &H0000FF00&
      X1              =   6240
      X2              =   6240
      Y1              =   5640
      Y2              =   5040
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   3615
      Left            =   120
      Top             =   720
      Width           =   6135
   End
   Begin VB.Label lblMove 
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   6960
      TabIndex        =   41
      ToolTipText     =   "Move"
      Top             =   120
      Width           =   315
   End
   Begin VB.Label lblTimeStamp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Time Stamp"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6360
      TabIndex        =   39
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Shape shpTimeStamp 
      BorderColor     =   &H0000FF00&
      Height          =   495
      Left            =   6360
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Shape move1 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Left            =   6960
      Shape           =   3  'Circle
      Top             =   160
      Width           =   255
   End
   Begin VB.Line lineExtra 
      BorderColor     =   &H0000FF00&
      Visible         =   0   'False
      X1              =   8040
      X2              =   7920
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label lblExit 
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   7620
      TabIndex        =   26
      ToolTipText     =   "Close"
      Top             =   120
      Width           =   315
   End
   Begin VB.Label lblMinimize 
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   7320
      TabIndex        =   25
      ToolTipText     =   "Minimize"
      Top             =   120
      Width           =   315
   End
   Begin VB.Line exit2 
      BorderColor     =   &H0000FF00&
      X1              =   7860
      X2              =   7665
      Y1              =   180
      Y2              =   375
   End
   Begin VB.Line exit1 
      BorderColor     =   &H0000FF00&
      X1              =   7680
      X2              =   7875
      Y1              =   180
      Y2              =   375
   End
   Begin VB.Line min1 
      BorderColor     =   &H0000FF00&
      X1              =   7320
      X2              =   7560
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Line min3 
      BorderColor     =   &H0000FF00&
      X1              =   7565
      X2              =   7430
      Y1              =   180
      Y2              =   375
   End
   Begin VB.Line min2 
      BorderColor     =   &H0000FF00&
      X1              =   7320
      X2              =   7440
      Y1              =   180
      Y2              =   360
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   3840
      Top             =   6720
      Width           =   495
   End
   Begin VB.Line Line20 
      BorderColor     =   &H0000FF00&
      X1              =   3960
      X2              =   3960
      Y1              =   600
      Y2              =   120
   End
   Begin VB.Label lblExtra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Extra"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6360
      TabIndex        =   24
      Top             =   2640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Shape shpExtra 
      BorderColor     =   &H0000FF00&
      Height          =   495
      Left            =   6360
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Line Line19 
      BorderColor     =   &H0000FF00&
      X1              =   4680
      X2              =   4680
      Y1              =   600
      Y2              =   120
   End
   Begin VB.Label lblLeaveNo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7320
      TabIndex        =   23
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape shpLeaveNo 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Left            =   7200
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblLeaveYes 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6480
      TabIndex        =   22
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape shpLeaveYes 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Left            =   6360
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblLeaveAsk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Leave?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6360
      TabIndex        =   21
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDND 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DND"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6360
      TabIndex        =   20
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Shape shpDND 
      BorderColor     =   &H0000FF00&
      Height          =   495
      Left            =   6360
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Line Line14 
      BorderColor     =   &H0000FF00&
      X1              =   3120
      X2              =   6240
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line12 
      BorderColor     =   &H0000FF00&
      X1              =   3120
      X2              =   6240
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label txtNickName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Unknown"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1080
      TabIndex        =   19
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label lblLeave 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Leave"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6360
      TabIndex        =   18
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Shape shpLeave 
      BorderColor     =   &H0000FF00&
      Height          =   495
      Left            =   6360
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label lblStatusAway 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Unknown is Away"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   3240
      TabIndex        =   17
      Top             =   5880
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblChangeName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   6135
      Width           =   2655
   End
   Begin VB.Shape shpChangeName 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Left            =   120
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5760
      Width           =   735
   End
   Begin VB.Line Line18 
      BorderColor     =   &H0000FF00&
      X1              =   960
      X2              =   1080
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line17 
      BorderColor     =   &H0000FF00&
      X1              =   2880
      X2              =   3000
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line16 
      BorderColor     =   &H0000FF00&
      X1              =   960
      X2              =   1080
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line15 
      BorderColor     =   &H0000FF00&
      X1              =   2880
      X2              =   3000
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line13 
      BorderColor     =   &H0000FF00&
      X1              =   960
      X2              =   960
      Y1              =   6000
      Y2              =   5760
   End
   Begin VB.Line Line11 
      BorderColor     =   &H0000FF00&
      X1              =   3000
      X2              =   3000
      Y1              =   5760
      Y2              =   6000
   End
   Begin VB.Shape shpAway 
      BorderColor     =   &H0000FF00&
      Height          =   495
      Left            =   6360
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label lblAway 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Away"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6360
      TabIndex        =   14
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Shape shpAction 
      BorderColor     =   &H0000FF00&
      Height          =   495
      Left            =   6360
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblAction 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Action"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6360
      TabIndex        =   13
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Shape shpMessage 
      BorderColor     =   &H0000FF00&
      Height          =   495
      Left            =   6360
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6360
      TabIndex        =   12
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Shape shpBeep 
      BorderColor     =   &H0000FF00&
      Height          =   495
      Left            =   6360
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblBeep 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Beep"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6360
      TabIndex        =   11
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5160
      TabIndex        =   10
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5520
      TabIndex        =   9
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblTimeSecond 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblTimeMinute 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblTimeHour 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   240
      Width           =   375
   End
   Begin VB.Line Line10 
      BorderColor     =   &H0000FF00&
      X1              =   120
      X2              =   6240
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0000FF00&
      X1              =   120
      X2              =   6240
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label lblSend 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Shape shpSend 
      BorderColor     =   &H0000FF00&
      Height          =   495
      Left            =   4680
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label lblWho 
      BackStyle       =   0  'Transparent
      Caption         =   "Unknown"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Chat with:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0000FF00&
      X1              =   4440
      X2              =   4560
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line8 
      BorderColor     =   &H0000FF00&
      X1              =   4560
      X2              =   4560
      Y1              =   4440
      Y2              =   4560
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0000FF00&
      X1              =   120
      X2              =   120
      Y1              =   4800
      Y2              =   4920
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FF00&
      X1              =   120
      X2              =   240
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FF00&
      X1              =   120
      X2              =   120
      Y1              =   4560
      Y2              =   4440
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FF00&
      X1              =   4560
      X2              =   4560
      Y1              =   4920
      Y2              =   4800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      X1              =   4440
      X2              =   4560
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      X1              =   120
      X2              =   240
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   495
      Left            =   120
      Top             =   120
      Width           =   6135
   End
   Begin VB.Shape shpSendBack 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4680
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Shape shpDNDBack 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   6360
      Top             =   3840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   6240
      Top             =   5040
      Width           =   4095
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Initializing Chat Program..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   5160
      Width           =   6135
   End
   Begin VB.Image FreakPic 
      Height          =   480
      Left            =   2040
      Picture         =   "frmFreakChatSZ.frx":0000
      Top             =   6840
      Width           =   480
   End
   Begin VB.Image BillPic 
      Height          =   480
      Left            =   1080
      Picture         =   "frmFreakChatSZ.frx":0C42
      Top             =   6840
      Width           =   480
   End
   Begin VB.Image AvPic 
      Height          =   480
      Left            =   240
      Picture         =   "frmFreakChatSZ.frx":1886
      Top             =   6840
      Width           =   480
   End
   Begin VB.Image InvicPic 
      Height          =   480
      Left            =   3240
      Picture         =   "frmFreakChatSZ.frx":24CA
      Top             =   6840
      Width           =   480
   End
   Begin VB.Image imgSZPic 
      Height          =   495
      Left            =   4080
      Top             =   120
      Width           =   495
   End
   Begin VB.Shape shpTimeStampBack 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   6360
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmSZChat"
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

Public GameDataNum As Integer
Public OppIndex As Integer
Public Opponent As String
Public IsOpponent As Boolean 'when false, means I'm the host
Public GameID As Integer
Public RemoteGameID As Integer
Dim MyScore As Integer





Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1


'game specific:
Dim IStarted As Boolean
Dim IAmAway As Boolean
Dim AwayCount As Integer
Dim DoNotDisturb As Boolean

Dim PackTime As String
Dim PackDate As String

Dim Msg2Title As String
Dim MQuestion As String
Dim OpLevel As Integer

Dim TimeStamp As Boolean

Dim PingCheck1 As String
Dim PingCheck2 As String
Dim NameCheck As String

Public Sub StartGame()

If IsOpponent Then
    'tell the other guy i am here
    SendIt "HI", "Ready to chat!"
    SendIt "CN", ""
    tmrDuration.Enabled = True
    tmrPingCheck.Enabled = True
Else

End If

End Sub

Public Sub GameInterprit(Txt As String, IndexFrom As Integer, WhoFrom As String)

lblWho = WhoFrom

'gets the stuff
'(244)(244)(244)(245)[CODE](245)[PARAMS](245)(243)(243)(243)

e = InStr(1, Txt, Chr(245))
f = InStr(e + 1, Txt, Chr(245))
Debug.Print TheWord
If e > 0 And f > e And f > 0 Then
    'code
    a$ = Mid(Txt, e + 1, f - e - 1)
    e = f
    f = InStr(e + 1, Txt, Chr(245))
    
    If e > 0 And f > e And f > 0 Then
        'params
        p$ = Mid(Txt, e + 1, f - e - 1)
        'decode the encoded shtuff
    End If
End If

'GAME SPECIFIC PACKETS

If a$ = "HI" Then
    IStarted = True
    lblStatus = p$
End If

If a$ = "AB" Then
    AddText "~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    AddText "Private Chat 2.0 - SillyZone Special Edition"
    AddText "* Private Chat designed by Freaky, made for Avatar-X's Server Assistant Client"
    AddText "* Private Chat is Copyright 2001 JSHC Productions"
    AddText "* Server Assistant Client is Copyright 2001 CyberWyre"
    AddText "~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
End If

If a$ = "AW" Then
    lblStatusAway = WhoFrom & " is Away"
    lblStatusAway.Visible = True
    lblStatus = WhoFrom & " is now in Away Mode."
    AddText ">>> " & WhoFrom & " is currently away."
End If

If a$ = "AX" Then
    lblStatusAway.Visible = False
    lblStatus = WhoFrom & " has returned."
    AddText ">>> " & WhoFrom & " is back in the chat."
End If

If a$ = "BU" Then
    If DoNotDisturb = True Then
        SendIt "BF", ""
    ElseIf DoNotDisturb = False Then
        Beep
        AddText "*** " & WhoFrom & " is getting your attention."
    End If
End If

If a$ = "BF" Then
    AddText "*** " & WhoFrom & " does not wish to be beeped! No beep was sent."
End If

If a$ = "CB" Then
    Timer1.Enabled = False
    txtDisplay.ForeColor = &H80000005
End If

If a$ = "DT" Then
    AddText "<<<  It is now " & Time & " on " & Date & " >>>"
End If

If a$ = "EM" Then
    AddText "<> " & WhoFrom & " " & p$
End If

If a$ = "FL" Then
    AddText "*** " & WhoFrom & " has left the chat."
End If

If a$ = "MI" Then
    If DoNotDisturb = True Then
        SendIt "MF", ""
    ElseIf DoNotDisturb = False Then
        MessBox (p$)
    End If
End If

If a$ = "MF" Then
    AddText "*** " & WhoFrom & " does not wish to get messages! No message was sent."
End If

If a$ = "MA" Then
    AddText "<<< " & p$ & " >>>"
End If

If a$ = "MB" Then
    AddText ">>> IMPORTANT: " & p$ & " <<<"
End If

If a$ = "TD" Then
    lblStatusAway = WhoFrom & " is in DND Mode"
    lblStatusAway.Visible = True
    lblStatus = WhoFrom & " is now in DND Mode."
    AddText ">>> " & WhoFrom & " has activated DND Mode."
End If

If a$ = "FD" Then
    lblStatusAway.Visible = False
End If

If a$ = "CN" Then
    SendIt "NS", WhoFrom
    AddText "*** " & WhoFrom & " has entered the chat."
    tmrDuration.Enabled = True
    tmrPingCheck.Enabled = True

    If lblWho = "Avatar-X" Then
        imgSZPic.Picture = AvPic.Picture
    ElseIf lblWho = "Rabid Llama" Then
        imgSZPic.Picture = BillPic.Picture
    ElseIf lblWho = "Freaky" Then
        imgSZPic.Picture = FreakPic.Picture
    ElseIf lblWho = "Invictus" Then
        imgSZPic.Picture = InvicPic.Picture
    End If
End If

If a$ = "NS" Then
    txtNickName = p$
    NameCheck = p$
    SendIt "NO", WhoFrom
    AddText "*** " & WhoFrom & " has entered the chat."
    
    If lblWho = "Avatar-X" Then
        imgSZPic.Picture = AvPic.Picture
    ElseIf lblWho = "Rabid Llama" Then
        imgSZPic.Picture = BillPic.Picture
    ElseIf lblWho = "Freaky" Then
        imgSZPic.Picture = FreakPic.Picture
    ElseIf lblWho = "Invictus" Then
        imgSZPic.Picture = InvicPic.Picture
    End If
    
    If NameCheck = "Avatar-X" Then
        OpLevel = 1
    ElseIf NameCheck = "Freaky" Then
        OpLevel = 1
    ElseIf NameCheck = "Invictus" Then
        OpLevel = 2
    ElseIf NameCheck = "BillDoor" Then
        OpLevel = 3
    ElseIf NameCheck = "Sky" Then
        OpLevel = 4
    Else
        OpLevel = 100
    End If
End If

If a$ = "NO" Then
    txtNickName = p$
    NameCheck = p$

    If NameCheck = "Avatar-X" Then
        OpLevel = 1
    ElseIf NameCheck = "Freaky" Then
        OpLevel = 1
    ElseIf NameCheck = "Invictus" Then
        OpLevel = 2
    ElseIf NameCheck = "BillDoor" Then
        OpLevel = 3
    ElseIf NameCheck = "Sky" Then
        OpLevel = 4
    Else
        OpLevel = 100
    End If

End If

If a$ = "KU" Then
    Unload Me
    MessBox "You have been kicked from the chat!"
End If

If a$ = "DN" Then
    AddText "* " & WhoFrom & " has changed his nickname to " & p$ & "."
End If

If a$ = "LC" Then
    OpLevel = p$
    AddText "*** Your operator level has been set to " & p$ & " by " & WhoFrom & "."
End If

If a$ = "QA" Then
    Msg2Title = "qanswer"
    MsgBox2 p$, "Question from " & WhoFrom, , True, "Answer"
End If

If a$ = "QB" Then
    AddText "[*?*] " & WhoFrom & " has replied: " & p$ & " , to the question: " & MQuestion & "."
End If

If a$ = "SM" Then
    Msg2Title = ""
    MsgBox2 p$, "Error 201", , False, "Yes", "No"
End If

If a$ = "TT" Then
    AddText p$
    lblStatus = "New messages have been received!"
End If

If a$ = "PS" Then
    SendIt "PT", ""
    AddText "--  " & WhoFrom & " is verifying his connection to you."
End If

If a$ = "PC" Then
    SendIt "PO", Time
End If

If a$ = "PO" Then
    PingCheck1 = p$
End If

If a$ = "PT" Then
    PackTime = Time
    PackDate = Date
End If


'standard messages
If a$ = "QU" Then
    tmrDuration.Enabled = False
    tmrPingCheck.Enabled = False
    AddText "*** " & WhoFrom & " has left the chat."
    lblStatus = WhoFrom & " has left the chat."
    Msg2Title = "quit"
    MsgBox2 WhoFrom & " has left the chat. Do you wish to leave this Private Chat too?", "Quit", , False, "Yes", "No"
End If

If a$ = "N!" Then
    MessBox "User has declined your request for private chat."
    Unload Me
End If

End Sub

Private Sub SendIt(Cde As String, Params As String)

SendGamePacket Cde, Params, OppIndex, GameID, RemoteGameID

End Sub

Sub VerParams()
    If txtSay.Text = "about" Then
        SendIt "AB", ""
        AddText "~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        AddText "Private Chat 2.0 - SillyZone Special Edition"
        AddText "* Private Chat designed by Freaky, made for Avatar-X's Server Assistant Client"
        AddText "* Private Chat is Copyright 2001 JSHC Productions"
        AddText "* Server Assistant Client is Copyright 2001 CyberWyre"
        AddText "~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        txtSay.Text = ""
    ElseIf txtSay.Text = "logtofile" Then
        pcl$ = App.Path & "\pchatlog.log"
        h = FreeFile
        Open pcl$ For Append As h
            Print #h, "Date: " & Date & "  Time: " & Time
            Print #h, "Private Chat with: " & lblWho
            Print #h, "Duration: " & lblTimeHour & ":" & lblTimeMinute & ":" & lblTimeSecond & " (HH:MM:SS)"
            Print #h, txtDisplay
            Print #h,
            Print #h,
        Close #h
        AddText "*** Chat has been successfully logged to " & App.Path & "\pchatlog.log."
        txtSay.Text = ""
    ElseIf txtSay.Text = "pingtest" Then
        SendIt "PS", ""
        AddText "--  The last packet " & lblWho & " received, was on " & PackDate & " at " & PackTime & "."
        txtSay.Text = ""
    ElseIf txtSay.Text = "helpme" Then
        AddText "~~~  Private Chat Parameters ~~~"
        AddText "about : Displays Private Chat info."
        AddText "logtofile : Logs the entire chat into pchatlog.log."
        AddText "pingtest : Checks if connection with opposing user is active. (Check twice)"
        AddText "helpme : Displays help."
        AddText "~~~~~~~~~~~~~~~~~~~~~~~~"
        txtSay.Text = ""
    ElseIf txtSay.Text = "openextra" Then
        shpExtra.Visible = True
        lblExtra.Visible = True
        txtSay.Text = ""
    Else
        If Not txtSay.Text = "" Then
            SendIt "TT", txtNickName & ": " & txtSay.Text
            AddText txtNickName & ": " & txtSay.Text
            txtSay.Text = ""
        Else
            lblStatus = "You can't send an empty message!"
        End If
    End If

End Sub

Sub CheckAway()
If IAmAway = True Then
    SendIt "AX", ""
    AddText "*** " & lblWho & " has been informed that you are back."
    Me.Caption = "Private Chat 2 - " & lblWho
    AwayCount = 0
    IAmAway = False
ElseIf IAmAway = False Then
    SendIt "AW", ""
    AddText "*** " & lblWho & " has been informed that you are away."
    Me.Caption = "Private Chat 2 - " & lblWho & " - Away Mode"
    lblStatus = "You are currently in Away Mode..."
    IAmAway = True
End If

End Sub

Sub Msgbox2Check()

On Error Resume Next
If Msg2Title = "message" Then
    SendIt "MI", msgboxDefault
    AddText "*** Sending Message Box: [" & msgboxDefault & "]..."
ElseIf Msg2Title = "action" Then
    SendIt "EM", msgboxDefault
    AddText "<> " & txtNickName & " " & msgboxDefault
ElseIf Msg2Title = "question" Then
    SendIt "QA", msgboxDefault
    AddText "[*?*] Sent question: " & msgboxDefault
    MQuestion = msgboxDefault
ElseIf Msg2Title = "qanswer" Then
    SendIt "QB", msgboxDefault
    AddText "[*?*] You have replied to " & lblWho & "'s question."
ElseIf Msg2Title = "goodmessage" Then
    SendIt "MA", msgboxDefault
    AddText "<<< " & msgboxDefault & " >>>"
ElseIf Msg2Title = "state" Then
    SendIt "SM", msgboxDefault
    AddText "*** Statement sent: " & msgboxDefault
ElseIf Msg2Title = "levelchange" Then
    SendIt "LC", msgboxDefault
    AddText "*** " & lblWho & "'s operator level has been set to " & msgboxDefault & "."
ElseIf Msg2Title = "changename" Then
    If Not msgboxDefault = txtNickName And Not msgboxDefault = "" And Not msgboxDefault = lblWho Then
        If InStr(1, msgboxDefault, "Avatar-X", vbTextCompare) Then
            If NameCheck = "Avatar-X" Then
                GoTo ChangeMyName
            Else
                AddText "* You are not allowed to change your name to " & msgboxDefault
                GoTo EndChange
            End If
        ElseIf InStr(1, msgboxDefault, "Freaky", vbTextCompare) Then
            If NameCheck = "Freaky" Then
                GoTo ChangeMyName
            Else
                AddText "* You are not allowed to change your name to " & msgboxDefault
                GoTo EndChange
            End If
        ElseIf InStr(1, msgboxDefault, "Invictus", vbTextCompare) Then
            If NameCheck = "Invictus" Then
                GoTo ChangeMyName
            Else
                AddText "* You are not allowed to change your name to " & msgboxDefault
                GoTo EndChange
            End If
        ElseIf InStr(1, msgboxDefault, "BillDoor", vbTextCompare) Then
            If NameCheck = "BillDoor" Then
                GoTo ChangeMyName
            Else
                AddText "* You are not allowed to change your name to " & msgboxDefault
                GoTo EndChange
            End If
        End If
ChangeMyName:
        SendIt "DN", msgboxDefault
        txtNickName = msgboxDefault
        AddText "* You have changed your nickname to " & msgboxDefault & "."
    End If
EndChange:
End If

picMsgBox.Visible = False
txtSay.SetFocus

If Msg2Title = "quit" Then
    Unload Me
End If

End Sub

Sub StartMove()
    
    If Me.WindowState = 2 Then Exit Sub
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, _
        HTCAPTION, 0&


End Sub

Sub StartMove2()
    
    ReleaseCapture
    SendMessage picMsgBox.hwnd, WM_NCLBUTTONDOWN, _
        HTCAPTION, 0&

End Sub

Sub PingCheckTest()

If PingCheck2 = PingCheck1 Then
    AddText "-/- You have lost your connection to " & lblWho & "."
    AddText "-/- Please verify which side the disconnection has occurred."
    tmrPingCheck.Enabled = False
Else
    PingCheck2 = PingCheck1
End If

End Sub

Sub OpError()

AddText "*** Your operator level is not cleared to perform this action."

End Sub
Private Sub flashAction_Timer()
If lblAction = "Action" Then
    lblAction = ""
    flashAction.Interval = 100
ElseIf lblAction = "" Then
    lblAction = "A"
    shpAction.BorderColor = &H8000&
ElseIf lblAction = "A" Then
    lblAction = "Ac"
    shpAction.BorderColor = &HFF00&
ElseIf lblAction = "Ac" Then
    lblAction = "Act"
    shpAction.BorderColor = &H8000&
ElseIf lblAction = "Act" Then
    lblAction = "Acti"
    shpAction.BorderColor = &HFF00&
ElseIf lblAction = "Acti" Then
    lblAction = "Actio"
    shpAction.BorderColor = &H8000&
ElseIf lblAction = "Actio" Then
    lblAction = "Action"
    flashAction.Interval = 1000
    shpAction.BorderColor = &HFF00&
End If

End Sub

Private Sub flashAway_Timer()
If lblAway = "Away" Then
    lblAway = ""
    flashAway.Interval = 100
ElseIf lblAway = "" Then
    lblAway = "A"
    shpAway.BorderColor = &H8000&
ElseIf lblAway = "A" Then
    lblAway = "Aw"
    shpAway.BorderColor = &HFF00&
ElseIf lblAway = "Aw" Then
    lblAway = "Awa"
    shpAway.BorderColor = &H8000&
ElseIf lblAway = "Awa" Then
    lblAway = "Away"
    flashAway.Interval = 1000
    shpAway.BorderColor = &HFF00&
End If

End Sub

Private Sub flashBeep_Timer()
If lblBeep = "Beep" Then
    lblBeep = ""
    flashBeep.Interval = 100
ElseIf lblBeep = "" Then
    lblBeep = "B"
    shpBeep.BorderColor = &H8000&
ElseIf lblBeep = "B" Then
    lblBeep = "Be"
    shpBeep.BorderColor = &HFF00&
ElseIf lblBeep = "Be" Then
    lblBeep = "Bee"
    shpBeep.BorderColor = &H8000&
ElseIf lblBeep = "Bee" Then
    lblBeep = "Beep"
    flashBeep.Interval = 1000
    shpBeep.BorderColor = &HFF00&
End If

End Sub

Private Sub flashCancel_Timer()
If shpMsgboxCancel.BorderColor = &HFF00& Then
shpMsgboxCancel.BorderColor = &H8000&
ElseIf shpMsgboxCancel.BorderColor = &H8000& Then
shpMsgboxCancel.BorderColor = &HFF00&
End If

End Sub

Private Sub flashExtra_Timer()
If lblExtra = "Extra" Then
    lblExtra = ""
    flashExtra.Interval = 100
    shpExtra.BorderColor = &HFFFF00
ElseIf lblExtra = "" Then
    lblExtra = "E"
    shpExtra.BorderColor = &HFF00&
ElseIf lblExtra = "E" Then
    lblExtra = "Ex"
    shpExtra.BorderColor = &HFFFF00
ElseIf lblExtra = "Ex" Then
    lblExtra = "Ext"
    shpExtra.BorderColor = &HFF00&
ElseIf lblExtra = "Ext" Then
    lblExtra = "Extr"
    shpExtra.BorderColor = &HFFFF00
ElseIf lblExtra = "Extr" Then
    lblExtra = "Extra"
    flashExtra.Interval = 1000
    shpExtra.BorderColor = &HFF00&
End If

End Sub

Private Sub flashLeave_Timer()
If lblLeave = "Leave" Then
    lblLeave = ""
    flashLeave.Interval = 100
    shpLeave.BorderColor = &H8000&
ElseIf lblLeave = "" Then
    lblLeave = "L"
    shpLeave.BorderColor = &HFF00&
ElseIf lblLeave = "L" Then
    lblLeave = "Le"
    shpLeave.BorderColor = &H8000&
ElseIf lblLeave = "Le" Then
    lblLeave = "Lea"
    shpLeave.BorderColor = &HFF00&
ElseIf lblLeave = "Lea" Then
    lblLeave = "Leav"
    shpLeave.BorderColor = &H8000&
ElseIf lblLeave = "Leav" Then
    lblLeave = "Leave"
    flashLeave.Interval = 1000
    shpLeave.BorderColor = &HFF00&
End If

End Sub

Private Sub flashMessage_Timer()
If lblMessage = "Message" Then
    lblMessage = ""
    flashMessage.Interval = 100
    shpMessage.BorderColor = &H8000&
ElseIf lblMessage = "" Then
    lblMessage = "M"
    shpMessage.BorderColor = &HFF00&
ElseIf lblMessage = "M" Then
    lblMessage = "Me"
    shpMessage.BorderColor = &H8000&
ElseIf lblMessage = "Me" Then
    lblMessage = "Mes"
    shpMessage.BorderColor = &HFF00&
ElseIf lblMessage = "Mes" Then
    lblMessage = "Mess"
    shpMessage.BorderColor = &H8000&
ElseIf lblMessage = "Mess" Then
    lblMessage = "Messa"
    shpMessage.BorderColor = &HFF00&
ElseIf lblMessage = "Messa" Then
    lblMessage = "Messag"
    shpMessage.BorderColor = &H8000&
ElseIf lblMessage = "Messag" Then
    lblMessage = "Message"
    flashMessage.Interval = 1000
    shpMessage.BorderColor = &HFF00&
End If

End Sub

Private Sub flashOk_Timer()
If shpMsgboxOK.BorderColor = &HFF00& Then
shpMsgboxOK.BorderColor = &H8000&
ElseIf shpMsgboxOK.BorderColor = &H8000& Then
shpMsgboxOK.BorderColor = &HFF00&
End If

End Sub

Private Sub flashSend_Timer()
If lblSend = "Send" Then
    lblSend = ""
    flashSend.Interval = 100
ElseIf lblSend = "" Then
    lblSend = "S"
    shpSend.BorderColor = &H8000&
ElseIf lblSend = "S" Then
    lblSend = "Se"
    shpSend.BorderColor = &HFF00&
ElseIf lblSend = "Se" Then
    lblSend = "Sen"
    shpSend.BorderColor = &H8000&
ElseIf lblSend = "Sen" Then
    lblSend = "Send"
    flashSend.Interval = 1000
    shpSend.BorderColor = &HFF00&
End If

End Sub

Private Sub flashTime_Timer()
If lblTime = "Time" Then
    lblTime = ""
    flashTime.Interval = 100
ElseIf lblTime = "" Then
    lblTime = "T"
    shpTime.BorderColor = &H8000&
ElseIf lblTime = "T" Then
    lblTime = "Ti"
    shpTime.BorderColor = &HFF00&
ElseIf lblTime = "Ti" Then
    lblTime = "Tim"
    shpTime.BorderColor = &H8000&
ElseIf lblTime = "Tim" Then
    lblTime = "Time"
    flashTime.Interval = 1000
    shpTime.BorderColor = &HFF00&
End If

End Sub

Private Sub Form_Load()
DoNotDisturb = False
IAmAway = False

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

StartMove


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
flashSend.Enabled = False
lblSend = "Send"
shpSend.BorderColor = &HFF00&
flashSend.Interval = 100

flashBeep.Enabled = False
lblBeep = "Beep"
shpBeep.BorderColor = &HFF00&
flashBeep.Interval = 100

flashMessage.Enabled = False
lblMessage = "Message"
shpMessage.BorderColor = &HFF00&
flashMessage.Interval = 100

flashAction.Enabled = False
lblAction = "Action"
shpAction.BorderColor = &HFF00&
flashAction.Interval = 100

flashTime.Enabled = False
lblTime = "Time"
shpTime.BorderColor = &HFF00&
flashTime.Interval = 100

flashAway.Enabled = False
lblAway = "Away"
shpAway.BorderColor = &HFF00&
flashAway.Interval = 100

flashLeave.Enabled = False
lblLeave = "Leave"
shpLeave.BorderColor = &HFF00&
flashLeave.Interval = 100

flashExtra.Enabled = False
lblExtra = "Extra"
shpExtra.BorderColor = &HFF00&
flashExtra.Interval = 100

End Sub

Private Sub Form_Unload(Cancel As Integer)

SendIt "QU", ""
RemoveGameData GameDataNum

End Sub

Private Sub Image1_Click()
txtDisplay.Visible = True

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
StartMove

End Sub

Private Sub lblAction_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpAction.BorderColor = &H80FFFF
lblAction.ForeColor = &H80FFFF

End Sub

Private Sub lblAction_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
flashAction.Enabled = True

End Sub

Private Sub lblAction_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpAction.BorderColor = &HFF00&
lblAction.ForeColor = &HFF00&

Msg2Title = "action"
MsgBox2 "Enter your action. Start off with a verb (e.g says hello.).", "Send Emotion", , True

End Sub

Private Sub lblAway_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpAway.BorderColor = &H80FFFF
lblAway.ForeColor = &H80FFFF

End Sub

Private Sub lblAway_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
flashAway.Enabled = True

End Sub

Private Sub lblAway_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpAway.BorderColor = &HFF00&
lblAway.ForeColor = &HFF00&

CheckAway

End Sub

Private Sub lblBeep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpBeep.BorderColor = &H80FFFF
lblBeep.ForeColor = &H80FFFF

End Sub

Private Sub lblBeep_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
flashBeep.Enabled = True

End Sub

Private Sub lblBeep_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpBeep.BorderColor = &HFF00&
lblBeep.ForeColor = &HFF00&

SendIt "BU", ""
AddText "*** Beeping " & lblWho & "'s client..."

End Sub

Private Sub lblChangeName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpChangeName.BorderColor = &H80FFFF
lblChangeName.ForeColor = &H80FFFF

End Sub

Private Sub lblChangeName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpChangeName.BorderColor = &HFF00&
lblChangeName.ForeColor = &HFF00&

NameChange:
Msg2Title = "changename"
MsgBox2 "Enter your new name:", "Change Name", txtNickName, True, "Change"

End Sub

Private Sub lblDND_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpDND.BorderColor = &H80FFFF
lblDND.ForeColor = &H80FFFF

End Sub

Private Sub lblDND_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpDND.BorderColor = &HFF00&
lblDND.ForeColor = &HFF00&

If DoNotDisturb = True Then
    DoNotDisturb = False
    shpDNDBack.Visible = False
    SendIt "FD", ""
ElseIf DoNotDisturb = False Then
    DoNotDisturb = True
    shpDNDBack.Visible = True
    SendIt "TD", ""
End If

End Sub

Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
exit1.BorderColor = &H80FFFF
exit2.BorderColor = &H80FFFF

End Sub

Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
exit1.BorderColor = &HFF00&
exit2.BorderColor = &HFF00&

Unload Me

End Sub

Private Sub lblExtra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpExtra.BorderColor = &H80FFFF
lblExtra.ForeColor = &H80FFFF

End Sub

Private Sub lblExtra_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
flashExtra.Enabled = True

End Sub

Private Sub lblExtra_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpExtra.BorderColor = &HFF00&
lblExtra.ForeColor = &HFF00&

If picExtras.Visible = False Then
    lineExtra.Visible = True
    picExtras.Visible = True
    Me.Width = 9990
ElseIf picExtras.Visible = True Then
    lineExtra.Visible = False
    picExtras.Visible = False
    Me.Width = 8070
End If

End Sub

Private Sub lblGoodMessage_Click()
Msg2Title = "goodmessage"
MsgBox2 "Enter your message:", "Message 2", "IMPORTANT: ", True

End Sub

Private Sub lblKick_Click()
If OpLevel <= 2 Then

SendIt "KU", ""
AddText "*** Kicking " & lblWho & " from the chat..."

Else
    OpError
End If
End Sub

Private Sub lblLeave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpLeave.BorderColor = &HFF&
lblLeave.ForeColor = &HFF&

End Sub

Private Sub lblLeave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
flashLeave.Enabled = True

End Sub

Private Sub lblLeave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpLeave.BorderColor = &HFF00&
lblLeave.ForeColor = &HFF00&

Msg2Title = "quit"
MsgBox2 "Are you sure you want to leave?", "Leave", , False, "Yes", "No"

End Sub

Private Sub lblLeaveNo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpLeaveNo.BorderColor = &H80FFFF
lblLeaveNo.ForeColor = &H80FFFF

End Sub

Private Sub lblLeaveNo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpLeaveNo.BorderColor = &HFF00&
lblLeaveNo.ForeColor = &HFF00&

lblLeaveAsk.Visible = False
lblLeaveYes.Visible = False
shpLeaveYes.Visible = False
lblLeaveNo.Visible = False
shpLeaveNo.Visible = False

End Sub

Private Sub lblLeaveYes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpLeaveYes.BorderColor = &H80FFFF
lblLeaveYes.ForeColor = &H80FFFF

End Sub

Private Sub lblLeaveYes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpLeaveYes.BorderColor = &HFF00&
lblLeaveYes.ForeColor = &HFF00&

Unload Me

End Sub

Private Sub lblLevel_Click()
If OpLevel = 1 Then

Msg2Title = "levelchange"
MsgBox2 "Enter the operator level for " & lblWho & ":", "Operator Level", "100", True, "Change"

Else
    OpError
End If
End Sub

Private Sub lblLogToFile_Click()
pcl$ = App.Path & "\pchatlog.log"
h = FreeFile
Open pcl$ For Append As h
    Print #h, "Date: " & Date & "  Time: " & Time
    Print #h, "Private Chat with: " & lblWho
    Print #h, "Duration: " & lblTimeHour & ":" & lblTimeMinute & ":" & lblTimeSecond & " (HH:MM:SS)"
    Print #h, txtDisplay
    Print #h,
    Print #h,
Close #h
Msg2Title = ""
MsgBox2 "Chat has been successfully logged to " & App.Path & "\pchatlog.log.", "Log Chat", "", False

End Sub

Private Sub lblMessage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpMessage.BorderColor = &H80FFFF
lblMessage.ForeColor = &H80FFFF

End Sub

Private Sub lblMessage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
flashMessage.Enabled = True

End Sub

Private Sub lblMessage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpMessage.BorderColor = &HFF00&
lblMessage.ForeColor = &HFF00&

Msg2Title = "message"
MsgBox2 "Enter your message for " & lblWho & ".", "Send Message Box", , True

End Sub

Private Sub lblMinimize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
min1.BorderColor = &H80FFFF
min2.BorderColor = &H80FFFF
min3.BorderColor = &H80FFFF

End Sub

Private Sub lblMinimize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
min1.BorderColor = &HFF00&
min2.BorderColor = &HFF00&
min3.BorderColor = &HFF00&

If Me.Height = 720 Then
    Me.Height = 6526
    
    Y1 = min1.Y1
    X1 = min1.X1
    X2 = min1.X2
    Y2 = min1.Y2
    
    y3 = min2.Y2
    y4 = min2.Y1
    x3 = min2.X2
    
    min1.Y1 = y4
    min1.Y2 = y4
    
    x4 = min2.X1
    min2.X1 = min2.X2
    min2.X2 = x4
    
    x4 = min3.X1
    min3.X1 = min3.X2
    min3.X2 = x4
 
Else
    Me.Height = 720

    Y1 = min1.Y1
    X1 = min1.X1
    X2 = min1.X2
    Y2 = min1.Y2
    
    y3 = min2.Y2
    x3 = min2.X2
    
    min1.Y1 = y3
    min1.Y2 = y3
    
    x4 = min2.X1
    min2.X1 = min2.X2
    min2.X2 = x4
    
    x4 = min3.X1
    min3.X1 = min3.X2
    min3.X2 = x4

End If

End Sub

Private Sub lblMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
StartMove

End Sub

Private Sub lblQuestion_Click()
If OpLevel <= 3 Then

Msg2Title = "question"
MsgBox2 "Enter your question:", "Question", , True

Else
    OpError
End If
End Sub

Private Sub lblSend_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpSend.BorderColor = &H80FFFF
'shpSendBack.FillColor = &HFF00&
lblSend.ForeColor = &H80FFFF

End Sub

Private Sub lblSend_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
flashSend.Enabled = True

End Sub

Private Sub lblSend_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpSend.BorderColor = &HFF00&
'shpSendBack.FillColor = &H0&
lblSend.ForeColor = &HFF00&

VerParams

End Sub

Private Sub lblState_Click()
If OpLevel <= 3 Then

Msg2Title = "state"
MsgBox2 "Enter your statement:", "State", , True

Else
    OpError
End If
End Sub

Private Sub lblStatus_Click()
If StatusMove.Interval = 10 Then
    StatusMove.Interval = 100
ElseIf StatusMove.Interval = 100 Then
    StatusMove.Interval = 1000
ElseIf StatusMove.Interval = 1000 Then
    StatusMove.Interval = 0
    lblStatus.Left = 360
ElseIf StatusMove.Interval = 0 Then
    StatusMove.Interval = 10
End If

End Sub

Private Sub lblTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpTime.BorderColor = &H80FFFF
lblTime.ForeColor = &H80FFFF

End Sub

Private Sub lblTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
flashTime.Enabled = True

End Sub

Private Sub lblTime_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpTime.BorderColor = &HFF00&
lblTime.ForeColor = &HFF00&

SendIt "DT", ""
AddText "<<<  It is now " & Time & " on " & Date & "  >>>"

End Sub

Private Sub lblTimeHour_Change()
lblTimeHour.Caption = Format$(lblTimeHour, "00")

End Sub

Private Sub lblTimeMinute_Change()
lblTimeMinute.Caption = Format$(lblTimeMinute, "00")

End Sub

Private Sub lblTimeSecond_Change()
lblTimeSecond.Caption = Format$(lblTimeSecond, "00")

End Sub

Private Sub lblTimeStamp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpTimeStamp.BorderColor = &H80FFFF
lblTimeStamp.ForeColor = &H80FFFF

End Sub

Private Sub lblTimeStamp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpTimeStamp.BorderColor = &HFF00&
lblTimeStamp.ForeColor = &HFF00&

If TimeStamp = True Then
    TimeStamp = False
    shpTimeStampBack.Visible = False
ElseIf TimeStamp = False Then
    TimeStamp = True
    shpTimeStampBack.Visible = True
End If

End Sub

Private Sub lblWho_Click()
If tmrWhoColour.Enabled = True Then
    tmrWhoColour.Enabled = False
    lblWho.ForeColor = &HFF00&
ElseIf tmrWhoColour.Enabled = False Then
    tmrWhoColour.Enabled = True
End If

End Sub

Private Sub msgboxCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpMsgboxCancel.BorderColor = &H80FFFF
msgboxCancel.ForeColor = &H80FFFF

End Sub

Private Sub msgboxCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
flashCancel.Enabled = True
End Sub

Private Sub msgboxCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next
shpMsgboxCancel.BorderColor = &HFF00&
msgboxCancel.ForeColor = &HFF00&

picMsgBox.Visible = False
txtSay.SetFocus

End Sub

Private Sub msgboxDefault_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    Msgbox2Check
End If

End Sub

Private Sub msgboxOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpMsgboxOK.BorderColor = &H80FFFF
msgboxOK.ForeColor = &H80FFFF

End Sub

Private Sub msgboxOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
flashOk.Enabled = True

End Sub

Private Sub msgboxOK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
shpMsgboxOK.BorderColor = &HFF00&
msgboxOK.ForeColor = &HFF00&

Msgbox2Check

End Sub

Private Sub msgboxTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
StartMove2

End Sub

Private Sub picExtras_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
flashSend.Enabled = False
lblSend = "Send"
shpSend.BorderColor = &HFF00&
flashSend.Interval = 100

flashBeep.Enabled = False
lblBeep = "Beep"
shpBeep.BorderColor = &HFF00&
flashBeep.Interval = 100

flashMessage.Enabled = False
lblMessage = "Message"
shpMessage.BorderColor = &HFF00&
flashMessage.Interval = 100

flashAction.Enabled = False
lblAction = "Action"
shpAction.BorderColor = &HFF00&
flashAction.Interval = 100

flashTime.Enabled = False
lblTime = "Time"
shpTime.BorderColor = &HFF00&
flashTime.Interval = 100

flashAway.Enabled = False
lblAway = "Away"
shpAway.BorderColor = &HFF00&
flashAway.Interval = 100

flashLeave.Enabled = False
lblLeave = "Leave"
shpLeave.BorderColor = &HFF00&
flashLeave.Interval = 100

flashExtra.Enabled = False
lblExtra = "Extra"
shpExtra.BorderColor = &HFF00&
flashExtra.Interval = 100

End Sub

Private Sub picMsgBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
StartMove2

End Sub

Private Sub picMsgBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
flashOk.Enabled = False
flashCancel.Enabled = False

shpMsgboxOK.BorderColor = &HFF00&
shpMsgboxCancel.BorderColor = &HFF00&

End Sub

Private Sub StatusMove_Timer()
lblStatus.Left = lblStatus.Left - 60

If lblStatus.Left < -5880 Then
lblStatus.Left = 6240
End If

End Sub

Private Sub tmrDuration_Timer()
lblTimeSecond.Caption = lblTimeSecond.Caption + 1

If lblTimeSecond = 60 Then
lblTimeMinute = lblTimeMinute + 1
lblTimeSecond = 0
End If

If lblTimeMinute = 60 Then
lblTimeHour = lblTimeHour + 1
lblTimeMinute = 0
End If

End Sub

Private Sub tmrPingCheck_Timer()
SendIt "PC", ""

If tmrPingCheck.Interval = 5000 Then
    PingCheckTest
    tmrPingCheck.Interval = 10000
Else
    tmrPingCheck.Interval = 5000
End If

End Sub

Private Sub tmrWhoColour_Timer()
If lblWho.ForeColor = &HFF& Then
    lblWho.ForeColor = &H80FF&
ElseIf lblWho.ForeColor = &H80FF& Then
    lblWho.ForeColor = &HFFFF&
ElseIf lblWho.ForeColor = &HFFFF& Then
    lblWho.ForeColor = &HFF00&
ElseIf lblWho.ForeColor = &HFF00& Then
    lblWho.ForeColor = &HFFFF00
ElseIf lblWho.ForeColor = &HFFFF00 Then
    lblWho.ForeColor = &HFF0000
ElseIf lblWho.ForeColor = &HFF0000 Then
    lblWho.ForeColor = &HFF00FF
ElseIf lblWho.ForeColor = &HFF00FF Then
    lblWho.ForeColor = &HFF&
End If

End Sub

Private Sub txtDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
flashSend.Enabled = False
lblSend = "Send"
shpSend.BorderColor = &HFF00&
flashSend.Interval = 200

flashBeep.Enabled = False
lblBeep = "Beep"
shpBeep.BorderColor = &HFF00&
flashBeep.Interval = 200

flashMessage.Enabled = False
lblMessage = "Message"
shpMessage.BorderColor = &HFF00&
flashMessage.Interval = 200

flashAction.Enabled = False
lblAction = "Action"
shpAction.BorderColor = &HFF00&
flashAction.Interval = 200

flashTime.Enabled = False
lblTime = "Time"
shpTime.BorderColor = &HFF00&
flashTime.Interval = 200

flashAway.Enabled = False
lblAway = "Away"
shpAway.BorderColor = &HFF00&
flashAway.Interval = 200

flashLeave.Enabled = False
lblLeave = "Leave"
shpLeave.BorderColor = &HFF00&
flashLeave.Interval = 100

flashExtra.Enabled = False
lblExtra = "Extra"
shpExtra.BorderColor = &HFF00&
flashExtra.Interval = 100

End Sub

Private Sub txtSay_Change()
lblStatus = "Conversation in progress..."
AwayCount = 0

End Sub

Private Sub txtSay_KeyPress(KeyAscii As Integer)
If IAmAway = True Then
    CheckAway
End If

If KeyAscii = 13 Then
    KeyAscii = 0
    VerParams
End If

End Sub

Sub AddText(Msg As String)

If TimeStamp = True Then
    ShortTime = Format(Time, "hh:mm:ss")
    txtDisplay = txtDisplay + vbCrLf + "[" & ShortTime & "] " + Msg
ElseIf TimeStamp = False Then
    txtDisplay = txtDisplay + vbCrLf + Msg
End If

txtDisplay.SelStart = Len(txtDisplay)

End Sub

Private Sub tmrAway_Timer()
AwayCount = AwayCount + 1
If AwayCount > 20000 Then AwayCount = 20000
If AwayCount >= 90 And IAmAway = False Then
    'AddText "*** Activating Auto-Away mode..."
    'IAmAway = False
    'CheckAway
End If

End Sub

Private Sub txtSay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
flashSend.Enabled = False
lblSend = "Send"
shpSend.BorderColor = &HFF00&
flashSend.Interval = 200

flashBeep.Enabled = False
lblBeep = "Beep"
shpBeep.BorderColor = &HFF00&
flashBeep.Interval = 200

flashMessage.Enabled = False
lblMessage = "Message"
shpMessage.BorderColor = &HFF00&
flashMessage.Interval = 200

flashAction.Enabled = False
lblAction = "Action"
shpAction.BorderColor = &HFF00&
flashAction.Interval = 200

flashTime.Enabled = False
lblTime = "Time"
shpTime.BorderColor = &HFF00&
flashTime.Interval = 200

flashAway.Enabled = False
lblAway = "Away"
shpAway.BorderColor = &HFF00&
flashAway.Interval = 200

flashLeave.Enabled = False
lblLeave = "Leave"
shpLeave.BorderColor = &HFF00&
flashLeave.Interval = 100

flashExtra.Enabled = False
lblExtra = "Extra"
shpExtra.BorderColor = &HFF00&
flashExtra.Interval = 100

End Sub

Function MsgBox2(CQuestion As String, Optional CTitle As String, Optional CDefault As String, Optional CTextEnable As Boolean, Optional COkLabel As String, Optional CCancelLabel As String) As String

On Error Resume Next
msgboxQuestion.Caption = CQuestion
msgboxTitle.Caption = CTitle
msgboxDefault.Text = CDefault
msgboxOK.Caption = COkLabel
msgboxCancel.Caption = CCancelLabel

If CTextEnable = True Then
    lineMsg1.Visible = True
    lineMsg2.Visible = True
    lineMsg3.Visible = True
    lineMsg4.Visible = True
    lineMsg5.Visible = True
    lineMsg6.Visible = True
    msgboxDefault.Visible = True
Else
    lineMsg1.Visible = False
    lineMsg2.Visible = False
    lineMsg3.Visible = False
    lineMsg4.Visible = False
    lineMsg5.Visible = False
    lineMsg6.Visible = False
    msgboxDefault.Visible = False
End If

If msgboxOK = "" Then
    msgboxOK.Caption = "OK"
End If
If msgboxCancel = "" Then
    msgboxCancel.Caption = "Cancel"
End If

picMsgBox.Visible = True
msgboxDefault.SetFocus
msgboxDefault.SelLength = Len(msgboxDefault)

End Function

