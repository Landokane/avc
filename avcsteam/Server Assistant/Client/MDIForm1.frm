VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Client"
   ClientHeight    =   6600
   ClientLeft      =   1080
   ClientTop       =   1470
   ClientWidth     =   11880
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":548A
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1980
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1140
      Top             =   1560
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6345
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Text            =   "Client"
            TextSave        =   "Client"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1774
            Text            =   "Server:"
            TextSave        =   "Server:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Upload Progress:"
            TextSave        =   "Upload Progress:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Download Progress:"
            TextSave        =   "Download Progress:"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "Users:"
            TextSave        =   "Users:"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Text            =   "HLDS:"
            TextSave        =   "HLDS:"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   300
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   179
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":AB1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B0B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B995
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":C26F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":C58B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":CE67
            Key             =   ""
            Object.Tag             =   "map"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":D743
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":E01D
            Key             =   ""
            Object.Tag             =   "exit"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":E8F7
            Key             =   ""
            Object.Tag             =   "web"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":F1D1
            Key             =   ""
            Object.Tag             =   "startscript"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":FAAB
            Key             =   ""
            Object.Tag             =   "namesincolour"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":FEFF
            Key             =   ""
            Object.Tag             =   "events"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":10353
            Key             =   ""
            Object.Tag             =   "timestamp"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":10C2D
            Key             =   ""
            Object.Tag             =   "hiddenmode"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":11081
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1195B
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":12235
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":12B0F
            Key             =   ""
            Object.Tag             =   "editclans"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":133E9
            Key             =   ""
            Object.Tag             =   "kickbans"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":13CC3
            Key             =   ""
            Object.Tag             =   "changepassword"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":14117
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":149F1
            Key             =   ""
            Object.Tag             =   "16"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":14E4D
            Key             =   ""
            Object.Tag             =   "salog"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":152A1
            Key             =   ""
            Object.Tag             =   "serverlog"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":156F5
            Key             =   ""
            Object.Tag             =   "messagesinwindow"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":15B49
            Key             =   ""
            Object.Tag             =   "compose"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":15F9D
            Key             =   ""
            Object.Tag             =   "serverinfo"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":16877
            Key             =   ""
            Object.Tag             =   "connectusers"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":17151
            Key             =   ""
            Object.Tag             =   "mailbox"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1746D
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":17D47
            Key             =   ""
            Object.Tag             =   "adminchat"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":18621
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":18EFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":197D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1A0AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1A989
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1B263
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1BB3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1C417
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1CCF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1D5CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1DEA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1E77F
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F059
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F933
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2020D
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":20AE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":213C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":21C9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":22575
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":22E4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":232A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":236F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":23A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":23E5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":242B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":24703
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":24B55
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":24FA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":253F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2584B
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":25C9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":260EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":26541
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":26993
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":26DE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":27237
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":27551
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":279A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":27DF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":28247
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":28699
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":28AEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":28F3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2938F
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":297E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":29C33
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2A085
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2A4D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2A929
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2AD7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2B1CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2B61F
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2BA71
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2BEC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2C315
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2C767
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2CBB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2D00B
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2D45D
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2D8AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2DD01
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2E153
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2E5A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2E9F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2ED11
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2F163
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2F5B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2FA07
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2FE59
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":302AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":306FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":30B4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":30FA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":313F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":31845
            Key             =   ""
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":31C97
            Key             =   ""
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":320E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3253B
            Key             =   ""
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3298D
            Key             =   ""
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":32DDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":33231
            Key             =   ""
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":33B0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":343E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":350BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":35999
            Key             =   ""
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":36273
            Key             =   ""
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":36B4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":37427
            Key             =   ""
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":37D01
            Key             =   ""
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":385DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":38EB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":39B8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3A869
            Key             =   ""
         EndProperty
         BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3B143
            Key             =   ""
         EndProperty
         BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3BA1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3C6F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage128 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3CFD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage129 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3DCAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage130 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3E585
            Key             =   ""
         EndProperty
         BeginProperty ListImage131 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3F25F
            Key             =   ""
         EndProperty
         BeginProperty ListImage132 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3FB39
            Key             =   ""
         EndProperty
         BeginProperty ListImage133 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":40413
            Key             =   ""
         EndProperty
         BeginProperty ListImage134 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":40CED
            Key             =   ""
         EndProperty
         BeginProperty ListImage135 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":415C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage136 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":41EA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage137 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4277B
            Key             =   ""
         EndProperty
         BeginProperty ListImage138 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":43055
            Key             =   ""
         EndProperty
         BeginProperty ListImage139 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4392F
            Key             =   ""
         EndProperty
         BeginProperty ListImage140 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":44209
            Key             =   ""
         EndProperty
         BeginProperty ListImage141 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":44AE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage142 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":457BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage143 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":46097
            Key             =   ""
         EndProperty
         BeginProperty ListImage144 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":46971
            Key             =   ""
         EndProperty
         BeginProperty ListImage145 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4724B
            Key             =   ""
         EndProperty
         BeginProperty ListImage146 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":47B25
            Key             =   ""
         EndProperty
         BeginProperty ListImage147 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":483FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage148 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":48CD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage149 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":499B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage150 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4A28D
            Key             =   ""
         EndProperty
         BeginProperty ListImage151 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4AB67
            Key             =   ""
         EndProperty
         BeginProperty ListImage152 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4B441
            Key             =   ""
         EndProperty
         BeginProperty ListImage153 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4BD1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage154 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4C5F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage155 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4CECF
            Key             =   ""
         EndProperty
         BeginProperty ListImage156 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4D7A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage157 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4E083
            Key             =   ""
         EndProperty
         BeginProperty ListImage158 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4E95D
            Key             =   ""
         EndProperty
         BeginProperty ListImage159 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4F237
            Key             =   ""
         EndProperty
         BeginProperty ListImage160 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4FB11
            Key             =   ""
         EndProperty
         BeginProperty ListImage161 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":503EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage162 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":50CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage163 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5159F
            Key             =   ""
         EndProperty
         BeginProperty ListImage164 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":51E79
            Key             =   ""
         EndProperty
         BeginProperty ListImage165 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":52753
            Key             =   ""
         EndProperty
         BeginProperty ListImage166 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5302D
            Key             =   ""
         EndProperty
         BeginProperty ListImage167 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":53907
            Key             =   ""
         EndProperty
         BeginProperty ListImage168 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":541E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage169 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":54ABB
            Key             =   ""
         EndProperty
         BeginProperty ListImage170 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":55395
            Key             =   ""
         EndProperty
         BeginProperty ListImage171 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":55C6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage172 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":56549
            Key             =   ""
         EndProperty
         BeginProperty ListImage173 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":56E23
            Key             =   ""
         EndProperty
         BeginProperty ListImage174 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":576FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage175 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":57FD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage176 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":588B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage177 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5918B
            Key             =   ""
         EndProperty
         BeginProperty ListImage178 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":59A65
            Key             =   ""
         EndProperty
         BeginProperty ListImage179 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5A33F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&Program"
      Begin VB.Menu mnuFileIn 
         Caption         =   "&Exit"
         Index           =   0
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuScripts 
      Caption         =   "&Scripts"
      Begin VB.Menu mnuScriptsIn 
         Caption         =   "&Edit Scripts..."
         Index           =   0
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuScriptsIn 
         Caption         =   "&Start Script..."
         Index           =   1
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "&Admin"
      Begin VB.Menu mnuAdminIn 
         Caption         =   "&Update EXE..."
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuAdminIn 
         Caption         =   "&Edit Users..."
         Index           =   1
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuAdminIn 
         Caption         =   "Edit &Kick-Ban List..."
         Index           =   2
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuAdminIn 
         Caption         =   "Edit &Server Info..."
         Index           =   3
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuAdminIn 
         Caption         =   "Edit &Clans..."
         Index           =   4
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuAdminIn 
         Caption         =   "Edit S&peech..."
         Index           =   5
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuAdminIn 
         Caption         =   "Edit RealPlayers..."
         Index           =   6
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuAdminIn 
         Caption         =   "Edit Web Info..."
         Index           =   7
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuAdminIn 
         Caption         =   "Edit General Info"
         Index           =   8
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuAdminIn 
         Caption         =   "Edit Events..."
         Index           =   9
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuAdminIn 
         Caption         =   "Edit Bad Word List..."
         Index           =   10
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuAdminIn 
         Caption         =   "File Manager"
         Index           =   11
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuAdminIn 
         Caption         =   "Edit Banlist..."
         Index           =   12
      End
      Begin VB.Menu mnuAdminIn 
         Caption         =   "Start Game Server"
         Index           =   13
      End
      Begin VB.Menu mnuAdminIn 
         Caption         =   "Stop Game Server"
         Index           =   14
      End
      Begin VB.Menu mnuAdminIn 
         Caption         =   "-"
         Index           =   20
      End
      Begin VB.Menu mnuAdminIn 
         Caption         =   "Hiding"
         Index           =   30
         Begin VB.Menu mnuAdminIn2 
            Caption         =   "Hide Me"
            Index           =   1
            Shortcut        =   ^H
         End
         Begin VB.Menu mnuAdminIn2 
            Caption         =   "Unhide Me"
            Index           =   2
            Shortcut        =   ^U
         End
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Se&ttings"
      Begin VB.Menu mnuSettingsIn 
         Caption         =   "&Change password..."
         Index           =   0
      End
      Begin VB.Menu mnuSettingsIn 
         Caption         =   "Display Names in Colour"
         Checked         =   -1  'True
         Index           =   1
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSettingsIn 
         Caption         =   "TimeStamp console messages"
         Index           =   2
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSettingsIn 
         Caption         =   "Show Messages in Text Window"
         Index           =   3
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSettingsIn 
         Caption         =   "Pop up Admin Chat on message"
         Checked         =   -1  'True
         Index           =   4
      End
      Begin VB.Menu mnuSettingsIn 
         Caption         =   "Reconnect on Disconnect"
         Checked         =   -1  'True
         Index           =   5
      End
      Begin VB.Menu mnuSettingsIn 
         Caption         =   "Snap Windows"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu mnuSettingsIn 
         Caption         =   "Enable Bing Sound"
         Checked         =   -1  'True
         Index           =   7
      End
      Begin VB.Menu mnuSettingsIn 
         Caption         =   "-"
         Index           =   20
      End
      Begin VB.Menu mnuSettingsIn 
         Caption         =   "Interface Colours..."
         Index           =   30
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuSettingsIn 
         Caption         =   "-"
         Index           =   40
      End
      Begin VB.Menu mnuSettingsIn 
         Caption         =   "Configure Half-Life..."
         Index           =   50
      End
      Begin VB.Menu mnuSettingsIn 
         Caption         =   "Customize Toolbar..."
         Index           =   51
      End
      Begin VB.Menu mnuSettingsIn 
         Caption         =   "-"
         Index           =   60
      End
      Begin VB.Menu mnuSettingsIn 
         Caption         =   "Set Away Mode..."
         Index           =   61
      End
      Begin VB.Menu mnuSettingsIn 
         Caption         =   "Cancel Away Mode"
         Index           =   62
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowsIn 
         Caption         =   "&Player List"
         Index           =   0
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuWindowsIn 
         Caption         =   "&Logged-In Users"
         Index           =   1
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuWindowsIn 
         Caption         =   "&Map"
         Index           =   2
      End
      Begin VB.Menu mnuWindowsIn 
         Caption         =   "&Admin Chat"
         Index           =   3
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuWindowsIn 
         Caption         =   "Map &Statistics..."
         Index           =   4
      End
      Begin VB.Menu mnuWindowsIn 
         Caption         =   "Whiteboard"
         Index           =   5
      End
   End
   Begin VB.Menu mnuMessages 
      Caption         =   "&Messages"
      Begin VB.Menu mnuMessagesIn 
         Caption         =   "&Open Mailbox"
         Index           =   0
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuMessagesIn 
         Caption         =   "Compose &Message..."
         Index           =   1
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu mnuLogs 
      Caption         =   "&Logs"
      Begin VB.Menu mnuLogsIn 
         Caption         =   "View current Server Log..."
         Index           =   0
      End
      Begin VB.Menu mnuLogsIn 
         Caption         =   "View current App Logs..."
         Index           =   1
      End
      Begin VB.Menu mnuLogsIn 
         Caption         =   "Log Search..."
         Index           =   2
      End
   End
   Begin VB.Menu mnuFunctions 
      Caption         =   "Functions"
      Visible         =   0   'False
      Begin VB.Menu mnuFunctionsIn 
         Caption         =   "Add RealPlayer using Current Name"
         Index           =   1
      End
      Begin VB.Menu mnuFunctionsIn 
         Caption         =   "More Player Options..."
         Index           =   2
         Begin VB.Menu mnuFunctionsMore 
            Caption         =   "Add RealPlayer..."
            Index           =   1
         End
         Begin VB.Menu mnuFunctionsMore 
            Caption         =   "Add Realplayer using Entry Name"
            Index           =   2
         End
         Begin VB.Menu mnuFunctionsMore 
            Caption         =   "Show in RealPlayers Window"
            Index           =   3
         End
         Begin VB.Menu mnuFunctionsMore 
            Caption         =   "Change Name to RealName"
            Index           =   4
         End
         Begin VB.Menu mnuFunctionsMore 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuFunctionsMore 
            Caption         =   "Add ClanPlayer"
            Index           =   6
         End
         Begin VB.Menu mnuFunctionsMore 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnuFunctionsMore 
            Caption         =   "Announce Player ID"
            Index           =   8
         End
      End
      Begin VB.Menu mnuFunctionsIn 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFunctionsIn 
         Caption         =   "Kill Player"
         Index           =   4
      End
      Begin VB.Menu mnuFunctionsIn 
         Caption         =   "Change Player Name..."
         Index           =   5
      End
      Begin VB.Menu mnuFunctionsIn 
         Caption         =   "Change Class to..."
         Index           =   6
         Begin VB.Menu mnuFunctionsClassIn 
            Caption         =   "Scout"
            Index           =   0
         End
         Begin VB.Menu mnuFunctionsClassIn 
            Caption         =   "Sniper"
            Index           =   1
         End
         Begin VB.Menu mnuFunctionsClassIn 
            Caption         =   "Soldier"
            Index           =   2
         End
         Begin VB.Menu mnuFunctionsClassIn 
            Caption         =   "Demoman"
            Index           =   3
         End
         Begin VB.Menu mnuFunctionsClassIn 
            Caption         =   "Medic"
            Index           =   4
         End
         Begin VB.Menu mnuFunctionsClassIn 
            Caption         =   "HWGuy"
            Index           =   5
         End
         Begin VB.Menu mnuFunctionsClassIn 
            Caption         =   "Pyro"
            Index           =   6
         End
         Begin VB.Menu mnuFunctionsClassIn 
            Caption         =   "Spy"
            Index           =   7
         End
         Begin VB.Menu mnuFunctionsClassIn 
            Caption         =   "Engineer"
            Index           =   8
         End
         Begin VB.Menu mnuFunctionsClassIn 
            Caption         =   "Civilian"
            Index           =   9
         End
      End
      Begin VB.Menu mnuFunctionsIn 
         Caption         =   "Voicing"
         Index           =   7
         Begin VB.Menu mnuFunctionsIn2 
            Caption         =   "De-voice"
            Index           =   0
         End
         Begin VB.Menu mnuFunctionsIn2 
            Caption         =   "Re-voice"
            Index           =   1
         End
      End
      Begin VB.Menu mnuFunctionsIn 
         Caption         =   "Points..."
         Index           =   8
         Begin VB.Menu mnuFunctionsPoints 
            Caption         =   "Add points..."
            Index           =   1
         End
         Begin VB.Menu mnuFunctionsPoints 
            Caption         =   "Announce Points"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFunctionsIn 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuFunctionsIn 
         Caption         =   "Ban for x Min..."
         Index           =   10
      End
      Begin VB.Menu mnuFunctionsIn 
         Caption         =   "Send Private Message..."
         Index           =   11
      End
      Begin VB.Menu mnuFunctionsIn 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuFunctionsIn 
         Caption         =   "Show in Map"
         Index           =   13
      End
      Begin VB.Menu mnuFuncScripts 
         Caption         =   "Scripts..."
         Visible         =   0   'False
         Begin VB.Menu mnuFuncScriptsIn 
            Caption         =   "Select a Script:"
            Enabled         =   0   'False
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuAdminEmail 
      Caption         =   "AdminEmail"
      Visible         =   0   'False
      Begin VB.Menu mnuAdminEmailIn 
         Caption         =   "Get All Users Messages"
         Index           =   0
      End
   End
   Begin VB.Menu mnuTeleporters 
      Caption         =   "TeleportMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuTeleportIn 
         Caption         =   "Delete"
         Index           =   0
      End
      Begin VB.Menu mnuTeleportIn 
         Caption         =   "Rename"
         Index           =   1
      End
   End
   Begin VB.Menu mnuSend 
      Caption         =   "ScriptSendMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuSendIn 
         Caption         =   "Send Changed Scripts"
         Index           =   0
      End
      Begin VB.Menu mnuSendIn 
         Caption         =   "Send All Scripts"
         Index           =   1
      End
      Begin VB.Menu mnuSendIn 
         Caption         =   "Send All && Close"
         Index           =   2
      End
   End
   Begin VB.Menu mnuGames 
      Caption         =   "GamesMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuGamesIn 
         Caption         =   "Read Away Message..."
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGamesIn 
         Caption         =   "Private Chat"
         Index           =   1
      End
      Begin VB.Menu mnuGamesIn 
         Caption         =   "Private Beep"
         Index           =   2
      End
      Begin VB.Menu mnuGamesIn 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuGamesIn 
         Caption         =   "Games:"
         Enabled         =   0   'False
         Index           =   10
      End
      Begin VB.Menu mnuGamesIn 
         Caption         =   "Tic Tac Toe"
         Index           =   11
      End
      Begin VB.Menu mnuGamesIn 
         Caption         =   "Hangman"
         Index           =   12
      End
      Begin VB.Menu mnuGamesIn 
         Caption         =   "Battleship"
         Index           =   13
      End
      Begin VB.Menu mnuGamesIn 
         Caption         =   "SillyZone Private Chat"
         Index           =   14
      End
      Begin VB.Menu mnuGamesIn 
         Caption         =   "Freaky's Car Race"
         Index           =   15
      End
   End
   Begin VB.Menu mnuBanTime 
      Caption         =   "Ban Time menu"
      Visible         =   0   'False
      Begin VB.Menu mnuBanTimeIn 
         Caption         =   "10 min"
         Index           =   0
      End
      Begin VB.Menu mnuBanTimeIn 
         Caption         =   "20 min"
         Index           =   1
      End
      Begin VB.Menu mnuBanTimeIn 
         Caption         =   "1 hr"
         Index           =   2
      End
      Begin VB.Menu mnuBanTimeIn 
         Caption         =   "2 hrs"
         Index           =   3
      End
      Begin VB.Menu mnuBanTimeIn 
         Caption         =   "4 hrs"
         Index           =   4
      End
      Begin VB.Menu mnuBanTimeIn 
         Caption         =   "Unlimited"
         Index           =   5
      End
   End
End
Attribute VB_Name = "MDIForm1"
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

Private Sub MDIForm_Load()

'restore window settings
'MDIForm1.Show

'MDIForm1.WindowState = 0
'If MDIForm1.WindowState = 0 Then

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
UnloadTime = True

End Sub

Private Sub MDIForm_Terminate()
End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)


'UnloadReady

ImGone = True
'save window settings

End Sub

Private Sub mnuAdminEmailIn_Click(Index As Integer)

frmMessageList.MnuClick Index



End Sub

Public Sub mnuAdminIn_Click(Index As Integer)

If Index = 0 Then SendExe
If Index = 1 Then SendPacket "EU", ""
If Index = 2 Then SendPacket "KB", ""
If Index = 3 Then SendPacket "ES", ""
If Index = 4 Then SendPacket "CL", ""
If Index = 5 Then SendPacket "SL", ""
If Index = 6 Then SendPacket "RP", ""
If Index = 7 Then SendPacket "WS", ""
If Index = 8 Then SendPacket "GW", ""
If Index = 9 Then SendPacket "LE", ""
If Index = 10 Then SendPacket "W1", ""
If Index = 11 Then SendPacket "F0", "": frmFileBrowser.Show
If Index = 12 Then SendPacket "GB", ""
If Index = 13 Then SendPacket "Z1", ""
If Index = 14 Then SendPacket "Z2", ""


End Sub

Private Sub mnuAdminIn2_Click(Index As Integer)

If Index = 1 Then SendPacket "H1", "": HiddenNow = True
If Index = 2 Then SendPacket "H2", "": HiddenNow = False

If HiddenNow = True Then
'    MDIForm1.Toolbar1.Buttons(17).Value = tbrPressed
    ChangeButton 17, 1, tbrPressed

Else
 '   MDIForm1.Toolbar1.Buttons(17).Value = tbrUnpressed
     ChangeButton 17, 1, tbrUnpressed

End If

End Sub

Private Sub mnuBanTimeIn_Click(Index As Integer)


On Error Resume Next
a$ = Form6.ListView1.SelectedItem

If a$ = "" Then Exit Sub

For i = 1 To Form6.ListView1.ListItems.Count
    If Form6.ListView1.ListItems.Item(i).Text = a$ Then j = i: Exit For
Next i

b$ = Form6.ListView1.ListItems.Item(j).SubItems(2) 'userid
'send ban command

If Index = 0 Then tme$ = "10"
If Index = 1 Then tme$ = "20"
If Index = 2 Then tme$ = "60"
If Index = 3 Then tme$ = "120"
If Index = 4 Then tme$ = "240"
If Index = 5 Then tme$ = "0"

If Index = 5 Then
    RS$ = InBox("Why are you perm-banning this player?", "Enter Ban Reason", "")
    If RS$ = "" Then MessBox ("You must enter a reason!"): Exit Sub
End If

a$ = Chr(251)
a$ = a$ + b$ + Chr(250)
a$ = a$ + tme$ + Chr(250)
a$ = a$ + RS$ + Chr(250)
a$ = a$ + Chr(251)

SendPacket "SB", b$

End Sub

Private Sub mnuFileIn_Click(Index As Integer)
If Index = 0 Then End

End Sub

Private Sub mnuSetupIn_Click(Index As Integer)

If Index = 1 Then Form3.Show

End Sub

Private Sub mnuViewIn_Click(Index As Integer)
If Index = 0 Then
    Form1.WindowState = 1
    Form6.Hide
End If
If Index = 1 Then
    Form6.Show
End If

End Sub



Private Sub mnuFuncScriptsIn_Click(Index As Integer)

StartFuncScript Index


End Sub

Private Sub mnuFunctionsClassIn_Click(Index As Integer)


Form6.FunctionsClass Index

End Sub

Private Sub mnuFunctionsIn_Click(Index As Integer)

Form6.Functions Index



End Sub

Private Sub mnuFunctionsIn2_Click(Index As Integer)

Form6.Functions Index + 60

End Sub

Private Sub mnuFunctionsMore_Click(Index As Integer)

Form6.Functions Index + 30

End Sub

Private Sub mnuFunctionsPoints_Click(Index As Integer)

Form6.Functions Index + 50

End Sub

Private Sub mnuGamesIn_Click(Index As Integer)

If Index = 0 Then

    ReadAwayMenu

Else

    GamesMenu Index

End If

End Sub

Private Sub mnuLogsIn_Click(Index As Integer)

If Index = 0 Then SendPacket "VL", "server"
If Index = 1 Then SendPacket "VL", "local"
If Index = 2 Then frmStartLogSearch.Show



End Sub

Private Sub mnuMessagesIn_Click(Index As Integer)
If Index = 0 Then SendPacket "M6", ""
If Index = 1 Then SendPacket "M5", ""

End Sub

Private Sub mnuScriptsIn_Click(Index As Integer)
ButtonShowMode = 0
If Index = 0 Then SendPacket "EC", ""
If Index = 1 Then SendPacket "BS", ""


End Sub

Private Sub mnuSendIn_Click(Index As Integer)

Form3.SendMenu Index


End Sub

Public Sub mnuSettingsIn_Click(Index As Integer)

If Index = 0 Then

    a$ = InBox("Please enter your old password:", "Change Password")
    b$ = InBox("Please enter your new password:", "Change Password")
       
    SendPacket "CP", a$ + Chr(250) + b$

End If

If Index = 1 Then

    If mnuSettingsIn(Index).Checked = True Then
        mnuSettingsIn(Index).Checked = False
        'Toolbar1.Buttons(19).Value = tbrUnpressed
        ChangeButton 19, 1, tbrUnpressed

    Else
        mnuSettingsIn(Index).Checked = True
        'Toolbar1.Buttons(19).Value = tbrPressed
        ChangeButton 19, 1, tbrPressed
    End If
End If

If Index = 2 Then

    If mnuSettingsIn(Index).Checked = True Then
        mnuSettingsIn(Index).Checked = False
        'Toolbar1.Buttons(20).Value = tbrUnpressed
        ChangeButton 20, 1, tbrUnpressed
    Else
        mnuSettingsIn(Index).Checked = True
        'Toolbar1.Buttons(20).Value = tbrPressed
        ChangeButton 20, 1, tbrPressed
    End If
End If

If Index = 3 Then
    If mnuSettingsIn(3).Checked = True Then
        mnuSettingsIn(3).Checked = False
        'Toolbar1.Buttons(21).Value = tbrUnpressed
        ChangeButton 21, 1, tbrUnpressed
    Else
        mnuSettingsIn(3).Checked = True
        'Toolbar1.Buttons(21).Value = tbrPressed
        ChangeButton 21, 1, tbrPressed
    End If
End If

If Index = 4 Then
    If mnuSettingsIn(4).Checked = True Then
        mnuSettingsIn(4).Checked = False
        'Toolbar1.Buttons(21).Value = tbrUnpressed
        
    Else
        mnuSettingsIn(4).Checked = True
        'Toolbar1.Buttons(21).Value = tbrPressed
        
    End If
End If

If Index = 5 Then
    If mnuSettingsIn(5).Checked = True Then
        mnuSettingsIn(5).Checked = False
        'Toolbar1.Buttons(21).Value = tbrUnpressed
        
    Else
        mnuSettingsIn(5).Checked = True
        'Toolbar1.Buttons(21).Value = tbrPressed
        
    End If
End If

If Index = 6 Then
    If mnuSettingsIn(6).Checked = True Then
        mnuSettingsIn(6).Checked = False
        'Toolbar1.Buttons(21).Value = tbrUnpressed
        
    Else
        mnuSettingsIn(6).Checked = True
        'Toolbar1.Buttons(21).Value = tbrPressed
        
    End If
End If

If Index = 7 Then
    If mnuSettingsIn(7).Checked = True Then
        mnuSettingsIn(7).Checked = False
        'Toolbar1.Buttons(21).Value = tbrUnpressed
        
    Else
        mnuSettingsIn(7).Checked = True
        'Toolbar1.Buttons(21).Value = tbrPressed
        
    End If
End If


If Index = 30 Then frmColors.Show
If Index = 50 Then frmHalfConfig.Show
If Index = 51 Then
    
    frmCustomize.Show
    
End If
SaveSetting "Server Assistant", "Settings", "NamesInColor", Ts(CLng(mnuSettingsIn(1).Checked))
SaveSetting "Server Assistant", "Settings", "TimeStamp", Ts(CLng(mnuSettingsIn(2).Checked))
SaveSetting "Server Assistant", "Settings", "ShowMessages", Ts(CLng(mnuSettingsIn(3).Checked))
SaveSetting "Server Assistant", "Settings", "PopUpAdmin", Ts(CLng(mnuSettingsIn(4).Checked))
SaveSetting "Server Assistant", "Settings", "AutoReconnect", Ts(CLng(mnuSettingsIn(5).Checked))
SaveSetting "Server Assistant", "Settings", "SnapWindows", Ts(CLng(mnuSettingsIn(6).Checked))
SaveSetting "Server Assistant", "Settings", "EnableBing", Ts(CLng(mnuSettingsIn(7).Checked))


If Index = 61 Then

    frmSetAway.Show
    
End If
If Index = 62 Then

    MyAwayMode = 0
    MyAwayMsg = 0
    UpdateAwayMode
    
End If

End Sub

Private Sub mnuTeleportIn_Click(Index As Integer)

frmMap.Functions Index



End Sub

Private Sub mnuWindowsIn_Click(Index As Integer)

If Index = 0 Then SendPacket "SU", "": ShowPlayers = True: Form6.Show
If Index = 1 Then SendPacket "CU", "": ShowUsers = True: frmConnectUsers.Show
If Index = 2 Then

    frmMap.Show


End If

If Index = 3 Then

    SendPacket "SC", ""
    frmAdminChat.Show


End If
If Index = 4 Then SendPacket "MP", ""
If Index = 5 Then frmWhiteBoard.Show


End Sub

Private Sub Timer1_Timer()
'If ImGone Then End

If UnloadTime Then
    UnloadReady
    Unload Me
End If
End Sub

Private Sub Timer2_Timer()

'Refresh MDIForm1



    MDIForm1.Height = GetSetting("Server Assistant Client", "Window", "winh", 800 * Screen.TwipsPerPixelX)
    MDIForm1.Top = GetSetting("Server Assistant Client", "Window", "wint", 50 * Screen.TwipsPerPixelX)
    MDIForm1.Left = GetSetting("Server Assistant Client", "Window", "winl", 50 * Screen.TwipsPerPixelX)
    MDIForm1.Width = GetSetting("Server Assistant Client", "Window", "winw", 800 * Screen.TwipsPerPixelX)
'End If

wwdd = Val(GetSetting("Server Assistant Client", "Window", "winmd", 2))

If wwdd <> 1 Then MDIForm1.WindowState = wwdd


Timer2.Enabled = False

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

nb = Val(Button.Tag)

Select Case nb
    Case 1
        SendPacket "BS", ""
    Case 2
        frmMap.Show
    Case 3
        SendPacket "F0", "": frmFileBrowser.Show
    Case 4
        If HLEXEPath = "" Then
            MsgBox "You must configure this feature first."
            frmHalfConfig.Show
            Exit Sub
        End If
       
       
        a$ = HLEXEPath + " -console " + HLExtraArgs + " -game " + HLGame + " +connect " + HLIP + ":" + HLPort
        'MsgBox a$
    
        If Dir(HLEXEPath) = "" Then
            MsgBox "Half-Life EXE Not Found!"
            Exit Sub
        End If
        Shell a$, vbNormalFocus
        
        If HLQuitSA = 1 Then UnloadTime = True
        
        If HLSetAway Then
            MyAwayMode = 4
            MyAwayMsg = "I just joined the game at " + HLIP + ":" + HLPort
            AutoSet = False
            UpdateAwayMode
            
        End If
    
    Case 5
        mnuFileIn_Click 0
    Case 6
        SendPacket "EC", ""
    Case 7
        mnuAdminIn_Click 1
    Case 8
        mnuAdminIn_Click 2
    Case 9
        mnuAdminIn_Click 3
    Case 10
        mnuAdminIn_Click 4
    Case 11
        mnuAdminIn_Click 5
    Case 12
        mnuAdminIn_Click 6
    Case 13
        mnuAdminIn_Click 7
    Case 14
        mnuAdminIn_Click 8
    Case 15
        mnuAdminIn_Click 9
    Case 16
        mnuAdminIn_Click 10
    Case 17
        'hidden button - special case
        
        If HiddenNow = True Then
            mnuAdminIn2_Click 2
        Else
            mnuAdminIn2_Click 1
        End If
        
    Case 18
        mnuSettingsIn_Click 0
    Case 19
        mnuSettingsIn_Click 1
    Case 20
        mnuSettingsIn_Click 2
    Case 21
        mnuSettingsIn_Click 3
    Case 22
        mnuSettingsIn_Click 30
    Case 23
        mnuSettingsIn_Click 7
    Case 24
        mnuWindowsIn_Click 0
    Case 25
        mnuWindowsIn_Click 1
    Case 26
        mnuMessagesIn_Click 0
    Case 27
        mnuMessagesIn_Click 1
    Case 28
        mnuLogsIn_Click 0
    Case 29
        mnuLogsIn_Click 1
    Case 30
        mnuLogsIn_Click 2
        
    Case 31
        mnuWindowsIn_Click 3
    Case 32
         mnuWindowsIn_Click 5
    Case 33
        mnuSettingsIn_Click 61
    Case 34
        mnuSettingsIn_Click 62
    Case 35
        mnuAdminIn_Click 13
    Case 36
        mnuAdminIn_Click 14
        
        
    End Select
    
    If nb > 50 Then
    
        ' figure out what the script is called
        
        a$ = Button.ToolTipText
        a$ = ReplaceString(a$, "Script - ", "")
        
        ScriptButtonName = a$
                
        ButtonShowMode = 2
        SendPacket "BS", ""
    
    End If
End Sub

