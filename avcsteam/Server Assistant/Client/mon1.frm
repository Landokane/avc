VERSION 5.00
Object = "{B7FC3580-8CE7-11CF-9754-00AA00C00908}#1.0#0"; "WINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Server Monitor"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8445
   Icon            =   "mon1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   6840
      Top             =   2520
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save Now"
      Height          =   435
      Left            =   7560
      TabIndex        =   29
      Top             =   1980
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   7980
      TabIndex        =   26
      Text            =   "7"
      Top             =   1620
      Width           =   435
   End
   Begin VB.TextBox Text5 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   5220
      TabIndex        =   24
      Text            =   "27015"
      Top             =   720
      Width           =   3195
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Add"
      Height          =   255
      Left            =   3960
      TabIndex        =   23
      Top             =   2940
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Del"
      Height          =   255
      Left            =   3960
      TabIndex        =   22
      Top             =   2700
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Disable"
      Height          =   435
      Left            =   6360
      TabIndex        =   20
      Top             =   1140
      Width           =   915
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add Selected"
      Height          =   315
      Left            =   0
      TabIndex        =   19
      Top             =   2100
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   2460
      Width           =   615
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2715
      Left            =   0
      TabIndex        =   17
      Top             =   3240
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4789
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Name"
         Text            =   "Name"
         Object.Width           =   4419
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Unque"
         Text            =   "Unique ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Last Alias"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Last Time Seen"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   3735
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7800
      Top             =   2520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check Now!"
      Height          =   435
      Left            =   7320
      TabIndex        =   12
      Top             =   1140
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   5220
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   360
      Width           =   3195
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   5220
      TabIndex        =   9
      Text            =   "127.0.0.1"
      Top             =   0
      Width           =   3195
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      Top             =   2820
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   2460
      Width           =   2415
   End
   Begin WINSOCKLib.UDP UDP1 
      Left            =   7320
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      Blocking        =   0   'False
      SleepTime       =   10
      RemoteHost      =   ""
      RemotePort      =   0
      LocalPort       =   26000
   End
   Begin VB.Label ServStatus 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3840
      TabIndex        =   28
      Top             =   1980
      Width           =   45
   End
   Begin VB.Line Line2 
      X1              =   60
      X2              =   8460
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Delay"
      Height          =   195
      Left            =   7560
      TabIndex        =   27
      Top             =   1680
      Width           =   405
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Port"
      Height          =   195
      Left            =   3780
      TabIndex        =   25
      Top             =   780
      Width           =   285
   End
   Begin VB.Label LastCheck 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3840
      TabIndex        =   21
      Top             =   1620
      Width           =   45
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Left            =   4620
      TabIndex        =   15
      Top             =   2940
      Width           =   450
   End
   Begin VB.Label Status 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6060
      TabIndex        =   14
      Top             =   2940
      Width           =   45
   End
   Begin VB.Label NextUpd 
      AutoSize        =   -1  'True
      Caption         =   "Time Till Next Update:"
      Height          =   195
      Left            =   3840
      TabIndex        =   13
      Top             =   1260
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   3780
      X2              =   8400
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Server Password"
      Height          =   195
      Left            =   3780
      TabIndex        =   10
      Top             =   420
      Width           =   1200
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Server IP"
      Height          =   195
      Left            =   3780
      TabIndex        =   8
      Top             =   60
      Width           =   660
   End
   Begin VB.Label LastTime 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6060
      TabIndex        =   7
      Top             =   2700
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Last Time Seen:"
      Height          =   195
      Left            =   4620
      TabIndex        =   6
      Top             =   2700
      Width           =   1155
   End
   Begin VB.Label Current 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6060
      TabIndex        =   5
      Top             =   2460
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Current Alias:"
      Height          =   195
      Left            =   4620
      TabIndex        =   4
      Top             =   2460
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Target Name:"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Target Unique ID's"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   2520
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Monitorss
    Name As String
    Numbers As String
    LastFound As String
    Status As String
    Current As String
    Checked As Boolean
End Type


Dim LastRCON As String
Dim UpdateTime As Integer
Dim Names(1 To 32) As String
Dim Numbers(1 To 32) As String
Dim Monitor(1 To 200) As Monitorss
Dim NumMonitors As Integer
Dim CurrentIndex As Integer
Dim FormName As String


Private Sub AddMonitors()

ListView1.ListItems.Clear
ListView1.Sorted = False
For i = 1 To NumMonitors
    If Monitor(i).Status = "" Then Monitor(i).Status = "Un-Updated"
    ListView1.ListItems.Add 1, , Monitor(i).Name
    ListView1.ListItems.Item(1).SubItems(1) = Monitor(i).Numbers
    ListView1.ListItems.Item(1).SubItems(2) = Monitor(i).Current
    ListView1.ListItems.Item(1).SubItems(3) = Monitor(i).Status
    ListView1.ListItems.Item(1).SubItems(4) = Monitor(i).LastFound
    ListView1.ListItems.Item(1).Tag = Str(i)
    ListView1.ListItems.Item(1).Checked = Monitor(i).Checked
Next i
ListView1.Sorted = True

End Sub

Private Function FindMonitor(Unique As String) As Integer

'finds a person by their unique id
For i = 1 To NumMonitors
    If InStr(1, Monitor(i).Numbers, Unique) Then j = i
Next i

FindMonitor = j

End Function


Private Sub Check()

LastRCON = ""
UDP1.LocalPort = 26000
UDP1.RemotePort = Text5
UDP1.RemoteHost = Text3

hed$ = Chr(255) + Chr(255) + Chr(255) + Chr(255) + "rcon" + Chr(32)
UDP1.SendData hed$ + " " + Text4 + " users"
'UDP1.SendData hed$ + " " + Text4 + " logaddress 24.64.165.49 26000"
Timer2.Enabled = True

End Sub

Private Sub Check2()


If Left(LastRCON, 3) = "Bad" Then
    ServStatus = "Bad RCON Password!"
    Timer1.Enabled = False
    Command5.Caption = "Enable"
    Exit Sub
End If
'now sort out the uniquenumber stuff

t$ = Right(LastRCON, 8)

For i = 1 To 8
    Debug.Print Asc(Mid(t$, i, 1))
Next i

For i = 1 To NumMonitors
    If Monitor(i).Checked = True Then Monitor(i).Status = "Not Found!"
Next i

f = 1

List1.Clear
LastRCON = " " + LastRCON
Do
    e = f
    f = InStr(e + 1, LastRCON, Chr(10))
    
    If f <> 0 Then
        d$ = Mid(LastRCON, e + 1, f - e - 1)
        d$ = Trim(d$)
        If Left(d$, 6) = "userid" Then GoTo nxt
        If Left(d$, 6) = "------" Then GoTo nxt
        'sample: 89 : 10702879 : Dav
        
        h = InStr(1, d$, ":")
        j = InStr(h + 1, d$, ":")
        
        If h > 0 And j > 0 Then
            
            n$ = Mid(d$, h + 1, j - h - 1)
            n$ = Trim(n$)
            
            p$ = ""
            If Len(d$) > j Then p$ = Right(d$, Len(d$) - j)
            p$ = Trim(p$)
            
            nn$ = n$ + Space(10 - Len(n$))
                        
            List1.AddItem nn$ + " : " + p$
                        
            r = List1.NewIndex
            Names(r + 1) = p$
            Numbers(r + 1) = n$
                       
            a = FindMonitor(n$)
            
            If a > 0 Then
                FoundUser a, p$
            End If

        End If
    End If
nxt:
Loop Until f = 0
LastCheck = "Last Succesful Check: " + Time$ + " " + Date$
ServStatus = "Update Successful!"
AddMonitors
Form1.Caption = FormName + " - " + Trim(Str(List1.ListCount)) + " users"
End Sub

Private Sub FoundUser(a, p$)

If Monitor(a).Checked = False Then Exit Sub

Monitor(a).LastFound = Time$ + " " + Date$

notnew = 0

If Monitor(a).Current = p$ Then notnew = 1
Monitor(a).Current = p$


Monitor(a).Status = "Found, New Name!"

If p$ = Monitor(a).Name Then
    'user still has original name
    Monitor(a).Status = "Found, Original Name"
    orig = 1
End If

If notnew = 0 And orig = 0 Then
    hed$ = Chr(255) + Chr(255) + Chr(255) + Chr(255) + "rcon" + Chr(32)
    UDP1.SendData hed$ + " " + Text4 + " say Alert! " + Monitor(a).Name + " is playing as " + p$
End If

End Sub

Private Sub Command1_Click()
Check
UpdateTime = Val(Text6)

End Sub

Private Sub Command2_Click()
If CurrentIndex = 0 Then Exit Sub
a = CurrentIndex

Monitor(a).Numbers = Text1
Monitor(a).Name = Text2
AddMonitors
CurrentIndex = 0
Text1 = ""
Text2 = ""
Current = ""
Status = ""
LastTime = ""
End Sub

Private Sub Command3_Click()
SaveSettings
End Sub

Private Sub Command4_Click()
a = List1.ListIndex
If a = -1 Then Exit Sub
a = a + 1
NumMonitors = NumMonitors + 1
i = NumMonitors
Monitor(i).Name = Names(a)
Monitor(i).Numbers = Numbers(a)
Monitor(i).Status = "Un-Updated"
Monitor(i).Checked = False
Monitor(i).LastFound = ""
Monitor(i).Current = ""
AddMonitors
CurrentIndex = 0
Text1 = ""
Text2 = ""
Current = ""
Status = ""
LastTime = ""

End Sub


Private Sub Command5_Click()

If Timer1.Enabled = True Then
    Command5.Caption = "Enable"
    Timer1.Enabled = False
    Exit Sub
End If
If Timer1.Enabled = False Then
    Command5.Caption = "Disable"
    Timer1.Enabled = True
    Exit Sub
End If

End Sub

Private Sub Command6_Click()

If CurrentIndex > 0 Then
    a = CurrentIndex
    If ListView1.ListItems.Count = 0 Then Exit Sub
    If a = -1 Then Exit Sub
    
    b$ = Monitor(a).Name
    c = MsgBox("Are you sure you want to delete " + Chr(34) + b$ + Chr(34) + "?", vbYesNo, "Delete Monitored User")
    
    If c = vbYes Then
        If a < NumMonitors Then
            For i = a To NumMonitors - 1
                Monitor(i).Checked = Monitor(i + 1).Checked
                Monitor(i).Name = Monitor(i + 1).Name
                Monitor(i).Numbers = Monitor(i + 1).Numbers
                Monitor(i).Status = Monitor(i + 1).Status
                Monitor(i).LastFound = Monitor(i + 1).LastFound
                Monitor(i).Current = Monitor(i + 1).Current
            Next i
        End If
        
        NumMonitors = NumMonitors - 1
        AddMonitors
        CurrentIndex = 0
        Text1 = ""
        Text2 = ""
        Current = ""
        Status = ""
        LastTime = ""
    End If
End If
End Sub

Private Sub Command7_Click()
NumMonitors = NumMonitors + 1
i = NumMonitors
Monitor(i).Name = "Unnamed"
Monitor(i).Numbers = "00000000"
Monitor(i).Status = "Un-Updated"
Monitor(i).Checked = False
Monitor(i).LastFound = ""
Monitor(i).Current = ""
AddMonitors
CurrentIndex = 0
Text1 = ""
Text2 = ""
Current = ""
Status = ""
LastTime = ""
End Sub

Private Sub Form_Load()
FormName = Me.Caption

LoadSettings

UpdateTime = Val(Text6)

AddMonitors
End Sub

Private Sub Form_Resize()
ListView1.Width = Form1.Width - 120
ListView1.Height = Form1.Height - ListView1.Top - 520

End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSettings
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
a = Val(Item.Tag)
Monitor(a).Checked = Item.Checked
If a = CurrentIndex Then ListView1_ItemClick Item
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
a = Val(Item.Tag)
If Item.Checked = True Then c$ = "Active"
If Item.Checked = False Then c$ = "Inactive"

Text1 = Monitor(a).Numbers
Text2 = Monitor(a).Name
Current = Monitor(a).Current
LastTime = Monitor(a).LastFound
Status = Monitor(a).Status + ", " + c$
CurrentIndex = a

End Sub

Private Sub Text6_Change()
If Val(Text6) < 5 And Val(Text6) <> 0 Then Text6 = "5"
End Sub

Private Sub Timer1_Timer()
UpdateTime = UpdateTime - 1
If UpdateTime <= 0 Then UpdateTime = Val(Text6): Check

NextUpd = "Time Till Next Update: " + Trim(Str(UpdateTime))
End Sub


Private Sub Timer2_Timer()
If LastRCON = "" Then
    ServStatus = "Coundn't connect to server!"
    Timer2.Enabled = False
End If
End Sub

Private Sub UDP1_DataArrival(ByVal bytesTotal As Long)
    
    UDP1.GetData a$
    a$ = Right(a$, Len(a$) - 5)
    'a$ = ReplaceString(a$, Chr(10), vbCrLf)
    LastRCON = LastRCON + a$
    r$ = Right(LastRCON, 8)
    r$ = Left(r$, 6)
    If r$ = Chr(117) + Chr(115) + Chr(101) + Chr(114) + Chr(115) + Chr(10) Then
        Timer2.Enabled = False
        Check2
    End If
    If Left(LastRCON, 3) = "Bad" Then
        ServStatus = "Bad RCON Password!"
        Timer1.Enabled = False
        Command5.Caption = "Enable"
    End If
End Sub

Function ReplaceString(ByVal Txt As String, ByVal from_str As String, ByVal to_str As String)
Dim new_txt As String
Dim pos As Integer

    Do While Len(Txt) > 0
        pos = InStr(Txt, from_str)
        If pos = 0 Then
            ' No more occurrences.
            new_txt = new_txt & Txt
            Txt = ""
        Else
            ' Found it.
            new_txt = new_txt & Left$(Txt, pos - 1) & to_str
            Txt = Mid$(Txt, pos + Len(from_str))
        End If
    Loop

    ReplaceString = new_txt
End Function
Function CheckForFile(a$) As Boolean
    b$ = Dir(a$)
    If b$ = "" Then CheckForFile = False
    If b$ <> "" Then CheckForFile = True
    
End Function
Private Sub LoadSettings()

Dim Strings(1 To 4) As String


a$ = App.Path + "\servdata.dat"

If CheckForFile(a$) Then

    Open a$ For Binary As #1
    
        Get #1, , NumMonitors
        Get #1, , Monitor
        Get #1, , Strings
    Close #1
    Text3 = Strings(1)
    Text4 = Strings(2)
    Text5 = Strings(3)
    Text6 = Strings(4)
    

End If


End Sub

Private Sub SaveSettings()
Dim Strings(1 To 4) As String
a$ = App.Path + "\servdata.dat"

If CheckForFile(a$) Then Kill a$

Open a$ For Binary As #1

    Strings(1) = Text3
    Strings(2) = Text4
    Strings(3) = Text5
    Strings(4) = Text6
    
    Put #1, , NumMonitors
    Put #1, , Monitor
    Put #1, , Strings
Close #1




End Sub
