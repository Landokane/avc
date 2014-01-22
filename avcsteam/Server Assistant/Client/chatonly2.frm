VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form6 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Server Status"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   8550
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1260
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "chatonly2.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "chatonly2.frx":059C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "chatonly2.frx":0B38
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "chatonly2.frx":10D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "chatonly2.frx":1670
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "chatonly2.frx":1C0C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3480
      Width           =   1155
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   6059
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3200
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Team"
         Object.Width           =   1401
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Class"
         Object.Width           =   2379
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Time"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ListPerc As Double
Dim Text1Perc As Double
Dim Text2Perc As Double

Dim Dragger As Integer

Public Sub Functions(Index As Integer)

On Error Resume Next

If Index = 0 Then
    a$ = ListView1.SelectedItem
    If a$ = "" Then Exit Sub
    
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(I).Text = a$ Then j = I: Exit For
    Next I
    
    B$ = ListView1.ListItems.Item(j).SubItems(3) 'uniqueid
    
    'get real name
    nn$ = InputBox("Enter the REAL name for player " + a$, "Add RealPlayer", a$)
    If nn$ = "" Then Exit Sub
    
    AddRealPlayer nn$, B$

ElseIf Index = 1 Then
    a$ = ListView1.SelectedItem
    If a$ = "" Then Exit Sub
    
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(I).Text = a$ Then j = I: Exit For
    Next I
    
    B$ = ListView1.ListItems.Item(j).SubItems(3) 'uniqueid
    
    'get real name
    AddRealPlayer a$, B$
    
ElseIf Index = 5 Then
    'kick
    
    a$ = ListView1.SelectedItem
    
    If a$ = "" Then Exit Sub
    
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(I).Text = a$ Then j = I: Exit For
    Next I
    
    B$ = Trim(ListView1.ListItems.Item(j).SubItems(2)) 'userid
    'send kick command
    SendPacket "CA", B$

ElseIf Index = 7 Then
    'kill
    
    a$ = ListView1.SelectedItem
    
    If a$ = "" Then Exit Sub
    
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(I).Text = a$ Then j = I: Exit For
    Next I
    B$ = Trim(ListView1.ListItems.Item(j).SubItems(2)) 'userid
    
    SendPacket "RC", "kill " + B$

ElseIf Index = 9 Then
    'kill
    
    a$ = ListView1.SelectedItem
    
    If a$ = "" Then Exit Sub
    
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(I).Text = a$ Then j = I: Exit For
    Next I
    B$ = Trim(ListView1.ListItems.Item(j).SubItems(2)) 'userid
    
    nn1$ = InputBox("Enter new name:", "Change Player Name", a$)
    If nn1$ = "" Then Exit Sub
    
    SendPacket "RC", "changename " + B$ + " " + nn1$

ElseIf Index = 10 Then
    'kill
    
    a$ = ListView1.SelectedItem
    
    If a$ = "" Then Exit Sub
    
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(I).Text = a$ Then j = I: Exit For
    Next I
    B$ = Trim(ListView1.ListItems.Item(j).SubItems(2)) 'userid
    
    SendPacket "RC", "setreal " + a$

ElseIf Index = 2 Then
    'kill
    
    a$ = ListView1.SelectedItem
    
    If a$ = "" Then Exit Sub
    
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(I).Text = a$ Then j = I: Exit For
    Next I
    B$ = Trim(ListView1.ListItems.Item(j).SubItems(2)) 'userid
    
    SendPacket "RC", "annid " + a$

ElseIf Index = 12 Then
    'kill
    
    a$ = ListView1.SelectedItem
    
    If a$ = "" Then Exit Sub
    
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(I).Text = a$ Then j = I: Exit For
    Next I
    B$ = ListView1.ListItems.Item(j).SubItems(3) 'unique
    
    nn1$ = InputBox("Enter minutes to ban for:", "Ban Player " + a$ + " for X Minutes", "30")
    If nn1$ = "" Then Exit Sub
    
    SendPacket "RC", "banid " + nn1$ + " " + B$ + " kick"

ElseIf Index = 3 Then
    'kill
    
    a$ = ListView1.SelectedItem
    
    If a$ = "" Then Exit Sub
    
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(I).Text = a$ Then j = I: Exit For
    Next I
    B$ = ListView1.ListItems.Item(j).SubItems(3) 'unique
    
    FindReal = B$
    SendPacket "RP", ""

ElseIf Index = 14 Then
    'kill
    
    a$ = ListView1.SelectedItem
    
    If a$ = "" Then Exit Sub
    
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(I).Text = a$ Then j = I: Exit For
    Next I
    B$ = Trim(ListView1.ListItems.Item(j).SubItems(2)) 'userid
    
    nn1$ = InputBox("Enter Private Message:", "Send Private Message", "")
    If nn1$ = "" Then Exit Sub
    
    SendPacket "RC", "talkto " + B$ + " " + nn1$

ElseIf Index = 60 Then
    'kill
    
    a$ = ListView1.SelectedItem
    
    SendPacket "RC", "devoice " + a$
    
ElseIf Index = 61 Then
    'kill
    
    a$ = ListView1.SelectedItem
    
   
    SendPacket "RC", "revoice " + a$
    
ElseIf Index = 19 Then
    'kill
    
    a$ = ListView1.SelectedItem
   
    If a$ = "" Then Exit Sub
    
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(I).Text = a$ Then j = I: Exit For
    Next I
    B$ = Trim(ListView1.ListItems.Item(j).SubItems(2)) 'userid
    
    frmMap.RotPlayerNum = Val(B$)
    frmMap.Update2
    
End If





End Sub


Public Sub FunctionsClass(Index As Integer)

On Error Resume Next

'change class

a$ = ListView1.SelectedItem
If a$ = "" Then Exit Sub

For I = 1 To ListView1.ListItems.Count
    If ListView1.ListItems.Item(I).Text = a$ Then j = I: Exit For
Next I
B$ = Trim(ListView1.ListItems.Item(j).SubItems(2)) 'userid


cl$ = Ts(Index + 1)
If Index = 9 Then cl$ = "11"

SendPacket "RC", "changeclass " + B$ + " " + cl$

End Sub

Private Sub Command1_Click()

'kick
On Error Resume Next
a$ = ListView1.SelectedItem

If a$ = "" Then Exit Sub

For I = 1 To ListView1.ListItems.Count
    If ListView1.ListItems.Item(I).Text = a$ Then j = I: Exit For
Next I

B$ = ListView1.ListItems.Item(j).SubItems(2) 'userid

'send kick command
SendPacket "SK", B$

End Sub

Private Sub Command2_Click()
On Error Resume Next
a$ = ListView1.SelectedItem

If a$ = "" Then Exit Sub

For I = 1 To ListView1.ListItems.Count
    If ListView1.ListItems.Item(I).Text = a$ Then j = I: Exit For
Next I

B$ = ListView1.ListItems.Item(j).SubItems(2) 'userid
'send ban command
SendPacket "SB", B$

End Sub

Private Sub Command3_Click()

SendPacket "SU", ""

End Sub

Private Sub Command4_Click()

MDIForm1.PopupMenu MDIForm1.mnuFunctions

End Sub

Private Sub Command5_Click()

frmMap.Show

End Sub

Private Sub Command6_Click()
On Error Resume Next
'kill

a$ = ListView1.SelectedItem

If a$ = "" Then Exit Sub

For I = 1 To ListView1.ListItems.Count
    If ListView1.ListItems.Item(I).Text = a$ Then j = I: Exit For
Next I
B$ = ListView1.ListItems.Item(j).SubItems(2) 'userid

SendPacket "RC", "kill " + B$

End Sub

Private Sub Form_Load()
ListPerc = 0.3
Text1Perc = 0.5
Text2Perc = 0.2
Me.Width = 10500
ShowPlayers = True
UpdatePlayerList

Command5.Enabled = DllEnabled

End Sub


Sub Update()

If Me.WindowState = 1 Then Exit Sub

h = Me.Height
h = h - Command1.Height - ListView1.Top - 420

If Me.Width < 3705 Then Me.Width = 3705

ListView1.Height = h '* ListPerc
'Text1.Top = ListView1.Height + 45

'Text1.Height = h * Text1Perc
'Text2.Top = Text1.Height + 45 + Text1.Top
'Text2.Height = h * Text2Perc
Command1.Top = h + 60
Command2.Top = h + 60
Command3.Top = h + 60
Command4.Top = h + 60
Command5.Top = h + 60


Command5.Left = Me.Width - Command1.Width - Command2.Width - Command4.Width - Command5.Width - 300
Command4.Left = Me.Width - Command1.Width - Command2.Width - Command4.Width - 240
Command1.Left = Me.Width - Command1.Width - Command2.Width - 180
Command2.Left = Me.Width - Command2.Width - 120
Command3.Left = 60


End Sub

Private Sub Form_Resize()

If Me.Width < 2000 Then Me.Width = 2000
If Me.Height < 2500 Then Me.Height = 2500

Update

ListView1.Width = Me.Width - 120
'Text1.Width = Me.Width - 120
'Text2.Width = Me.Width - 120

'Label1.Top = ListView1.Height - (Label1.Height / 2)
'Label2.Top = ListView1.Height - (Label1.Height / 2)





End Sub

Private Sub Form_Unload(Cancel As Integer)
ShowPlayers = False

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

'istview1.ColumnHeaders(ListView1.SortKey + 1).for

ListView1.Sorted = True
k = ListView1.SortKey

If k = (ColumnHeader.Index - 1) Then
    If ListView1.SortOrder = lvwDescending Then
        ListView1.SortOrder = lvwAscending
    Else
        ListView1.SortOrder = lvwDescending
    End If
End If

ListView1.SortKey = (ColumnHeader.Index - 1)






End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then MDIForm1.PopupMenu MDIForm1.mnuFunctions

End Sub
