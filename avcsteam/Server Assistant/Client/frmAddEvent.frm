VERSION 5.00
Begin VB.Form frmAddEvent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add an Event"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "frmAddEvent.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "Event"
      Height          =   615
      Left            =   60
      TabIndex        =   36
      Top             =   240
      Width           =   4875
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   600
         TabIndex        =   38
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   60
         TabIndex        =   37
         Top             =   300
         Width           =   420
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "What to do:"
      Height          =   1215
      Left            =   60
      TabIndex        =   26
      Top             =   5820
      Width           =   4875
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   60
         TabIndex        =   29
         Top             =   840
         Width           =   4755
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   2340
         TabIndex        =   28
         Top             =   240
         Width           =   2475
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   240
         Width           =   2235
      End
      Begin VB.Label Label2 
         Caption         =   "Parameters"
         Height          =   195
         Left            =   60
         TabIndex        =   30
         Top             =   600
         Width           =   2475
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Run this event for the first time at:"
      Height          =   3135
      Left            =   60
      TabIndex        =   18
      Top             =   2640
      Width           =   4875
      Begin VB.CommandButton DateButton 
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1080
         Width           =   675
      End
      Begin VB.CommandButton DayButton 
         Caption         =   "Sun"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   840
         Width           =   675
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   3840
         TabIndex        =   32
         Text            =   "2000"
         Top             =   420
         Width           =   915
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   480
         TabIndex        =   21
         Top             =   2760
         Width           =   435
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1380
         TabIndex        =   20
         Top             =   2760
         Width           =   435
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   2280
         TabIndex        =   19
         Top             =   2760
         Width           =   435
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "October 2000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   180
         TabIndex        =   33
         Top             =   420
         Width           =   2115
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hour"
         Height          =   195
         Left            =   60
         TabIndex        =   25
         Top             =   2820
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Min"
         Height          =   195
         Left            =   1020
         TabIndex        =   24
         Top             =   2820
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Sec"
         Height          =   195
         Left            =   1920
         TabIndex        =   23
         Top             =   2820
         Width           =   285
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "(24 hour clock)"
         Height          =   195
         Left            =   2820
         TabIndex        =   22
         Top             =   2820
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Run this Event under these conditions:"
      Height          =   1695
      Left            =   60
      TabIndex        =   3
      Top             =   900
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Width           =   2235
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   240
         Width           =   1515
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   2340
         TabIndex        =   15
         Text            =   "1"
         Top             =   240
         Width           =   915
      End
      Begin VB.Frame Frame1 
         Caption         =   "Select Days"
         Height          =   1035
         Left            =   60
         TabIndex        =   4
         Top             =   600
         Width           =   4755
         Begin VB.CheckBox CheckDays 
            Caption         =   "Monday"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox CheckDays 
            Caption         =   "Tuesday"
            Height          =   195
            Index           =   1
            Left            =   1620
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox CheckDays 
            Caption         =   "Wednesday"
            Height          =   195
            Index           =   2
            Left            =   3180
            TabIndex        =   12
            Top             =   240
            Width           =   1275
         End
         Begin VB.CheckBox CheckDays 
            Caption         =   "Thursday"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   1095
         End
         Begin VB.CheckBox CheckDays 
            Caption         =   "Friday"
            Height          =   195
            Index           =   4
            Left            =   1620
            TabIndex        =   10
            Top             =   480
            Width           =   1095
         End
         Begin VB.CheckBox CheckDays 
            Caption         =   "Saturday"
            Height          =   195
            Index           =   5
            Left            =   3180
            TabIndex        =   9
            Top             =   480
            Width           =   1095
         End
         Begin VB.CheckBox CheckDays 
            Caption         =   "Sunday"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Select All"
            Height          =   255
            Left            =   1620
            TabIndex        =   7
            Top             =   720
            Width           =   915
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Weekdays"
            Height          =   255
            Left            =   2580
            TabIndex        =   6
            Top             =   720
            Width           =   1035
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Select None"
            Height          =   255
            Left            =   3660
            TabIndex        =   5
            Top             =   720
            Width           =   1035
         End
      End
   End
   Begin VB.CommandButton Command5 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add this Event"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   7080
      Width           =   3375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Please choose the details of this event below."
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4860
   End
End
Attribute VB_Name = "frmAddEvent"
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

Public EditMode As Boolean
Public EditNum As Integer
Dim CalenYear As Integer
Dim CalenMonth As Integer
Dim CalenDay As Integer


Private Sub Combo1_Click()
Update

End Sub


Private Sub Combo3_Click()

If Combo3.ListIndex = 0 Then
    
    Label2 = "Parameters"
    Combo4.Enabled = True
Else
    
    Label2 = "Command"
    Combo4.Enabled = False

End If

End Sub



Private Sub Combo5_Click()

CalenMonth = Combo5.ListIndex + 1
UpdateCalender

End Sub

Private Sub Command1_Click()
    For i = 0 To 6
        CheckDays(i).Value = 1
    Next i
End Sub

Private Sub Command2_Click()
    For i = 0 To 4
        CheckDays(i).Value = 1
    Next i
    For i = 5 To 6
        CheckDays(i).Value = 0
    Next i
End Sub

Private Sub Command3_Click()
    For i = 0 To 6
        CheckDays(i).Value = 0
    Next i
End Sub

Private Sub Command4_Click()

'first, ensure everything is valid.
If Combo1.ListIndex = 1 Or Combo1.ListIndex = 2 Then
    'ensure at least one box is checked
    For i = 0 To 6
        a = a + CheckDays(i)
    Next i
    
    If a = 0 Then MessBox "You must select at least one day on which this script can start!", vbCritical, "Error Adding Event": Exit Sub
End If

If Combo1.ListIndex = 1 And Val(Text1) <= 0 Then
    MessBox "This event cannot run 0 times!", vbCritical, "Error Adding Event": Exit Sub
End If

If Combo1.ListIndex = 1 And Combo2.ListIndex = 4 And Val(Text1) <= 10 Then
    MessBox "Events cannot be run more frequently than once every 11 seconds!", vbCritical, "Error Adding Event": Exit Sub
End If


'If Combo3.ListIndex = 0 And Combo4.ListIndex = -1 Then
'    MessBox "You must select a script to run!", vbCritical, "Error Adding Event": Exit Sub
'End If

If Combo3.ListIndex = 1 And Text5 = "" Then
    MessBox "You must enter an RCON command!", vbCritical, "Error Adding Event": Exit Sub
End If

Text7 = Trim(Text7)
If Text7 = "" Then
    MessBox "You must enter a Name for this event!", vbCritical, "Error Adding Event": Exit Sub
End If

'all set, start filling the array
Dim NewEvent As typEvent
Dim FirstCheck As Date

NewEvent.ComPara = Text5
For i = 0 To 6
    NewEvent.Days(i) = CBool(CheckDays(i))
Next i
NewEvent.Every = Combo2.ListIndex
NewEvent.Mde = Combo1.ListIndex
NewEvent.ScriptName = Combo4.Text
NewEvent.Times = Val(Text1)
NewEvent.WhatToDo = Combo3.ListIndex

'the date/time
d$ = Ts(CalenMonth) + "/" + Ts(CalenDay) + "/" + Ts(CalenYear)
FirstCheck = d$
t$ = Ts(Val(Text2)) + ":" + Ts(Val(Text3)) + ":" + Ts(Val(Text4))
FirstCheck = FirstCheck + t$
NewEvent.FirstCheck = FirstCheck
NewEvent.Name = Text7

'all done, send this to the server
PackageNewEvent NewEvent
Unload Me

End Sub

Private Sub Command5_Click()
Unload Me

End Sub

Private Sub DateButton_Click(Index As Integer)

CalenDay = Val(DateButton(Index).Caption)
UpdateCalender


End Sub

Private Sub DayButton_GotFocus(Index As Integer)
Combo5.SetFocus

End Sub

Private Sub Form_Load()

If EditMode = True Then
    Me.Caption = "Edit Event"
    Command4.Caption = "Update this Event"
End If

'fucked up dates
Dim DDay As String
Dim DMonth As String
Dim DYear As String
Dim FullDate As String

DMonth = Left(Date$, 2)
DDay = Mid(Date$, 4, 2)
DYear = Mid(Date$, 7, 4)

FullDate = DDay + "/" + DMonth + "/" + DYear

CalenDay = DDay
CalenMonth = DMonth
CalenYear = DYear

'If EditMode = False Then Calendar1.Value = FullDate

Combo1.AddItem "Once"
Combo1.AddItem "Every"
Combo1.AddItem "Once on these days..."

Combo2.AddItem "Weeks"
Combo2.AddItem "Days"
Combo2.AddItem "Hours"
Combo2.AddItem "Minutes"
Combo2.AddItem "Seconds"

Combo3.AddItem "Run a script"
Combo3.AddItem "Send RCON command"

Combo5.AddItem "January"
Combo5.AddItem "February"
Combo5.AddItem "March"
Combo5.AddItem "April"
Combo5.AddItem "May"
Combo5.AddItem "June"
Combo5.AddItem "July"
Combo5.AddItem "August"
Combo5.AddItem "September"
Combo5.AddItem "October"
Combo5.AddItem "November"
Combo5.AddItem "December"

Combo1.ListIndex = 0
Combo2.ListIndex = 0
Combo3.ListIndex = 0
Combo2.Enabled = False
Text1.Enabled = False
Frame1.Enabled = False

'Fill the SCRIPTS combobox
For i = 1 To NumCommands
    Combo4.AddItem Commands(i).Name
Next i

If Combo4.ListCount > 0 Then Combo4.ListIndex = 0

MakeCalendar

If EditMode = False Then
    Text2 = Hour(Time)
    Text3 = Minute(Time)
    Text4 = Second(Time)
    
    'Text6 = DYear
    'Combo5.ListIndex = DMonth - 1
Else
    'Fill in the values
    Combo1.ListIndex = Events(EditNum).Mde
    Combo2.ListIndex = Events(EditNum).Every
    Update
    Text1 = Ts(Events(EditNum).Times)
    For i = 0 To 6
        If Events(EditNum).Days(i) Then CheckDays(i).Value = 1
        If Events(EditNum).Days(i) = False Then CheckDays(i).Value = 0
    Next i
    'Calendar1.Value = Events(EditNum).FirstCheck
    CalenDay = Day(Events(EditNum).FirstCheck)
    CalenMonth = Month(Events(EditNum).FirstCheck)
    CalenYear = Year(Events(EditNum).FirstCheck)
    
    Text2 = Ts(Hour(Events(EditNum).FirstCheck))
    Text3 = Ts(Minute(Events(EditNum).FirstCheck))
    Text4 = Ts(Second(Events(EditNum).FirstCheck))
    Text7 = Events(EditNum).Name
    
    Combo3.ListIndex = Events(EditNum).WhatToDo
    Text5 = Events(EditNum).ComPara
    
    j = -1
    'For I = 0 To Combo4.ListCount - 1
    '    'If Combo4.List(I) = Events(EditNum).ScriptName Then j = I: Exit For
    'Next I
    Combo4.Text = Events(EditNum).ScriptName

End If

UpdateCalender

End Sub

Sub MakeCalendar()

'make day buttons

For i = 0 To 6

    If i > 0 Then
        Load DayButton(i)
    End If

    DayButton(i).Left = DayButton(0).Left + ((DayButton(0).Width - Screen.TwipsPerPixelX) * (i Mod 7))
    DayButton(i).Top = DayButton(0).Top
    DayButton(i).Visible = True
    DayButton(i).ZOrder 0
    
    If i = 1 Then DayButton(i).Caption = "Mon"
    If i = 2 Then DayButton(i).Caption = "Tue"
    If i = 3 Then DayButton(i).Caption = "Wed"
    If i = 4 Then DayButton(i).Caption = "Thu"
    If i = 5 Then DayButton(i).Caption = "Fri"
    If i = 6 Then DayButton(i).Caption = "Sat"
    

Next i


'creates the calendar buttons

For i = 0 To 41

    If i > 0 Then
        Load DateButton(i)
    End If

    DateButton(i).Left = DateButton(0).Left + ((DateButton(0).Width - Screen.TwipsPerPixelX) * (i Mod 7))
    DateButton(i).Top = DateButton(0).Top + ((DateButton(0).Height - Screen.TwipsPerPixelY) * (i \ 7))
    DateButton(i).Visible = True
    DateButton(i).ZOrder 0

Next i


End Sub

Sub UpdateCalender()

'updates calender based on selections

Dim TmpDate As Date
Dim TmpDate2 As Date
Dim FirstDay As Date

'month and year
Combo5.ListIndex = CalenMonth - 1
Text6 = Ts(CalenYear)
Label3 = Combo5.List(CalenMonth - 1) + " " + Text6

're-number the calendar
'----------
'start by finding what day the FIRST day of this month is...

TmpDate = Ts(CalenMonth) + "/01/" + Ts(CalenYear)
wk = Weekday(TmpDate)

'lets us find how many days we go back into the previous month
dysbck = wk - 1

If dysbck > 0 Then
    'get the date at this time...
    TmpDate2 = "0" + Ts(dysbck)
    FirstDay = TmpDate - TmpDate2
Else
    FirstDay = TmpDate
End If

'ok, now lets start numbering

'one day
TmpDate = "01"

For i = 0 To 41
    
    dy$ = Ts(Day(FirstDay))
    DateButton(i).Caption = dy$
    
    If Month(FirstDay) <> CalenMonth Then 'we are still in last / next month
        DateButton(i).Enabled = False
    Else
        DateButton(i).Enabled = True
    End If
    
    'finally, if this is the current date, colour the button
    If Day(FirstDay) = CalenDay And Month(FirstDay) = CalenMonth Then
        DateButton(i).BackColor = RGB(255, 255, 0)
    Else
        DateButton(i).BackColor = DayButton(0).BackColor
    End If
    
    'add a day
    FirstDay = FirstDay + TmpDate
Next i



End Sub

Sub Update()

a = Combo1.ListIndex
b = Combo1.ListIndex

If a = 0 Then
    
    Combo2.Enabled = False
    Text1.Enabled = False
    Frame1.Enabled = False
    
ElseIf a = 1 Then

    Combo2.Enabled = True
    Text1.Enabled = True
    Frame1.Enabled = True
    For i = 0 To 6
        CheckDays(i).Value = 1
    Next i

ElseIf a = 2 Then
    
    Combo2.Enabled = False
    Text1.Enabled = False
    Frame1.Enabled = True

End If





End Sub

Private Sub Text1_LostFocus()
Text1 = Trim(str(Int(Val(Text1))))
If Val(Text1) > 30000 Then Text1 = "30000"
End Sub

Private Sub Text2_LostFocus()
Text2 = Trim(str(Int(Val(Text2))))
If Len(Text2) = 1 Then Text2 = "0" + Text2
If Val(Text2) > 23 Then Text2 = "00"

End Sub

Private Sub Text3_LostFocus()

Text3 = Trim(str(Int(Val(Text3))))
If Len(Text3) = 1 Then Text3 = "0" + Text3
If Val(Text3) > 59 Then Text3 = "59"


End Sub

Private Sub Text4_LostFocus()
Text4 = Trim(str(Int(Val(Text4))))
If Len(Text4) = 1 Then Text4 = "0" + Text4
If Val(Text4) > 59 Then Text4 = "59"

End Sub

Private Sub Text5_LostFocus()
Text5 = Trim(Text5)

End Sub

Private Sub Text6_LostFocus()
b = Int(Val(Text6))
If b > 2100 Then b = 2100
If b < 1900 Then b = 1900
Text6 = Ts(b)
CalenYear = b
UpdateCalender




End Sub
