VERSION 5.00
Begin VB.Form frmCalendar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendar"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   Icon            =   "frmCalendar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4905
   Begin VB.Frame Frame3 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4875
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   255
         Left            =   4020
         TabIndex        =   13
         Top             =   2820
         Width           =   795
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   2280
         TabIndex        =   7
         Top             =   2760
         Width           =   435
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1380
         TabIndex        =   6
         Top             =   2760
         Width           =   435
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   480
         TabIndex        =   5
         Top             =   2760
         Width           =   435
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   3840
         TabIndex        =   3
         Text            =   "2000"
         Top             =   420
         Width           =   915
      End
      Begin VB.CommandButton DayButton 
         Caption         =   "Sun"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   675
      End
      Begin VB.CommandButton DateButton 
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "(24 hour clock)"
         Height          =   195
         Left            =   2820
         TabIndex        =   12
         Top             =   2820
         Width           =   1065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Sec"
         Height          =   195
         Left            =   1920
         TabIndex        =   11
         Top             =   2820
         Width           =   285
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Min"
         Height          =   195
         Left            =   1020
         TabIndex        =   10
         Top             =   2820
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hour"
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Top             =   2820
         Width           =   345
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
         TabIndex        =   8
         Top             =   420
         Width           =   2115
      End
   End
End
Attribute VB_Name = "frmCalendar"
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

Public CalenYear As Integer
Public CalenMonth As Integer
Public CalenDay As Integer
Public ReturnDate As Date



Private Sub Combo5_Click()

CalenMonth = Combo5.ListIndex + 1
UpdateCalender

End Sub

Private Sub Command1_Click()

Dim FirstCheck As Date

'the date/time
d$ = Ts(CalenMonth) + "/" + Ts(CalenDay) + "/" + Ts(CalenYear)
FirstCheck = d$
t$ = Ts(Val(Text2)) + ":" + Ts(Val(Text3)) + ":" + Ts(Val(Text4))
FirstCheck = FirstCheck + t$
ReturnDate = FirstCheck


End Sub

Private Sub DateButton_Click(Index As Integer)

CalenDay = Val(DateButton(Index).Caption)
UpdateCalender


End Sub
Private Sub DayButton_GotFocus(Index As Integer)
Combo5.SetFocus

End Sub

Public Sub DefaultDate()


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


UpdateCalender

End Sub

Private Sub Form_Load()


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

MakeCalendar

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
    'Me.Show
    
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

Private Sub Text6_LostFocus()
b = Int(Val(Text6))
If b > 2100 Then b = 2100
If b < 1900 Then b = 1900
Text6 = Ts(b)
CalenYear = b
UpdateCalender




End Sub
