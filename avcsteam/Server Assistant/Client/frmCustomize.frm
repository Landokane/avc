VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomize 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customize Toolbar"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   Icon            =   "frmCustomize.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7170
   Begin VB.CommandButton Command2 
      Height          =   615
      Index           =   8
      Left            =   3600
      Picture         =   "frmCustomize.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Reset Toolbar"
      Top             =   4920
      Width           =   315
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Index           =   7
      Left            =   3300
      Picture         =   "frmCustomize.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Reset Toolbar"
      Top             =   4920
      Width           =   315
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add Scripts to Available"
      Height          =   375
      Left            =   4020
      TabIndex        =   12
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      Top             =   5520
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Index           =   6
      Left            =   3300
      Picture         =   "frmCustomize.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Reset Toolbar"
      Top             =   4260
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Index           =   5
      Left            =   3300
      Picture         =   "frmCustomize.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Move Down"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Index           =   4
      Left            =   3300
      Picture         =   "frmCustomize.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Move Up"
      Top             =   1620
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Index           =   3
      Left            =   3300
      Picture         =   "frmCustomize.frx":198C
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Remove all from Toolbar"
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Index           =   2
      Left            =   3300
      Picture         =   "frmCustomize.frx":1DCE
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Add All to Toolbar"
      Top             =   2940
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Index           =   1
      Left            =   3300
      Picture         =   "frmCustomize.frx":2210
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Remove from Toolbar"
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Index           =   0
      Left            =   3300
      Picture         =   "frmCustomize.frx":2652
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Add to Toolbar"
      Top             =   300
      Width           =   615
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5235
      Left            =   60
      TabIndex        =   1
      Top             =   300
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   9234
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Description"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   5235
      Left            =   3960
      TabIndex        =   2
      Top             =   300
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   9234
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Description"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   900
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   49
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":2A94
            Key             =   ""
            Object.Tag             =   "22"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":3030
            Key             =   ""
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":390C
            Key             =   ""
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":3C28
            Key             =   ""
            Object.Tag             =   "30"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":3F44
            Key             =   ""
            Object.Tag             =   "11"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":4820
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":50FC
            Key             =   ""
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":59D8
            Key             =   ""
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":5E2C
            Key             =   ""
            Object.Tag             =   "13"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":6280
            Key             =   ""
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":66D4
            Key             =   ""
            Object.Tag             =   "19"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":6B28
            Key             =   ""
            Object.Tag             =   "15"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":6F7C
            Key             =   ""
            Object.Tag             =   "20"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":73D0
            Key             =   ""
            Object.Tag             =   "17"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":7824
            Key             =   ""
            Object.Tag             =   "14"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":7C78
            Key             =   ""
            Object.Tag             =   "24"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":80CC
            Key             =   ""
            Object.Tag             =   "12"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":8520
            Key             =   ""
            Object.Tag             =   "10"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":8974
            Key             =   ""
            Object.Tag             =   "8"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":8DC8
            Key             =   ""
            Object.Tag             =   "18"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":921C
            Key             =   ""
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":9670
            Key             =   ""
            Object.Tag             =   "16"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":9ACC
            Key             =   ""
            Object.Tag             =   "29"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":9F20
            Key             =   ""
            Object.Tag             =   "28"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":A374
            Key             =   ""
            Object.Tag             =   "21"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":A7C8
            Key             =   ""
            Object.Tag             =   "27"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":AC1C
            Key             =   ""
            Object.Tag             =   "9"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":B370
            Key             =   ""
            Object.Tag             =   "25"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":B7C4
            Key             =   ""
            Object.Tag             =   "26"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":BAE0
            Key             =   ""
            Object.Tag             =   "23"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":CD64
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":D1C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":D4DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":D7F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":DB12
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":DF64
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":E3B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":E808
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":EB22
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":EE3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":F28E
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":F6E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":FB32
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":FF84
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":103D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":10828
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":10C7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":110CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomize.frx":1151E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Available Toolbar Buttons"
      Height          =   195
      Left            =   4020
      TabIndex        =   3
      Top             =   60
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Current Toolbar Buttons"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1680
   End
End
Attribute VB_Name = "frmCustomize"
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

Sub ShowAvailable()

For i = 1 To ListView2.ListItems.Count
    With ListView2.ListItems.Item(i)

        If .Selected = True Then k = Val(.Index): Exit For
    End With
Next i

ListView2.ListItems.Clear

For i = 1 To UBound(DefaultToolBar)

    'see if we can add this
    
    noadd = 0
    For j = 1 To UBound(CurrentToolBar)
        
        cutj = CurrentToolBar(j).Tag
        If cutj < 31 Then
            If cutj = DefaultToolBar(i).Tag Then noadd = 1
        Else
            If DefaultToolBar(i).Description = CurrentToolBar(j).Description Then noadd = 1
        End If
    Next j
    
    
    If noadd = 0 Then
        b = b + 1
        ListView2.ListItems.Add b, , DefaultToolBar(i).Description, DefaultToolBar(i).IconID, DefaultToolBar(i).IconID
        ListView2.ListItems.Item(b).Tag = DefaultToolBar(i).Tag
        
    End If
    
Next i

j = ListView2.ListItems.Count
ListView2.ListItems.Add j + 1, , "Separator"

For i = 1 To ListView2.ListItems.Count
    With ListView2.ListItems.Item(i)
        .Selected = False
        If k = Val(.Index) Then .Selected = True
    End With
Next i


End Sub

Sub ShowCurrent()

For i = 1 To ListView1.ListItems.Count
    With ListView1.ListItems.Item(i)
        If .Selected = True Then k = Val(.Index): Exit For
    End With
Next i

ListView1.ListItems.Clear
For i = 1 To UBound(CurrentToolBar)

    'see if we can add this
    
    ListView1.ListItems.Add i, , CurrentToolBar(i).Description, CurrentToolBar(i).IconID, CurrentToolBar(i).IconID
    ListView1.ListItems.Item(i).Tag = i
    
Next i

For i = 1 To ListView1.ListItems.Count
    With ListView1.ListItems.Item(i)
        .Selected = False
        If k = Val(.Index) Then .Selected = True
    End With
Next i

End Sub


Private Sub Command1_Click()
ApplyToToolbar
Unload Me


End Sub

Sub AddIconToCurrent(m)

'adds icon at current position

For i = 1 To ListView1.ListItems.Count
    
    With ListView1.ListItems.Item(i)
        If .Selected = True Then k = Val(.Index): Exit For
    End With
Next i

If k = 0 Then k = UBound(CurrentToolBar) + 1

ReDim Preserve CurrentToolBar(0 To UBound(CurrentToolBar) + 1)
n = UBound(CurrentToolBar)

For i = n To k + 1 Step -1
    
    CurrentToolBar(i).Description = CurrentToolBar(i - 1).Description
    CurrentToolBar(i).Tag = CurrentToolBar(i - 1).Tag
    CurrentToolBar(i).Type = CurrentToolBar(i - 1).Type
    CurrentToolBar(i).IconID = CurrentToolBar(i - 1).IconID

Next i

'now add

If m > 0 Then
    CurrentToolBar(k).Description = DefaultToolBar(m).Description
    CurrentToolBar(k).Type = DefaultToolBar(m).Type
    CurrentToolBar(k).IconID = DefaultToolBar(m).IconID
    CurrentToolBar(k).Tag = DefaultToolBar(m).Tag
Else
    CurrentToolBar(k).Description = "Separator"
    CurrentToolBar(k).Type = tbrSeparator
    CurrentToolBar(k).IconID = 0
    CurrentToolBar(k).Tag = 0
End If


End Sub



Sub MoveButtonUp()

'moves selected button up

For i = 1 To ListView1.ListItems.Count
    
    With ListView1.ListItems.Item(i)
        If .Selected = True Then k = Val(.Tag): b = Val(.Index): Exit For
    End With
Next i

n = UBound(CurrentToolBar)

m = k - 1

If k <= 1 Or k <= 0 Then Exit Sub

Swap CurrentToolBar(k).Description, CurrentToolBar(m).Description
Swap CurrentToolBar(k).Tag, CurrentToolBar(m).Tag
Swap CurrentToolBar(k).IconID, CurrentToolBar(m).IconID
Swap CurrentToolBar(k).Type, CurrentToolBar(m).Type

ListView1.ListItems.Item(b).Selected = False
ListView1.ListItems.Item(b - 1).Selected = True

End Sub


Sub MoveButtonDown()

'moves selected button down

For i = 1 To ListView1.ListItems.Count
    
    With ListView1.ListItems.Item(i)
        If .Selected = True Then k = Val(.Tag): b = Val(.Index): Exit For
    End With
Next i

n = UBound(CurrentToolBar)

m = k + 1

If k >= n Then Exit Sub

Swap CurrentToolBar(k).Description, CurrentToolBar(m).Description
Swap CurrentToolBar(k).Tag, CurrentToolBar(m).Tag
Swap CurrentToolBar(k).IconID, CurrentToolBar(m).IconID
Swap CurrentToolBar(k).Type, CurrentToolBar(m).Type

ListView1.ListItems.Item(b).Selected = False
ListView1.ListItems.Item(b + 1).Selected = True

End Sub

Sub RemoveIconFromCurrent(m)

'removes this one

n = UBound(CurrentToolBar)

k = m

If k > 0 Then
    For i = k + 1 To n
        
        CurrentToolBar(i - 1).Description = CurrentToolBar(i).Description
        CurrentToolBar(i - 1).Tag = CurrentToolBar(i).Tag
        CurrentToolBar(i - 1).Type = CurrentToolBar(i).Type
        CurrentToolBar(i - 1).IconID = CurrentToolBar(i).IconID
    
    Next i
    
    ReDim Preserve CurrentToolBar(0 To UBound(CurrentToolBar) - 1)
End If

End Sub




Private Sub Command2_Click(Index As Integer)

If Index = 0 Then ' move over
    For i = 1 To ListView2.ListItems.Count
        
        With ListView2.ListItems.Item(i)
            If .Selected = True Then
                
                'add this type
                
                k = Val(.Tag)
                
                If k > 0 Then
                    For j = 1 To UBound(DefaultToolBar)
                        If Val(DefaultToolBar(j).Tag) = k Then m = j: Exit For
                    Next j
                    
                    'found the icon, now add it
                    AddIconToCurrent m
                Else
                    'adding a seperator
                        
                    AddIconToCurrent 0
                        
                End If
            End If
        End With
        
    Next i
End If


If Index = 1 Then ' remove
    For i = 1 To ListView1.ListItems.Count
        
        With ListView1.ListItems.Item(i)
            If .Selected = True Then
                
                'remove this one
                
                k = Val(.Index)
                
                RemoveIconFromCurrent k
                
                Exit For
            End If
        End With
        
    Next i
End If

If Index = 2 Or Index = 6 Then ' move all over / reset
    CopyDefaultToCurrent
End If

If Index = 3 Then ' remove all
    ReDim CurrentToolBar(0 To 0)
End If

If Index = 4 Then ' move up
    MoveButtonUp
End If


If Index = 5 Then ' move down
    MoveButtonDown
End If

If Index = 7 Then

    For i = 1 To ListView1.ListItems.Count
        
        With ListView1.ListItems.Item(i)
            If .Selected = True Then
                
                'dec icon id
                k = Val(.Index)
                If Val(CurrentToolBar(k).Tag) > 31 Then
                    id = CurrentToolBar(k).IconID - 1
                    If id < 32 Then id = 49
                    CurrentToolBar(k).IconID = id
                End If
                
                Exit For
            End If
        End With
        
    Next i
End If
If Index = 8 Then

    For i = 1 To ListView1.ListItems.Count
        
        With ListView1.ListItems.Item(i)
            If .Selected = True Then
                
                'dec icon id
                k = Val(.Index)
                If Val(CurrentToolBar(k).Tag) > 31 Then
                    id = CurrentToolBar(k).IconID + 1
                    If id > 49 Then id = 32
                    CurrentToolBar(k).IconID = id
                End If
                
                Exit For
            End If
        End With
        
    Next i
End If

ShowAvailable
ShowCurrent

ApplyToToolbar







End Sub

Public Sub AddScripts()

'' add scripts to toolbar list

LoadDefaultToolbar
n = UBound(DefaultToolBar)


For i = 1 To NumCommands
    If Commands(i).ScriptName <> "" Then
        doit = 1
        
        If Commands(i).NumButtons > 0 Then
            If Commands(i).Buttons(1).Type = 3 Then doit = 0
        End If
        
        If doit = 1 Then
            Num = Num + 1
            ReDim Preserve DefaultToolBar(1 To n + Num)
        
            DefaultToolBar(n + Num).Description = "Script - " + Commands(i).ScriptName
            DefaultToolBar(n + Num).Tag = Ts(n + Num)
            'DefaultToolBar(n + Num).Tag =
            DefaultToolBar(n + Num).IconID = 32 + ((i - 1) Mod 18)
            DefaultToolBar(n + Num).Type = 10
        End If
    End If
Next i


ShowAvailable

End Sub

Private Sub Command6_Click()

Open App.Path + "\toolbar.dat" For Append As #1


With MDIForm1.Toolbar1.Buttons

    For i = 1 To .Count
        
        Print #1, "DefaultToolBar(" + Ts(i) + ").Description = " + Chr(34) + .Item(i).Description + Chr(34)
        Print #1, "DefaultToolBar(" + Ts(i) + ").Tag = " + Chr(34) + .Item(i).Tag + Chr(34)
        Print #1, "DefaultToolBar(" + Ts(i) + ").IconID = " + Ts(.Item(i).Image)
        Print #1, "DefaultToolBar(" + Ts(i) + ").Type = " + Ts(.Item(i).Style)
        Print #1, ""
        
    Next i

End With


End Sub

Private Sub Command7_Click()
ShowCurrent
ShowAvailable


End Sub

Private Sub Command3_Click()
ButtonShowMode = 1
SendPacket "BS", ""

End Sub

Private Sub Form_Load()
ShowCurrent
ShowAvailable
End Sub
