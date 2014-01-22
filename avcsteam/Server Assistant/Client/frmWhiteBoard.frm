VERSION 5.00
Begin VB.Form frmWhiteBoard 
   Caption         =   "WhiteBoard"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   14910
   Icon            =   "frmWhiteBoard.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   994
   Begin VB.CommandButton Command5 
      Caption         =   "Delete Object"
      Height          =   315
      Left            =   12420
      TabIndex        =   36
      Top             =   7320
      Width           =   2475
   End
   Begin VB.PictureBox Picture4 
      Height          =   7395
      Left            =   3060
      ScaleHeight     =   489
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   601
      TabIndex        =   32
      Top             =   0
      Width           =   9075
      Begin VB.PictureBox picBoard 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   18000
         Left            =   0
         ScaleHeight     =   1200
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1600
         TabIndex        =   33
         Top             =   0
         Width           =   24000
         Begin VB.TextBox TempText 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   1740
            MultiLine       =   -1  'True
            TabIndex        =   34
            Top             =   720
            Visible         =   0   'False
            Width           =   3435
         End
         Begin VB.Shape Shape 
            Height          =   2715
            Index           =   0
            Left            =   1800
            Shape           =   2  'Oval
            Top             =   2280
            Visible         =   0   'False
            Width           =   4335
         End
         Begin VB.Line Liner 
            Index           =   0
            Visible         =   0   'False
            X1              =   116
            X2              =   296
            Y1              =   372
            Y2              =   372
         End
         Begin VB.Line TempLine 
            Index           =   0
            Visible         =   0   'False
            X1              =   116
            X2              =   240
            Y1              =   388
            Y2              =   388
         End
         Begin VB.Image ImageDa 
            Height          =   1275
            Index           =   0
            Left            =   660
            Top             =   1140
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   540
            TabIndex        =   35
            Top             =   300
            Visible         =   0   'False
            Width           =   480
            WordWrap        =   -1  'True
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3795
      Left            =   0
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.PictureBox Picture5 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2100
         Picture         =   "frmWhiteBoard.frx":08CA
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   38
         Top             =   2100
         Width           =   240
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   0
         ScaleHeight     =   735
         ScaleWidth      =   1215
         TabIndex        =   9
         Top             =   2880
         Width           =   1215
         Begin VB.OptionButton FillMode 
            Height          =   615
            Index           =   1
            Left            =   600
            Picture         =   "frmWhiteBoard.frx":0C0E
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   0
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton FillMode 
            Height          =   615
            Index           =   0
            Left            =   0
            Picture         =   "frmWhiteBoard.frx":14D8
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   2400
         ScaleHeight     =   2055
         ScaleWidth      =   615
         TabIndex        =   12
         Top             =   0
         Width           =   615
         Begin VB.OptionButton LineWidth 
            Height          =   315
            Index           =   5
            Left            =   0
            Picture         =   "frmWhiteBoard.frx":1DA2
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1440
            Width           =   615
         End
         Begin VB.OptionButton LineWidth 
            Height          =   315
            Index           =   4
            Left            =   0
            Picture         =   "frmWhiteBoard.frx":266C
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   1140
            Width           =   615
         End
         Begin VB.OptionButton LineWidth 
            Height          =   315
            Index           =   3
            Left            =   0
            Picture         =   "frmWhiteBoard.frx":2F36
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   900
            Width           =   615
         End
         Begin VB.OptionButton LineWidth 
            Height          =   315
            Index           =   2
            Left            =   0
            Picture         =   "frmWhiteBoard.frx":3800
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   600
            Width           =   615
         End
         Begin VB.OptionButton LineWidth 
            Height          =   315
            Index           =   1
            Left            =   0
            Picture         =   "frmWhiteBoard.frx":40CA
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   300
            Width           =   615
         End
         Begin VB.OptionButton LineWidth 
            Height          =   315
            Index           =   0
            Left            =   0
            Picture         =   "frmWhiteBoard.frx":4994
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   0
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.OptionButton Tools 
         Height          =   615
         Index           =   9
         Left            =   1200
         Picture         =   "frmWhiteBoard.frx":525E
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1200
         Width           =   615
      End
      Begin VB.OptionButton Tools 
         Height          =   615
         Index           =   5
         Left            =   1200
         Picture         =   "frmWhiteBoard.frx":5B28
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton Tools 
         Height          =   615
         Index           =   8
         Left            =   0
         Picture         =   "frmWhiteBoard.frx":63F2
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1800
         Width           =   615
      End
      Begin VB.OptionButton Tools 
         Height          =   615
         Index           =   7
         Left            =   600
         Picture         =   "frmWhiteBoard.frx":6CBC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1200
         Width           =   615
      End
      Begin VB.OptionButton Tools 
         Height          =   615
         Index           =   6
         Left            =   0
         Picture         =   "frmWhiteBoard.frx":7586
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1200
         Width           =   615
      End
      Begin VB.OptionButton Tools 
         Height          =   615
         Index           =   4
         Left            =   600
         Picture         =   "frmWhiteBoard.frx":7E50
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton Tools 
         Height          =   615
         Index           =   3
         Left            =   1200
         Picture         =   "frmWhiteBoard.frx":871A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   615
      End
      Begin VB.OptionButton Tools 
         Height          =   615
         Index           =   2
         Left            =   0
         Picture         =   "frmWhiteBoard.frx":8FE4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton Tools 
         Height          =   615
         Index           =   1
         Left            =   600
         Picture         =   "frmWhiteBoard.frx":98AE
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   615
      End
      Begin VB.OptionButton Tools 
         Height          =   615
         Index           =   0
         Left            =   0
         Picture         =   "frmWhiteBoard.frx":A178
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   50
      Left            =   3060
      SmallChange     =   5
      TabIndex        =   30
      Top             =   7380
      Width           =   9075
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7335
      LargeChange     =   50
      Left            =   12120
      SmallChange     =   5
      TabIndex        =   29
      Top             =   0
      Width           =   255
   End
   Begin VB.ListBox List1 
      Height          =   7275
      Left            =   12420
      TabIndex        =   28
      Top             =   0
      Width           =   2475
   End
   Begin VB.PictureBox fillColour 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1560
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   25
      Top             =   3780
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Control"
      Height          =   1395
      Left            =   0
      TabIndex        =   21
      Top             =   6240
      Width           =   3015
      Begin VB.CommandButton Command6 
         Height          =   195
         Left            =   2400
         TabIndex        =   37
         Top             =   1080
         Width           =   315
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Request Contents"
         Height          =   495
         Left            =   1500
         TabIndex        =   26
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Load Contents"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1500
         TabIndex        =   24
         Top             =   780
         Width           =   1395
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save Contents"
         Height          =   495
         Left            =   60
         TabIndex        =   23
         Top             =   780
         Width           =   1395
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Clear Whiteboard"
         Height          =   495
         Left            =   60
         TabIndex        =   22
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.PictureBox selColour 
      BackColor       =   &H000000C0&
      Height          =   375
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   20
      Top             =   3780
      Width           =   1515
   End
   Begin VB.PictureBox picColours 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2010
      Left            =   0
      Picture         =   "frmWhiteBoard.frx":AA42
      ScaleHeight     =   134
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   19
      Top             =   4200
      Width           =   3015
   End
End
Attribute VB_Name = "frmWhiteBoard"
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

Dim LneWidth As Integer
Dim Filled As Boolean
Dim CurrTool As Integer

Dim MouseDn As Boolean
Dim FlipMode As Boolean
Dim Pos1X As Integer
Dim Pos1Y As Integer
Dim Pos2X As Integer
Dim Pos2Y As Integer

Dim PointData As String
Dim MoveObj As Object
Dim MovingNow As Boolean
Public OldWindowProc As Long
Private Sub Command1_Click()

SendPacket "CB", ""

End Sub

Private Sub Command2_Click()
a$ = InBox("Name:")

Open App.Path + "\" + a$ + ".dat" For Binary As #1
Put #1, , NumShapes
Put #1, , Shapes
Close #1
End Sub

Private Sub Command4_Click()

SendPacket "AS", ""
ClearBoard ""

End Sub

Private Sub Command5_Click()
'
i = List1.ListIndex
If i = -1 Then Exit Sub
j = List1.ItemData(i)

SendPacket "DS", Ts(Shapes(j).ShapeID)


'Open "e:\vbproj~1\udp\wb1.dat" For Binary As #1
'Put #1, , NumShapes
'Put #1, , Shapes
'Close #1

End Sub

Private Sub Command6_Click()
'scan picture

'PackageNewShape CurrTool, Liner(0).BorderColor, -1, Liner(0).BorderWidth, Liner(0).X1, Liner(0).Y1, Liner(0).X2, Liner(0).Y2, ""

'For Y = 1 To Picture1.ScaleHeight
 '   For X = 1 To Picture5.ScaleWidth
Dim c As Long

basex = 50
basey = 50

For Y = 1 To Picture1.ScaleHeight
    For X = 1 To Picture5.ScaleWidth
    
         c = Picture5.Point(X, Y)
         
         X1 = (X - 1) * 20
         Y1 = (Y - 1) * 20
         
         PackageNewShape 2, c, c, 1, basex + CInt(X1), basey + CInt(Y1), 20, 20, ""
    Next X
Next Y


End Sub

Private Sub FillMode_Click(Index As Integer)

If Index = 0 Then Filled = True
If Index = 1 Then Filled = False


End Sub

Private Sub Form_Load()

Filled = False
CurrTool = 0
LneWidth = 1

ShowWhiteBoard = True
SendPacket "AS", ""


End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then Exit Sub

If Me.WindowState = 0 Then
If Me.Height > Screen.TwipsPerPixelY * 1200 Then Me.Height = Screen.TwipsPerPixelY * 1200
If Me.Height < 8055 Then Me.Height = 8055

If Me.Width > Screen.TwipsPerPixelX * 1600 Then Me.Width = Screen.TwipsPerPixelX * 1600
If Me.Width < 7620 Then Me.Width = 7620
End If
w = Me.ScaleWidth
h = Me.ScaleHeight

w2 = w - picColours.Width - List1.Width - VScroll1.Width - 8
Picture4.Width = w2
VScroll1.Left = Picture4.Left + Picture4.Width + 2
List1.Left = VScroll1.Left + VScroll1.Width + 2

Picture4.Height = h - HScroll1.Height - 4
HScroll1.Top = Picture4.Top + Picture4.Height + 2

VScroll1.Height = Picture4.Height
HScroll1.Width = Picture4.Width

List1.Height = Picture4.Height
Me.Caption = Ts(w) + " (" + Ts(Me.Width) + ") - " + Ts(h) + "(" + Ts(Me.Height) + ")"


VScroll1.Max = picBoard.Height - Picture4.Height
HScroll1.Max = picBoard.Width - Picture4.Width

Command5.Left = List1.Left + 1
Command5.Top = HScroll1.Top - 2




End Sub

Private Sub Form_Unload(Cancel As Integer)
ShowWhiteBoard = False

End Sub

Private Sub HScroll1_Change()

picBoard.Left = -HScroll1.Value
picBoard.Refresh
End Sub

Private Sub HScroll1_Scroll()
picBoard.Left = -HScroll1.Value
picBoard.Refresh
End Sub

Private Sub ImageDa_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

picBoard_MouseDown 9, Index, (X / Screen.TwipsPerPixelX) + ImageDa(Index).Left, (Y / Screen.TwipsPerPixelY) + ImageDa(Index).Top


End Sub

Private Sub ImageDa_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
picBoard_MouseMove 9, Index, (X / Screen.TwipsPerPixelX) + ImageDa(Index).Left, (Y / Screen.TwipsPerPixelY) + ImageDa(Index).Top

End Sub

Private Sub ImageDa_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
picBoard_MouseUp 9, Index, (X / Screen.TwipsPerPixelX) + ImageDa(Index).Left, (Y / Screen.TwipsPerPixelY) + ImageDa(Index).Top

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label_DblClick(Index As Integer)

If CurrTool = 0 Then
    
    'start text edit mode
    
    
    TempText.Top = Label(Index).Top
    TempText.Left = Label(Index).Left
    TempText.Width = Label(Index).Width
    TempText.Height = Label(Index).Height
    TempText.ForeColor = Label(Index).ForeColor
    TempText.FontSize = Label(Index).FontSize

    If Label(Index).BackStyle = 1 Then TempText.BackColor = Label(Index).BackColor
    If Label(Index).BackStyle = 0 Then TempText.BackColor = RGB(255, 255, 255)


    TempText = Label(Index).Caption
    TempText.Tag = Label(Index).Tag
    TempText.ZOrder 0
    TempText.Visible = True
    TempText.SetFocus
End If


End Sub

Private Sub Label_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
picBoard_MouseDown 10, Index, (X / Screen.TwipsPerPixelX) + Label(Index).Left, (Y / Screen.TwipsPerPixelY) + Label(Index).Top

End Sub

Private Sub Label_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
picBoard_MouseMove 10, Index, (X / Screen.TwipsPerPixelX) + Label(Index).Left, (Y / Screen.TwipsPerPixelY) + Label(Index).Top

End Sub

Private Sub Label_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
picBoard_MouseUp 10, Index, (X / Screen.TwipsPerPixelX) + Label(Index).Left, (Y / Screen.TwipsPerPixelY) + Label(Index).Top

End Sub

Private Sub LineWidth_Click(Index As Integer)
    
    LneWidth = Index + 1
End Sub


Private Sub picColours_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

c = picColours.Point(X, Y)

If Button = 1 Then
    selColour.BackColor = c
End If
If Button = 2 Then
    fillColour.BackColor = c
End If

End Sub

Private Sub picBoard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If TempText.Visible = True Then

    DoneTextEdit
    
    Exit Sub
End If


MouseDn = True

If CurrTool = 0 Then 'arrow - move

    If Button >= 9 Then
        
        If Button = 9 Then Set MoveObj = ImageDa(Shift): MovingNow = True
        If Button = 10 Then Set MoveObj = Label(Shift): MovingNow = True
        
        If MovingNow Then
            Pos1X = X
            Pos1Y = Y
            
            Pos2X = MoveObj.Left
            Pos2Y = MoveObj.Top
        End If
    End If

ElseIf CurrTool = 1 Then 'line
    Liner(0).X1 = X
    Liner(0).Y1 = Y
    
    Pos1X = X
    Pos1Y = Y
    
    Liner(0).X2 = X
    Liner(0).Y2 = Y
    
    Liner(0).BorderColor = selColour.BackColor
    Liner(0).BorderWidth = LneWidth
    
    Liner(0).ZOrder 0
    Liner(0).Visible = True

ElseIf CurrTool >= 2 And CurrTool <= 4 Then
    
    Shape(0).Left = X
    Shape(0).Top = Y
    
    Pos1X = X
    Pos1Y = Y
    
    Shape(0).Width = 1
    Shape(0).Height = 1
    
    Shape(0).BorderColor = selColour.BackColor
    Shape(0).BorderStyle = 1
    Shape(0).FillColor = fillColour.BackColor
    Shape(0).BorderWidth = LneWidth
    
    If Filled Then Shape(0).FillStyle = 0
    If Not Filled Then Shape(0).FillStyle = 1
    
    If CurrTool = 2 Then Shape(0).Shape = 0
    If CurrTool = 3 Then Shape(0).Shape = 2
    If CurrTool = 4 Then Shape(0).Shape = 4
    
    Shape(0).ZOrder 0
    Shape(0).Visible = True
ElseIf CurrTool = 5 Then ' pencil

    PointData = ""
    
    Pos1X = X
    Pos1Y = Y
    Me.Caption = 0
    
    PointData = PointData + ConvertPoint(CInt(X)) + "," + ConvertPoint(CInt(Y)) + ","
    
ElseIf CurrTool = 6 Then
    
    'Form1.Dlg1.FileT = ""
    
'    e = InStrRev(Form1.Dlg1.FileName, "\")
'    If e > 1 Then b$ = Left(Form1.Dlg1.FileName, e - 1)
'    Form1.Dlg1.FileName = ""
'    Form1.Dlg1.InitDir = b$ + "\"
'
'
'    Form1.Dlg1.DialogTitle = "Select Graphic file (must be less than 20 k)"
'    Form1.Dlg1.Filter = "Image Files|*.bmp;gif;jpg"
'    Form1.Dlg1.MaxFileSize = 20480
'
'    Form1.Dlg1.ShowOpen

    Dim Filter$, Flags&, FileName$
    
    Flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or _
            OFN_PATHMUSTEXIST
            
    Filter$ = "Image Files" + Chr(0) + "*.bmp;gif;jpg"
    
    FileName = ShowOpen(Filter, Flags, Me.hwnd)
        
    a$ = FileName
    
    If a$ = "" Then Exit Sub
    If FileLen(a$) > 20480 Then MessBox "Too big, must be 20k or less": Exit Sub
   
    'add this file
    
    If CheckForFile(a$) Then
    
        mn = FileLen(a$)
        Label1 = mn
        
        h = FreeFile
        Open a$ For Binary As h
            
            ' read the file
            ret$ = Input(50000, #h)
        Close h
    End If
    'send the image data
    PackageNewShape CurrTool, -1, -1, 0, CInt(X), CInt(Y), 0, 0, ret$
    
    Tools_Click 0
    Tools(0).Value = True
    
ElseIf CurrTool = 7 Then

    ' text tool
   
    Shape(0).Left = X
    Shape(0).Top = Y
    
    Pos1X = X
    Pos1Y = Y
    
    Shape(0).Width = 1
    Shape(0).Height = 1
    
    Shape(0).BorderColor = RGB(0, 0, 0)
    Shape(0).BorderWidth = LneWidth
    Shape(0).BorderStyle = 3
    
    Shape(0).FillStyle = 1
    
    Shape(0).Shape = 0
    Shape(0).ZOrder 0
    Shape(0).Visible = True

ElseIf CurrTool = 8 Then 'eraser - erase

    If Button >= 9 Then
        
        If Button = 9 Then Set MoveObj = ImageDa(Shift): MovingNow = True
        If Button = 10 Then Set MoveObj = Label(Shift): MovingNow = True
        

        SendPacket "DS", MoveObj.Tag
    
    End If
    
ElseIf CurrTool = 9 Then 'paste


    a$ = Clipboard.GetText
    If a$ <> "" Then Exit Sub
    
    
    ImageDa(0) = Clipboard.GetData

    SavePicture ImageDa(0).Picture, App.Path + "\wbdata\temp.bmp"

    a$ = App.Path + "\wbdata\temp.bmp"

    If a$ = "" Then Exit Sub
    If FileLen(a$) > 20480 Then MessBox "Too big, must be 20k or less": Exit Sub

    'add this file

    If CheckForFile(a$) Then

        mn = FileLen(a$)
        Label1 = mn

        h = FreeFile
        Open a$ For Binary As h

            ' read the file
            ret$ = Input(50000, #h)
        Close h
    End If
    'send the image data
    PackageNewShape 6, -1, -1, 0, CInt(X), CInt(Y), 0, 0, ret$

    Tools_Click 0
    Tools(0).Value = True

    
End If


End Sub

Function ConvertPoint(pt As Integer) As String
    'converts to proper point
    
    nt = pt \ 200
    nt2 = pt - (200 * nt)
    
    If nt2 < 0 Then nt2 = 0
        
    a$ = Chr(nt + 30) + Chr(nt2 + 30)
    
    If Len(a$) < 2 Then MsgBox "problem"
    
    ConvertPoint = a$

End Function

Function ConvertBack(dat As String) As Integer
    'converts to proper point
    
    If Len(dat) < 2 Then
        MsgBox "error"
    Else
    
        nt = Asc(Mid(dat, 1, 1))
        nt2 = Asc(Mid(dat, 2, 1)) - 30
        nt2 = nt2 + 200 * (nt - 30)
        
        ConvertBack = nt2
    End If
End Function

Private Sub picBoard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If MouseDn Then
    
    If CurrTool = 0 Then
       'moving
        If MovingNow Then
        
            disx = Pos1X - X
            disy = Pos1Y - Y
            On Error Resume Next
            
            MoveObj.Top = Pos2Y - disy
            MoveObj.Left = Pos2X - disx
        
        End If
    
    ElseIf CurrTool = 1 Then
        
        Liner(0).X2 = X
        Liner(0).Y2 = Y
        
    ElseIf CurrTool >= 2 And CurrTool <= 4 Then
        
        l = Pos1X
        t = Pos1Y
        
        w = X - Pos1X
        h = Y - Pos1Y
            
        If w < 0 Then w = Pos1X - X: l = X
        If h < 0 Then h = Pos1Y - Y: t = Y
        
        Shape(0).Left = l
        Shape(0).Top = t
        Shape(0).Width = w
        Shape(0).Height = h
        

    ElseIf CurrTool = 5 Then ' pencil
    
        d = Sqr((Pos1X - X) ^ 2 + (Pos1Y - Y) ^ 2)
        If d >= 5 And Len(PointData) < 3000 Then
        
            'draw line
            
            'picBoard.DrawWidth = LneWidth
            'picBoard.Line (Pos1X, Pos1Y)-(X, Y), selColour.BackColor
                    
            i = TempLine.Count
            Load TempLine(i)
                        
            TempLine(i).BorderWidth = LneWidth
            TempLine(i).BorderColor = selColour.BackColor
            TempLine(i).X1 = Pos1X
            TempLine(i).Y1 = Pos1Y
            TempLine(i).X2 = X
            TempLine(i).Y2 = Y
            TempLine(i).ZOrder 0
            TempLine(i).Visible = True
            
            Pos1X = X
            Pos1Y = Y
            
            
            PointData = PointData + ConvertPoint(CInt(X)) + "," + ConvertPoint(CInt(Y)) + ","
            Me.Caption = Len(PointData)
        End If
    ElseIf CurrTool = 7 Then
    
        l = Pos1X
        t = Pos1Y
        
        w = X - Pos1X
        h = Y - Pos1Y
            
        If w < 0 Then w = Pos1X - X: l = X
        If h < 0 Then h = Pos1Y - Y: t = Y
        
        Shape(0).Left = l
        Shape(0).Top = t
        Shape(0).Width = w
        Shape(0).Height = h
        
    End If


End If


End Sub

Private Sub picBoard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If MouseDn Then
    MouseDn = False
    
    ' add the shape
    If CurrTool = 0 Then
        
        If MovingNow Then

            On Error Resume Next
            'update
            a$ = Chr(251)
            a$ = a$ + MoveObj.Tag + Chr(250)
            a$ = a$ + Ts(MoveObj.Left) + Chr(250)
            a$ = a$ + Ts(MoveObj.Top) + Chr(250)
            a$ = a$ + Chr(251)
        
            SendPacket "SM", a$
            
            Set MoveObj = Nothing
            MovingNow = False
                    
        End If
    
    ElseIf CurrTool = 1 Then
        
        Liner(0).Visible = False
        
        PackageNewShape CurrTool, Liner(0).BorderColor, -1, Liner(0).BorderWidth, Liner(0).X1, Liner(0).Y1, Liner(0).X2, Liner(0).Y2, ""
        
    ElseIf CurrTool >= 2 And CurrTool <= 4 Then
        
        Shape(0).Visible = False
        If Filled Then PackageNewShape CurrTool, Shape(0).BorderColor, Shape(0).FillColor, Shape(0).BorderWidth, Shape(0).Left, Shape(0).Top, Shape(0).Width, Shape(0).Height, ""
        If Not Filled Then PackageNewShape CurrTool, Shape(0).BorderColor, -1, Shape(0).BorderWidth, Shape(0).Left, Shape(0).Top, Shape(0).Width, Shape(0).Height, ""
        
    ElseIf CurrTool = 5 Then ' pencil
    
        PackageNewShape CurrTool, selColour.BackColor, -1, LneWidth, 0, 0, 0, 0, PointData
        PointData = ""
            
        'remove temp lines
        
        For i = 1 To TempLine.Count - 1
            Unload TempLine(i)
        Next i
    
    ElseIf CurrTool = 7 Then 'text tool
        
        Shape(0).Visible = False
            
        'make the text box visible
        
        TempText.Top = Shape(0).Top
        TempText.Left = Shape(0).Left
        TempText.Width = Shape(0).Width
        TempText.Height = Shape(0).Height
        TempText.ForeColor = selColour.BackColor
        
        TempText.FontSize = 8 + LneWidth * 4
        
        If Filled Then TempText.BackColor = fillColour.BackColor
        If Not Filled Then TempText.BackColor = RGB(255, 255, 255)
        
        
        TempText = ""
        TempText.Tag = "-1"
        TempText.Visible = True
        TempText.SetFocus
        
    End If
End If

End Sub

Sub DoneTextEdit()

If TempText.Tag = "-1" Then
    'new text mode

    
    TempText.Visible = False
    
    If TempText.Text <> "" Then
    
        If Filled Then PackageNewShape 7, TempText.ForeColor, TempText.BackColor, CInt(TempText.FontSize), TempText.Left, TempText.Top, TempText.Width, TempText.Height, TempText.Text
        If Not Filled Then PackageNewShape 7, TempText.ForeColor, -1, CInt(TempText.FontSize), TempText.Left, TempText.Top, TempText.Width, TempText.Height, TempText.Text

    End If
Else
    'editing old text mode
    
    TempText.Visible = False
    
    'tell server we changed the text
    
    a$ = Chr(251)
    a$ = a$ + TempText.Tag + Chr(250)
    a$ = a$ + TempText.Text + Chr(250)
    a$ = a$ + Chr(251)

    SendPacket "TC", a$

End If



End Sub

Private Sub Tools_Click(Index As Integer)

CurrTool = Index

End Sub

Sub PackageNewShape(Typ As Integer, LineColour As Long, fillColour As Long, _
    LineWidth As Integer, Pos1X As Integer, Pos1Y As Integer, _
    Pos2X As Integer, Pos2Y As Integer, ExtraData As String)

' Creates a new shape and sends it to the server

Dim NewShape As typWhiteboard

NewShape.ObjType = Typ
NewShape.LineColour = LineColour
NewShape.fillColour = fillColour
NewShape.LineWidth = LineWidth
NewShape.Pos1X = Pos1X
NewShape.Pos1Y = Pos1Y
NewShape.Pos2X = Pos2X
NewShape.Pos2Y = Pos2Y
NewShape.ExtraData = ExtraData

Randomize
NewShape.ShapeID = Int(Rnd * 30000) + 1

' all done, now send

a$ = ""

a$ = Chr(251)
a$ = a$ + Ts(NewShape.ObjType) + Chr(250)
a$ = a$ + Ts(NewShape.LineColour) + Chr(250)
a$ = a$ + Ts(NewShape.fillColour) + Chr(250)
a$ = a$ + Ts(NewShape.LineWidth) + Chr(250)
a$ = a$ + Ts(NewShape.Pos1X) + Chr(250)
a$ = a$ + Ts(NewShape.Pos1Y) + Chr(250)
a$ = a$ + Ts(NewShape.Pos2X) + Chr(250)
a$ = a$ + Ts(NewShape.Pos2Y) + Chr(250)
b$ = NewShape.ExtraData
If b$ <> "" Then b$ = Convert255(b$)
a$ = a$ + b$ + Chr(250)
a$ = a$ + Ts(NewShape.ShapeID) + Chr(250)

a$ = a$ + Chr(251)

'send the data
SendPacket "NS", a$

End Sub


Public Sub NewShapeLoad(p$)
On Error Resume Next
'load up the new shape and display it


Dim NewShape As typWhiteboard

f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then
                    NumShapes = NumShapes + 1
                    ReDim Preserve Shapes(0 To NumShapes)
                    Shapes(NumShapes).ObjType = Val(m$)
                End If
                    
                If j = 2 Then Shapes(NumShapes).LineColour = Val(m$)
                If j = 3 Then Shapes(NumShapes).fillColour = Val(m$)
                If j = 4 Then Shapes(NumShapes).LineWidth = Val(m$)
                If j = 5 Then Shapes(NumShapes).Pos1X = Val(m$)
                If j = 6 Then Shapes(NumShapes).Pos1Y = Val(m$)
                If j = 7 Then Shapes(NumShapes).Pos2X = Val(m$)
                If j = 8 Then Shapes(NumShapes).Pos2Y = Val(m$)
                
                
                If j = 9 Then 'decode the data
                    If m$ <> "" Then
                        Shapes(NumShapes).ExtraData = DeCode255(m$)
                    End If
                End If
                If j = 10 Then Shapes(NumShapes).ShapeID = Val(m$)
                If j = 11 Then Shapes(NumShapes).Creator = m$
            End If
        Loop Until h = 0
        
        'now add this shape to the whiteboard data
        DisplayShape NumShapes

    End If
Loop Until f = 0 Or e = 0




'Shapes(NumShapes).ExtraData = NewShape.ExtraData
'Shapes(NumShapes).fillColour = NewShape.fillColour
'Shapes(NumShapes).LineColour = NewShape.LineColour
'Shapes(NumShapes).LineWidth = NewShape.LineWidth
'Shapes(NumShapes).ObjType = NewShape.ObjType
'Shapes(NumShapes).Pos1X = NewShape.Pos1X
'Shapes(NumShapes).Pos1Y = NewShape.Pos1Y
'Shapes(NumShapes).Pos2X = NewShape.Pos2X
'Shapes(NumShapes).Pos2Y = NewShape.Pos2Y
'Shapes(NumShapes).ShapeID = NewShape.ShapeID


End Sub


Sub DisplayShape(Num)

'displays this shape

With Shapes(Num)


    If .ObjType = 1 Then
        
        ' create a new shape here
        
        i = Liner.UBound
        i = i + 1
        Load Liner(i)
        
        Liner(i).X1 = .Pos1X
        Liner(i).Y1 = .Pos1Y
        Liner(i).X2 = .Pos2X
        Liner(i).Y2 = .Pos2Y
        
        Liner(i).BorderColor = .LineColour
        Liner(i).BorderWidth = .LineWidth
        Liner(i).Tag = Ts(.ShapeID)
            
        Liner(i).ZOrder 0
        Liner(i).Visible = True
        
        List1.AddItem .Creator + ": New Line"
        List1.ListIndex = List1.NewIndex
        
    ElseIf .ObjType >= 2 And .ObjType <= 4 Then
        
        ' create a new shape here
        
        i = Shape.UBound
        i = i + 1
        Load Shape(i)
        
        Shape(i).Top = .Pos1Y
        Shape(i).Left = .Pos1X
        Shape(i).Width = .Pos2X
        Shape(i).Height = .Pos2Y
        Shape(i).BorderStyle = 1
        Shape(i).BorderColor = .LineColour
        Shape(i).BorderWidth = .LineWidth
        If .fillColour >= 0 Then Shape(i).BackColor = .fillColour
        Shape(i).Tag = Ts(.ShapeID)
    
        If .fillColour >= 0 Then Shape(i).BackStyle = 1
        If .fillColour < 0 Then Shape(i).BackStyle = 0
        
    
        If .ObjType = 2 Then Shape(i).Shape = 0: List1.AddItem .Creator + ": New Rectangle"
        If .ObjType = 3 Then Shape(i).Shape = 2: List1.AddItem .Creator + ": New Ellipse"
        If .ObjType = 4 Then Shape(i).Shape = 4: List1.AddItem .Creator + ": New Rounded Rect"
        
        Shape(i).ZOrder 0
        Shape(i).Visible = True
        
        
        List1.ListIndex = List1.NewIndex
        
    ElseIf .ObjType = 5 Then       'pencil data
    
        ' create the pencil data
        'format:
        ' XX,YY,XX,YY,XX,YY,XX,YY ...
        
        f = 0
        Dim TheData() As String
        
        b$ = .ExtraData
        b$ = Left(b$, Len(b$) - 1)
        
        TheData = Split(b$, ",")
        
        'MsgBox "First One: " + Ts(Len(b$) + 1) + "   Second One: " + Ts(Len(c$))
        
        For j = 0 To UBound(TheData) - 3 Step 2
            errs = 0
            
            d1$ = Mid(b$, ((j) * 3) + 1, 2)
            If Len(d1$) < 2 Then errs = 1
            
            d2$ = Mid(b$, ((j + 1) * 3) + 1, 2)
            If Len(d2$) < 2 Then errs = 1
            
            d3$ = Mid(b$, ((j + 2) * 3) + 1, 2)
            If Len(d3$) < 2 Then errs = 1
            
            d4$ = Mid(b$, ((j + 3) * 3) + 1, 2)
            If Len(d4$) < 2 Then errs = 1

            
            If errs = 0 Then
                i = Liner.UBound
                i = i + 1
                Load Liner(i)
                Liner(i).X1 = ConvertBack(d1$)
                Liner(i).Y1 = ConvertBack(d2$)
                Liner(i).X2 = ConvertBack(d3$)
                Liner(i).Y2 = ConvertBack(d4$)
                
                Liner(i).BorderColor = .LineColour
                Liner(i).BorderWidth = .LineWidth
                Liner(i).Tag = Ts(.ShapeID)
                    
                Liner(i).ZOrder 0
                Liner(i).Visible = True
            End If
        Next j
    
        List1.AddItem .Creator + ": New Pencil Drawing"
        List1.ListIndex = List1.NewIndex
    
    ElseIf .ObjType = 6 Then       'image data
        
        'save the pic
        
        i = ImageDa.UBound
        i = i + 1
        Load ImageDa(i)
        
        d$ = App.Path + "\wbdata"
        If Dir(d$, vbDirectory) = "" Then MkDir d$
        
        d$ = d$ + "\" + Ts(.ShapeID) + ".dat"
        
        h = FreeFile
        Open d$ For Binary As h
            Put #h, , .ExtraData
        Close h
        
        Label2 = Len(.ExtraData)
        
        ImageDa(i).Picture = LoadPicture(d$)
        ImageDa(i).Left = .Pos1X
        ImageDa(i).Top = .Pos1Y
        ImageDa(i).Visible = True
        ImageDa(i).ZOrder 0
        ImageDa(i).Tag = Ts(.ShapeID)
        
        List1.AddItem .Creator + ": New Image"
        List1.ListIndex = List1.NewIndex
        
    ElseIf .ObjType = 7 Then       'text
        
        i = Label.UBound
        i = i + 1
        Load Label(i)
        
        Label(i).Left = .Pos1X
        Label(i).Top = .Pos1Y
        Label(i).Width = .Pos2X
        Label(i).Height = .Pos2Y
        Label(i).Tag = Ts(.ShapeID)
      
        If .fillColour = -1 Then Label(i).BackStyle = 0
        If .fillColour > -1 Then
            Label(i).BackStyle = 1
            Label(i).BackColor = .fillColour
        End If
        
        Label(i).ForeColor = .LineColour
        Label(i).FontSize = .LineWidth
        Label(i).Caption = .ExtraData
        
        Label(i).Visible = True
        Label(i).ZOrder 0
        
        List1.AddItem .Creator + ": New Text"
        List1.ListIndex = List1.NewIndex
        
    End If

End With

List1.ItemData(List1.NewIndex) = Num
DoEvents

End Sub

Sub ReList()

List1.Clear
For jj = 1 To NumShapes

    With Shapes(jj)
    
        If .ObjType = 1 Then
            
            List1.AddItem .Creator + ": New Line"
        
        ElseIf .ObjType >= 2 And .ObjType <= 4 Then
            
            If .ObjType = 2 Then List1.AddItem .Creator + ": New Rectangle"
            If .ObjType = 3 Then List1.AddItem .Creator + ": New Ellipse"
            If .ObjType = 4 Then List1.AddItem .Creator + ": New Rounded Rect"
            
        ElseIf .ObjType = 5 Then       'pencil data
        
             List1.AddItem .Creator + ": New Pencil Drawing"
        
        ElseIf .ObjType = 6 Then       'image data
            List1.AddItem .Creator + ": New Image"
        ElseIf .ObjType = 7 Then       'text
            List1.AddItem .Creator + ": New Text"
        End If
        
        List1.ItemData(List1.NewIndex) = jj
    End With
Next jj



End Sub

Public Sub ClearBoard(p$)
On Error Resume Next

List1.Clear
Me.Caption = "Clearing..."

If p$ <> "" Then
    List1.AddItem p$ + " cleared."
    List1.ListIndex = List1.NewIndex
End If

ReDim Shapes(0 To 0)
NumShapes = 0

For i = 1 To Liner.UBound
    Unload Liner(i)
Next i
   picBoard.Refresh

For i = 1 To Shape.UBound
    Unload Shape(i)
Next i
   picBoard.Refresh

For i = 1 To ImageDa.UBound
    Unload ImageDa(i)
Next i
    picBoard.Refresh


For i = 1 To Label.UBound
    Unload Label(i)
Next i
   picBoard.Refresh

Me.Caption = "Whiteboard"
End Sub

Public Sub ChangeText(p$)
On Error Resume Next
'change the contents of a textbox
f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then shpid = Val(m$)
                If j = 2 Then newtxt$ = m$

            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0


j = 0
For i = 1 To NumShapes
    If Shapes(i).ShapeID = shpid Then j = i: Exit For
Next i

If j > 0 Then
    Shapes(i).ExtraData = newtxt$
        
    For i = 1 To Label.UBound
        If Val(Label(i).Tag) = shpid Then
            Label(i).Caption = newtxt$
        End If
    Next i
End If

End Sub


Public Sub MoveObject(p$)
On Error Resume Next
'moves a shape
f = 0
i = 0
Do

    e = InStr(f + 1, p$, Chr(251))
    f = InStr(e + 1, p$, Chr(251))
    'extract this section
    
    If e > 0 And f > e Then
        a$ = Mid(p$, e + 1, f - e - 1)
        i = i + 1
            
        h = 0
        j = 0
        Do
            g = h
            h = InStr(g + 1, a$, Chr(250))
            g = g + 1
            j = j + 1
            If g > 0 And h > g - 1 Then
                m$ = Mid(a$, g, h - g)
                
                If j = 1 Then shpid = Val(m$)
                If j = 2 Then newx = Val(m$)
                If j = 3 Then newy = Val(m$)

            End If
        Loop Until h = 0
    
    End If
Loop Until f = 0 Or e = 0



j = 0
For i = 1 To NumShapes
    If Shapes(i).ShapeID = shpid Then j = i: Exit For
Next i

If j > 0 Then
    Shapes(i).Pos1X = newx
    Shapes(i).Pos1Y = newy
    On Error Resume Next
    
    For i = 1 To Label.UBound
        If Val(Label(i).Tag) = shpid Then
            Label(i).Left = newx
            Label(i).Top = newy
        End If
    Next i
    For i = 1 To ImageDa.UBound
        If Val(ImageDa(i).Tag) = shpid Then
            ImageDa(i).Left = newx
            ImageDa(i).Top = newy
        End If
    Next i
End If

End Sub

Public Sub DeleteObject(p$)
On Error Resume Next
shpid = Val(p$)

j = 0
For i = 1 To NumShapes
    If Shapes(i).ShapeID = shpid Then j = i: Exit For
Next i

If j > 0 Then
        
   'delete
    For i = j To NumShapes - 1
        
        Shapes(i).ExtraData = Shapes(i + 1).ExtraData
        Shapes(i).fillColour = Shapes(i + 1).fillColour
        Shapes(i).LineColour = Shapes(i + 1).LineColour
        Shapes(i).LineWidth = Shapes(i + 1).LineWidth
        Shapes(i).ObjType = Shapes(i + 1).ObjType
        Shapes(i).Pos1X = Shapes(i + 1).Pos1X
        Shapes(i).Pos1Y = Shapes(i + 1).Pos1Y
        Shapes(i).Pos2X = Shapes(i + 1).Pos2X
        Shapes(i).Pos2Y = Shapes(i + 1).Pos2Y
        Shapes(i).ShapeID = Shapes(i + 1).ShapeID
        
    Next i
    NumShapes = NumShapes - 1
    ReDim Preserve Shapes(0 To NumShapes)
    
    For i = 1 To Label.UBound
        If Val(Label(i).Tag) = shpid Then
            Unload Label(i)
        End If
    Next i
    For i = 1 To ImageDa.UBound
        If Val(ImageDa(i).Tag) = shpid Then
           Unload ImageDa(i)
        End If
    Next i
    For i = 1 To Shape.UBound
        If Val(Shape(i).Tag) = shpid Then
           Unload Shape(i)
        End If
    Next i
    For i = 1 To Liner.UBound
        If Val(Liner(i).Tag) = shpid Then
           Unload Liner(i)
        End If
    Next i
    
    ReList
    
End If

End Sub

Private Sub VScroll1_Change()
picBoard.Top = -VScroll1.Value
picBoard.Refresh
End Sub

Private Sub VScroll1_Scroll()
picBoard.Top = -VScroll1.Value
picBoard.Refresh

End Sub
