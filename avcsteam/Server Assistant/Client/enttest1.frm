VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "PutBack"
      Height          =   795
      Left            =   1980
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "enttest1.frx":0000
      Top             =   2100
      Width           =   8235
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Foxalicious"
      Height          =   795
      Left            =   420
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   1680
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   1200
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReportEnts Lib "redll.dll" Alias "reportents" (BSPPath As String, BankData As String) As Long
Private Declare Function ImportEnts Lib "redll.dll" Alias "importents" (BSPPath As String, BankData As String) As Long
Private Declare Function DllCanUnloadNow Lib "redll.dll" () As Integer
Private Const BankSize = 512000

Private Sub Command1_Click()

Dim Size As Long
Dim CurrBank As String
Dim String1 As String
String1 = App.Path + "\2fort.bsp"
Label1 = String1
CurrBank = Space(BankSize)

Size = ReportEnts(String1, CurrBank)
Label2 = Size

If Size > 0 Then
    CurrBank = Left(CurrBank, Size)
End If

Text1 = CurrBank


CurrBank = vbNullString
DllCanUnloadNow

End Sub

Private Sub Command2_Click()

Dim CurrBank As String
Dim String1 As String

String1 = App.Path + "\2fort.bsp"
CurrBank = Text1

Size = ImportEnts(String1, CurrBank)

Label1 = Size

CurrBank = vbNullString
DllCanUnloadNow


End Sub
