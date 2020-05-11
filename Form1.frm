VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4425
   ClientLeft      =   3360
   ClientTop       =   2850
   ClientWidth     =   9345
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   9345
   Begin VB.CommandButton Command33 
      Caption         =   "Command33"
      Height          =   615
      Left            =   7320
      TabIndex        =   50
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Command32"
      Height          =   615
      Left            =   7320
      TabIndex        =   49
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Command31"
      Height          =   615
      Left            =   7320
      TabIndex        =   48
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Command30"
      Height          =   615
      Left            =   7320
      TabIndex        =   47
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Command29"
      Height          =   615
      Left            =   7320
      TabIndex        =   46
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Command28"
      Height          =   615
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   120
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   650
      Left            =   5280
      Top             =   3840
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Equation"
      Height          =   375
      Left            =   3720
      TabIndex        =   43
      Top             =   3960
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Trigonometry"
      Height          =   375
      Left            =   3720
      TabIndex        =   42
      Top             =   3360
      Width           =   2415
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Advance"
      Height          =   375
      Left            =   2040
      TabIndex        =   41
      Top             =   3960
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Standard"
      Height          =   375
      Left            =   2040
      TabIndex        =   40
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Cot (-x)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   39
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Tan (-x)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   38
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Cos (-x)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   37
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Sin (-x)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   36
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Cot^2(x)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   35
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Tan^2(x)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   34
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Cos^2(x)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   33
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Sin^2(x)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   32
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Cot 2x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   31
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Tan 2x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   30
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Cos 2x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   29
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Sin 2x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   28
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "+ -"
      Height          =   495
      Left            =   1320
      TabIndex        =   27
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "sqr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   26
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "x^2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   25
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "�--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   23
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   21
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "cot"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   20
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "tan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   19
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "cos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   18
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "sin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   17
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton exit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   16
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton result 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   15
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      Caption         =   "�"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   14
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command13 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   13
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command12 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   11
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton c9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   9
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton c8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton c7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton c6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton c5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton c4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton c3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton c2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton c1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox t1 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   4335
   End
   Begin VB.Line Line9 
      X1              =   7200
      X2              =   3600
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line8 
      X1              =   3600
      X2              =   2640
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line7 
      X1              =   7200
      X2              =   7200
      Y1              =   120
      Y2              =   4320
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   1920
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line5 
      X1              =   1920
      X2              =   1920
      Y1              =   3120
      Y2              =   3960
   End
   Begin VB.Line Line4 
      X1              =   2640
      X2              =   120
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line3 
      X1              =   3600
      X2              =   3600
      Y1              =   840
      Y2              =   3240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Arash hatami"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Line Line2 
      X1              =   2640
      X2              =   2640
      Y1              =   840
      Y2              =   3240
   End
   Begin VB.Line Line1 
      X1              =   1920
      X2              =   1920
      Y1              =   840
      Y2              =   3240
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Arash hatami"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   375
      Left            =   5760
      TabIndex        =   44
      Top             =   3360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Const SND_ASYNC = &H1
Private Const SND_FILENAME = &H20000
Dim a As String, q As String, op As String, f3_close As Boolean
Private Sub Command28_Click()        'Drawing mode
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Show
Form5.Hide
End Sub
Private Sub Command29_Click()        'Base
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Hide
Form5.Show
End Sub
Private Sub Command30_Click()        'Average
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Show
Form7.Hide
Form8.Hide
Form9.Hide
End Sub
Private Sub Command31_Click()        'Factoriel
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form7.Show
Form8.Hide
Form9.Hide
End Sub
Private Sub Command32_Click()        'Bast
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Show
Form9.Hide
End Sub
Private Sub Command33_Click()        'First Number
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Show
End Sub
Private Sub Form_Load()
f3_close = True
Form1.Icon = LoadPicture(App.Path & "\picture\Icon2.ico")
Form2.Icon = LoadPicture(App.Path & "\picture\Icon2.ico")
Form3.Icon = LoadPicture(App.Path & "\picture\Icon2.ico")
Form4.Icon = LoadPicture(App.Path & "\picture\Icon2.ico")
Form5.Icon = LoadPicture(App.Path & "\picture\Icon2.ico")
Form1.Show
Form2.Hide
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Hide
Form1.Caption = "Calculator    -     Write , Design , Compiled By Arash Hatami - hatamiarash7@gmail.com"
Form1.BackColor = RGB(223, 243, 249)
Form2.BackColor = RGB(223, 243, 249)
Form3.BackColor = RGB(223, 243, 249)
Form4.BackColor = RGB(223, 243, 249)
Form5.BackColor = RGB(223, 243, 249)
Form6.BackColor = RGB(223, 243, 249)
Form7.BackColor = RGB(223, 243, 249)
Form8.BackColor = RGB(223, 243, 249)
Form9.BackColor = RGB(223, 243, 249)

Label1.Caption = "Arash Hatami"
Label1.Font = arial
Label1.FontSize = 12
Label1.BackColor = RGB(223, 243, 249)
Label1.ForeColor = vbRed
Label2.Caption = "Music"
Label2.Font = arial
Label2.FontSize = 12
Label2.BackColor = RGB(223, 243, 249)
Label2.ForeColor = vbRed

Option1.Font = arial
Option2.Font = arial
Option3.Font = arial
Option4.Font = arial
Option1.BackColor = RGB(223, 243, 249)
Option2.BackColor = RGB(223, 243, 249)
Option3.BackColor = RGB(223, 243, 249)
Option4.BackColor = RGB(223, 243, 249)
Option1.FontSize = 14
Option2.FontSize = 14
Option3.FontSize = 14
Option4.FontSize = 14
Option1.SetFocus

Line1.BorderColor = vbBlack
Line2.BorderColor = vbBlack
Line3.BorderColor = vbBlack
Line4.BorderColor = vbBlack
Line5.BorderColor = vbBlack
Line6.BorderColor = vbBlack
Line7.BorderColor = vbBlack

t1.BackColor = RGB(153, 217, 234)
t1.FontSize = 20
t1.Text = ""
t1.Alignment = 1
t1.SetFocus

Command1.Caption = "Sin"
Command2.Caption = "Cos"
Command3.Caption = "Tan"
Command4.Caption = "Cot"
Command16.Caption = "Sin 2x"
Command17.Caption = "Cos 2x"
Command18.Caption = "Tan 2x"
Command19.Caption = "Cot 2x"
Command20.Caption = "Sin^2 x"
Command21.Caption = "Cos^2 x"
Command22.Caption = "Tan^2 x"
Command23.Caption = "Cot^2 x"
Command24.Caption = "Sin (-x)"
Command25.Caption = "Cos (-x)"
Command26.Caption = "Tan (-x)"
Command27.Caption = "Cot (-x)"
Command28.Caption = "Drawing Mode"
Command29.Caption = "Base"
Command30.Caption = "Average"
Command31.Caption = "Factoriel"
Command32.Caption = "Bast"
Command33.Caption = "First number"

Command16.FontSize = 12
Command17.FontSize = 12
Command18.FontSize = 12
Command19.FontSize = 12
Command20.FontSize = 12
Command21.FontSize = 12
Command22.FontSize = 12
Command23.FontSize = 12
Command24.FontSize = 12
Command25.FontSize = 12
Command26.FontSize = 12
Command27.FontSize = 12
Command28.FontSize = 12
Command29.FontSize = 12
Command30.FontSize = 12
Command31.FontSize = 12
Command32.FontSize = 12
Command33.FontSize = 12

c1.FontBold = False
c2.FontBold = False
c3.FontBold = False
c4.FontBold = False
c5.FontBold = False
c6.FontBold = False
c7.FontBold = False
c8.FontBold = False
c9.FontBold = False
result.FontBold = False
Command1.FontBold = False
Command2.FontBold = False
Command3.FontBold = False
Command4.FontBold = False
Command5.FontBold = False
Command6.FontBold = False
Command7.FontBold = False
Command8.FontBold = False
Command9.FontBold = False
Command10.FontBold = False
Command11.FontBold = False
Command12.FontBold = False
Command13.FontBold = False
Command14.FontBold = False
Command15.FontBold = False
Command16.FontBold = False
Command17.FontBold = False
Command18.FontBold = False
Command19.FontBold = False
Command20.FontBold = False
Command21.FontBold = False
Command22.FontBold = False
Command23.FontBold = False
Command24.FontBold = False
Command25.FontBold = False
Command26.FontBold = False
Command27.FontBold = False
Command28.FontBold = False
Command29.FontBold = False
Command30.FontBold = False
Command31.FontBold = False
Command32.FontBold = False
Command33.FontBold = False
End Sub
Private Sub c1_Click()               '1 num
If t1.Text = "0" Or a = "z" Then
    t1.Text = "1"
    a = ""
Else
    t1.Text = t1.Text + "1"
End If
t1.SetFocus
End Sub
Private Sub c2_Click()               '2 num
If t1.Text = "0" Or a = "z" Then
    t1.Text = "2"
    a = ""
Else
    t1.Text = t1.Text + "2"
End If
t1.SetFocus
End Sub
Private Sub c3_Click()               '3 num
If t1.Text = "0" Or a = "z" Then
    t1.Text = "3"
    a = ""
Else
    t1.Text = t1.Text + "3"
End If
t1.SetFocus
End Sub
Private Sub c4_Click()               '4 num
If t1.Text = "0" Or a = "z" Then
    t1.Text = "4"
    a = ""
Else
    t1.Text = t1.Text + "4"
End If
t1.SetFocus
End Sub
Private Sub c5_Click()               '5 num
If t1.Text = "0" Or a = "z" Then
    t1.Text = "5"
    a = ""
Else
    t1.Text = t1.Text + "5"
End If
t1.SetFocus
End Sub
Private Sub c6_Click()               '6 num
If t1.Text = "0" Or a = "z" Then
    t1.Text = "6"
    a = ""
Else
    t1.Text = t1.Text + "6"
End If
t1.SetFocus
End Sub
Private Sub c7_Click()               '7 num
If t1.Text = "0" Or a = "z" Then
    t1.Text = "7"
    a = ""
Else
    t1.Text = t1.Text + "7"
End If
t1.SetFocus
End Sub
Private Sub c8_Click()               '8 num
If t1.Text = "0" Or a = "z" Then
    t1.Text = "8"
    a = ""
Else
    t1.Text = t1.Text + "8"
End If
t1.SetFocus
End Sub
Private Sub c9_Click()               '9 num
If t1.Text = "0" Or a = "z" Then
    t1.Text = "9"
    a = ""
Else
    t1.Text = t1.Text + "9"
End If
t1.SetFocus
End Sub
Private Sub Command1_Click()               'sin x
t = Val(t1.Text)
t1 = Round(Sin(t * 3.14 / 180), 3)
q = t1.Text
op = "sin"
a = "z"
t1.SetFocus
End Sub
Private Sub Command2_Click()               'cos x
t = Val(t1.Text)
t1 = Round(Cos(t * 3.14 / 180), 3)
t1.SetFocus
End Sub
Private Sub Command3_Click()               'tan x
t = Val(t1.Text)
t1 = Round(Tan(t * 3.14 / 180), 3)
t1.SetFocus
End Sub
Private Sub Command4_Click()               'cot x
t = Val(t1.Text)
t1 = Round(Cos(t * 3.14 / 180) / Sin(t * 3.14 / 180), 3)
t1.SetFocus
End Sub
Private Sub Command5_Click()               'clear
t1 = ""
t1.Alignment = 1
t1.SetFocus
End Sub
Private Sub Command6_Click()               ' . key
q = Len(t1)
For I = 1 To q
   If Mid$(t1, q, 1) = "." Then
      k = "qq"
   End If
Next I
If k <> "qq" Then
   t1.Text = t1.Text + "."
End If
t1.SetFocus
End Sub
Private Sub Command7_Click()               'back key
q = Len(t1)
If q <> 0 Then
   e = Len(t1)
   t1 = Mid(t1, 1, e - 1)
End If
t1.SetFocus
End Sub
Private Sub Command8_Click()               'x^2
t = Val(t1.Text)
t1.Text = (t) ^ 2
t1.SetFocus
End Sub
Private Sub Command9_Click()               'sqr
t = Val(t1.Text)
t1.Text = Round(Sqr(t), 3)
t1.SetFocus
End Sub
Private Sub command10_Click()               '0 num
If t1.Text = "0" Or a = "z" Then
    t1.Text = "0"
    a = ""
Else
    t1.Text = t1.Text + "0"
End If
t1.SetFocus
End Sub
Private Sub Command11_Click()               ' + key
If op = "+" Then
    t1.Text = Val(q) + Val(t1)
End If
q = t1.Text
op = "+"
a = "z"
t1.SetFocus
End Sub
Private Sub Command12_Click()               ' - key
If op = "-" Then
    t1.Text = Val(q) - Val(t1)
End If
q = t1.Text
op = "-"
a = "z"
t1.SetFocus
End Sub
Private Sub Command13_Click()               ' x key
If op = "*" Then
    t1.Text = Val(q) * Val(t1)
End If
q = t1.Text
op = "*"
a = "z"
t1.SetFocus
End Sub
Private Sub Command14_Click()               ' � key
If op = "/" Then
    t1.Text = Val(q) / Val(t1)
End If
q = t1.Text
op = "/"
a = "z"
t1.SetFocus
End Sub
Private Sub Command15_Click()               ' +- key
t = Val(t1.Text)
t1.Text = (t) * (-1)
t1.SetFocus
End Sub
Private Sub Command16_Click()
t = Val(t1.Text)                                  'sin 2x
t1.Text = Round(2 * (Round(Sin(t * 3.14 / 180), 3)) * (Round(Cos(t), 3)), 3)
t1.SetFocus
End Sub
Private Sub Command17_Click()
t = Val(t1.Text)                                  'cos 2x
t1.Text = Round(1 - (2 * (Round((Sin(t * 3.14 / 180) ^ 2), 3))), 3)
t1.SetFocus
End Sub
Private Sub Command18_Click()
t = Val(t1.Text)                                  'tan 2x
t1.Text = Round((2 * ((Round(Sin(t * 3.14 / 180), 3)) / (Round(Cos(t * 3.14 / 180), 3)))) / (1 - (((Round(Sin(t * 3.14 / 180), 3)) / (Round(Cos(t * 3.14 / 180), 3))) ^ 2)), 3)
t1.SetFocus
End Sub
Private Sub Command19_Click()
t = Val(t1.Text)                                  'cot 2x
t1.Text = Round(2 * (Round((Round(Cos(t * 3.14 / 180), 3)) / (Round(Sin(t * 3.14 / 180), 3)), 3)) / (Round((Round(Cos(t * 3.14 / 180), 3)) / (Round(Sin(t * 3.14 / 180), 3)), 3) - 1), 3)
t1.SetFocus
End Sub
Private Sub Command20_Click()
t = Val(t1.Text)                                'sin^2 x
t1.Text = Round(1 - (Round(1 - (2 * (Sin(t * 3.14 / 180) ^ 2)), 3)) / (2), 3)
t1.SetFocus
End Sub
Private Sub Command21_Click()
t = Val(t1.Text)                                'cos^2 x
t1.Text = Round(1 + (Round(1 - (2 * (Sin(t * 3.14 / 180) ^ 2)), 3)) / (2), 3)
t1.SetFocus
End Sub
Private Sub Command22_Click()
t = Val(t1.Text)                                'tan^2 x
t1.Text = Round((Round(1 - (Round(1 - (2 * (Round((Sin(t * 3.14 / 180) ^ 2), 3))), 3)), 3)) / (Round(1 + (Round(1 - (2 * (Round((Sin(t * 3.14 / 180) ^ 2), 3))), 3)), 3)), 3)
t1.SetFocus
End Sub
Private Sub Command23_Click()
t = Val(t1.Text)                                'cot^2 x
t1.Text = (Round(1 + (Round(1 - (2 * (Round((Sin(t * 3.14 / 180) ^ 2), 3))), 3)), 3)) / (Round(1 - (Round(1 - (2 * (Round((Sin(t * 3.14 / 180) ^ 2), 3))), 3)), 3))
t1.SetFocus
End Sub
Private Sub Command24_Click()
t = Val(t1.Text)                                  'sin(-x)
t1.Text = Round(Sin(t * 3.14 / 180) * (-1), 3)
t1.SetFocus
End Sub
Private Sub Command25_Click()
t = Val(t1.Text)                                  'cos (-x)
t1.Text = Round(Cos(t * 3.14 / 180), 3)
t1.SetFocus
End Sub
Private Sub Command26_Click()
t = Val(t1.Text)                                  'tan (-x)
t1.Text = Round(Tan(t * 3.14 / 180) * (-1), 3)
t1.SetFocus
End Sub
Private Sub Command27_Click()
t = Val(t1.Text)                                  'cot (-x)
t1.Text = Round((Cos(t * 3.14 / 180) / Sin(t * 3.14 / 180)) * (-1), 3)
t1.SetFocus
End Sub
Private Sub Label1_Click()               'information
Form1.Hide
Form2.Show
End Sub
Private Sub Label2_Click()               'Music Pannel
If f3_close = False Then
   Form3.Hide
   f3_close = True
   Label2.ForeColor = &H8000000A
Else
   Form3.Show
   f3_close = False
   Label2.ForeColor = vbRed
End If
End Sub
Private Sub Option1_Click()               'standard
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command8.Visible = False
Command9.Visible = False
Command15.Visible = False
Command16.Visible = False
Command17.Visible = False
Command18.Visible = False
Command19.Visible = False
Command20.Visible = False
Command21.Visible = False
Command22.Visible = False
Command23.Visible = False
Command24.Visible = False
Command25.Visible = False
Command26.Visible = False
Command27.Visible = False
Line2.Visible = True
Line3.Visible = False
Line4.Visible = True
Line5.Visible = False
Line6.Visible = False
Line7.Visible = False
Line8.Visible = False
Line9.Visible = False
t1.SetFocus
End Sub
Private Sub Option2_Click()               'advance
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command8.Visible = True
Command9.Visible = True
Command15.Visible = True
Command16.Visible = False
Command17.Visible = False
Command18.Visible = False
Command19.Visible = False
Command20.Visible = False
Command21.Visible = False
Command22.Visible = False
Command23.Visible = False
Command24.Visible = False
Command25.Visible = False
Command26.Visible = False
Command27.Visible = False
Line2.Visible = True
Line3.Visible = False
Line4.Visible = True
Line5.Visible = True
Line6.Visible = True
Line7.Visible = False
Line8.Visible = False
Line9.Visible = False
t1.SetFocus
End Sub
Private Sub Option3_Click()               'S - C - T
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command8.Visible = False
Command9.Visible = False
Command15.Visible = False
Command16.Visible = False
Command17.Visible = False
Command18.Visible = False
Command19.Visible = False
Command20.Visible = False
Command21.Visible = False
Command22.Visible = False
Command23.Visible = False
Command24.Visible = False
Command25.Visible = False
Command26.Visible = False
Command27.Visible = False
Line2.Visible = True
Line3.Visible = True
Line4.Visible = True
Line5.Visible = False
Line6.Visible = False
Line7.Visible = False
Line8.Visible = True
Line9.Visible = False
t1.SetFocus
End Sub
Private Sub Option4_Click()               'equation
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command8.Visible = True
Command9.Visible = True
Command15.Visible = True
Command16.Visible = True
Command17.Visible = True
Command18.Visible = True
Command19.Visible = True
Command20.Visible = True
Command21.Visible = True
Command22.Visible = True
Command23.Visible = True
Command24.Visible = True
Command25.Visible = True
Command26.Visible = True
Command27.Visible = True
Line2.Visible = True
Line3.Visible = True
Line4.Visible = True
Line5.Visible = True
Line6.Visible = True
Line7.Visible = True
Line8.Visible = True
Line9.Visible = True
t1.SetFocus
End Sub
Private Sub result_Click()               ' = key
Select Case op
Case "+": t1.Text = Val(q) + Val(t1)
Case "-": t1.Text = Val(q) - Val(t1)
Case "*": t1.Text = Val(q) * Val(t1)
Case "/": t1.Text = Val(q) / Val(t1)
End Select
a = "z"
q = t1
End Sub
Private Sub exit_Click()               ' exit key / effect
l = MsgBox("Do You Want To Exit ?", vbYesNo + vbInformation, "End")
If l = 6 Then
   sndPlaySound App.Path & "\Sound\End.wav", SND_ASYNC Or SND_FILENAME
   
   Dim counter As Integer
   Dim I As Integer
   counter = Me.Height
     Do: DoEvents
       counter = counter - 10
       Me.Height = counter
       Me.Top = (Screen.Height - Me.Height) / 2
     Loop Until counter <= 10
   I = 15
   counter = Me.Width
     Do: DoEvents
       counter = counter + I
       Me.Width = counter
       Me.Left = (Screen.Width - Me.Width) / 2
       I = I + 1
     Loop Until counter >= Screen.Width
End
Else
    k = MsgBox("thank you baby !!", , "oh yeah ...")
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)               'exit form / effect
l = MsgBox("Do You Want To Exit ?", vbYesNo + vbInformation, "End")
If l = 6 Then
   sndPlaySound App.Path & "\Sound\End.wav", SND_ASYNC Or SND_FILENAME
   
   Dim counter As Integer
   Dim I As Integer
   counter = Me.Height
     Do: DoEvents
       counter = counter - 10
       Me.Height = counter
       Me.Top = (Screen.Height - Me.Height) / 2
     Loop Until counter <= 10
   I = 15
   counter = Me.Width
     Do: DoEvents
       counter = counter + I
       Me.Width = counter
       Me.Left = (Screen.Width - Me.Width) / 2
       I = I + 1
     Loop Until counter >= Screen.Width
End
Else
    k = MsgBox("thank you baby !!", , "oh yeah ...")
End If
End Sub
Private Sub Timer1_Timer()        'label color changing
Label1.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub
