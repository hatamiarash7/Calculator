VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form5"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11580
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   4680
      TabIndex        =   5
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   975
      Left            =   7560
      TabIndex        =   4
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As Long, c As Long, f As Long, I As Long, d As Long
a = Val(Text1.Text)
b = Val(Text2.Text)
Do While a > 0
    d = a Mod b
    c = d * 10 ^ I
    f = f + c
    a = Int(a / b)
    I = I + 1
Loop
Label3.Caption = f
End Sub
Private Sub Form_Load()
Form1.Icon = LoadPicture(App.Path & "\picture\Icon.ico")
Form2.Icon = LoadPicture(App.Path & "\picture\Icon.ico")
Form3.Icon = LoadPicture(App.Path & "\picture\Icon.ico")
Form4.Icon = LoadPicture(App.Path & "\picture\Icon.ico")
Form5.Icon = LoadPicture(App.Path & "\picture\Icon.ico")
Form5.Caption = "Math Base By Arash Hatami : hatamiarash7@gmail.com"

Label1.Caption = "Input Your Number"
Label2.Caption = "Input Your Base"
Label3.Caption = ""
Label1.FontSize = 14
Label2.FontSize = 14
Label3.FontSize = 20
Label1.Font = arial
Label2.Font = arial
Label3.Font = arial
Label1.BackColor = RGB(223, 243, 249)
Label2.BackColor = RGB(223, 243, 249)
Label3.BackColor = RGB(223, 243, 249)

Text1.Text = ""
Text2.Text = ""
Text1.Font = arial
Text2.Font = arial
Text1.FontSize = 20
Text2.FontSize = 20

Command1.Caption = "Calculate"
Command1.Font = arial
Command1.FontSize = 18
End Sub
Private Sub Form_Unload(Cancel As Integer)
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
Form1.Show
Form2.Hide
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Hide
End Sub
