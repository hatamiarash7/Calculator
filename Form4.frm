VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form4"
   ClientHeight    =   8550
   ClientLeft      =   5730
   ClientTop       =   2610
   ClientWidth     =   13455
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   13455
   Begin VB.CommandButton Command8 
      Caption         =   "Command6"
      Height          =   615
      Left            =   8760
      TabIndex        =   19
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command6"
      Height          =   615
      Left            =   8760
      TabIndex        =   18
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   855
      Left            =   8640
      TabIndex        =   12
      Top             =   6840
      Width           =   4695
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   1200
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   495
         Left            =   3480
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3375
      Left            =   8640
      TabIndex        =   11
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sinx"
      Height          =   615
      Left            =   8880
      TabIndex        =   10
      Top             =   360
      Width           =   1575
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7575
      Left            =   120
      ScaleHeight     =   7575
      ScaleWidth      =   8175
      TabIndex        =   9
      Top             =   120
      Width           =   8175
      Begin VB.Line L1 
         X1              =   0
         X2              =   5640
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line L2 
         X1              =   3240
         X2              =   3240
         Y1              =   480
         Y2              =   4560
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   960
         Shape           =   3  'Circle
         Top             =   4800
         Width           =   255
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sinx"
      Height          =   615
      Left            =   8880
      TabIndex        =   8
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sinx"
      Height          =   615
      Left            =   8880
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Sinx"
      Height          =   615
      Left            =   8880
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3255
      Left            =   8640
      TabIndex        =   1
      Top             =   3480
      Width           =   2055
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   1695
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Option4"
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   2280
         Width           =   1455
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   10200
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   7800
      Width           =   8295
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Const SND_ASYNC = &H1
Private Const SND_FILENAME = &H20000

Dim x0, y0, d As Single, co, fu As String, cc As Integer
Private Sub Form_Load()
Form4.Caption = "Trigonometry Drawing By Arash Hatami : hatamiarash7@gmail.com"

Frame1.Caption = "Functions"
Frame2.Caption = "Color"
Frame3.Caption = "Calculator"
Frame1.BackColor = RGB(223, 243, 249)
Frame2.BackColor = RGB(223, 243, 249)
Frame3.BackColor = RGB(223, 243, 249)

Option1.Value = True
Option1.Caption = "Black"
Option2.Caption = "Blue"
Option3.Caption = "Red"
Option4.Caption = "Green"
Option1.Font = arial
Option2.Font = arial
Option3.Font = arial
Option4.Font = arial
Option1.FontSize = 11
Option2.FontSize = 11
Option3.FontSize = 11
Option4.FontSize = 11
Option1.BackColor = RGB(223, 243, 249)
Option2.BackColor = RGB(223, 243, 249)
Option3.BackColor = RGB(223, 243, 249)
Option4.BackColor = RGB(223, 243, 249)

Command5.Caption = "Sin X"
Command6.Caption = "Cos X"
Command7.Caption = "Tan X"
Command8.Caption = "Cot X"

Command5.Font = arial
Command6.Font = arial
Command7.Font = arial
Command8.Font = arial

Command5.FontSize = 12
Command6.FontSize = 12
Command7.FontSize = 12
Command8.FontSize = 12

Label1.Caption = "Arash Hatami"
Label2.Caption = "Sin"
Label3.Caption = ""
Label1.Font = arial
Label2.Font = arial
Label3.Font = arial
Label1.FontSize = 18
Label2.FontSize = 18
Label3.FontSize = 15
Label1.BackColor = RGB(223, 243, 249)
Label2.BackColor = RGB(223, 243, 249)
Label3.BackColor = RGB(223, 243, 249)


Text2.Text = ""
Text2.Font = arial
Text2.FontSize = 25

Shape1.Left = -1000
Shape1.Top = -1000

Pic.ScaleMode = 3
x0 = Pic.ScaleWidth / 2
y0 = Pic.ScaleHeight / 2
L1.X1 = x0
L1.X2 = x0
L1.Y1 = 0
L1.Y2 = 2 * y0
'-------------------------------------------'
L2.X1 = 0:
L2.X2 = 2 * x0
L2.Y1 = y0
L2.Y2 = y0
End Sub
Private Sub Command5_Click()
Pic.Cls
xx = (Text2.Text * 3.14) / 180
z = -(Sin(xx)) * 2
Shape1.Top = (40 * z + y0) - 8
Shape1.Left = (40 * xx + x0) - 8
Shape1.BorderColor = co
Shape1.FillColor = co
For x = -10 * 3.14 To 10 * 3.14 Step 0.005
    y = -(Sin(x)) * 2
    Pic.PSet (40 * x + x0, 40 * y + y0), co
Next x
Label2.Caption = "Sin"
If Text2.Text = "" Then
    dd = MsgBox("Please Input A Number For Degree", , "Error")
Else
    fu = Round(Sin(Text2.Text * 3.14 / 180), 3)
    xx = (Text2.Text * 3.14) / 180
    Label3.Caption = " =" & fu
End If
End Sub
Private Sub Command6_Click()
Pic.Cls
xx = (Text2.Text * 3.14) / 180
z = -(Cos(xx)) * 2
Shape1.Top = (40 * z + y0) - 8
Shape1.Left = (40 * xx + x0) - 8
Shape1.BorderColor = co
Shape1.FillColor = co
For x = -10 * 3.14 To 10 * 3.14 Step 0.005
    y = -(Cos(x)) * 2
    Pic.PSet (40 * x + x0, 40 * y + y0), co
Next x
Label2.Caption = "Cos"
If Text2.Text = "" Then
    dd = MsgBox("Please Input A Number For Degree", , "Error")
Else
    fu = Round(Cos(Text2.Text * 3.14 / 180), 3)
    Label3.Caption = " =" & fu
End If
End Sub
Private Sub Command7_Click()
Pic.Cls
xx = (Text2.Text * 3.14) / 180
z = -(Tan(xx)) * 2
Shape1.Top = (40 * z + y0) - 8
Shape1.Left = (40 * xx + x0) - 8
Shape1.BorderColor = co
Shape1.FillColor = co
For x = -10 * 3.14 To 10 * 3.14 Step 0.005
    y = -(Tan(x)) * 2
    Pic.PSet (40 * x + x0, 40 * y + y0), co
Next x
Label2.Caption = "Tan"
If Text2.Text = "" Then
    dd = MsgBox("Please Input A Number For Degree", , "Error")
Else
    fu = Round(Tan(Text2.Text * 3.14 / 180), 3)
    Label3.Caption = " =" & fu
End If
If Val(Text2.Text) >= 73 Then
    mm = MsgBox("Your Dot Is Out Of Reach !!", , "Sorry")
    Text2.SetFocus
End If
End Sub
Private Sub Command8_Click()
Pic.Cls
xx = (Text2.Text * 3.14) / 180
z = -(1 / Tan(xx)) * 2
Shape1.Top = (40 * z + y0) - 8
Shape1.Left = (40 * xx + x0) - 8
Shape1.BorderColor = co
Shape1.FillColor = co
For x = -10 * 3.14 To -0.0005 * 3.14 Step 0.005
    y = -(1 / Tan(x)) * 2
    Pic.PSet (40 * x + x0, 40 * y + y0), co
Next x
For x = 0.0005 * 3.14 To 10 * 3.14 Step 0.005
    y = -(1 / Tan(x)) * 2
    Pic.PSet (40 * x + x0, 40 * y + y0), co
Next x
Label2.Caption = "Cot"
If Text2.Text = "" Then
    dd = MsgBox("Please Input A Number For Degree", , "Error")
Else
    fu = Round(1 / Tan(Text2.Text * 3.14 / 180), 3)
Label3.Caption = " =" & fu
End If
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
Private Sub Option1_Click()
Option2.Value = False
Option3.Value = False
Option4.Value = False
co = vbBlack
End Sub
Private Sub Option2_Click()
Option1.Value = False
Option3.Value = False
Option4.Value = False
co = vbBlue
End Sub
Private Sub Option3_Click()
Option2.Value = False
Option1.Value = False
Option4.Value = False
co = vbRed
End Sub
Private Sub Option4_Click()
Option2.Value = False
Option3.Value = False
Option1.Value = False
co = vbGreen
End Sub

