VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form7"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5265
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   960
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   2184
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Private Sub Command1_Click()
a = InputBox("input a number Between 0 & 17", "Faktoriel", 0)
While a > 17
    d = MsgBox("Please input a number Between 0 & 17", vbOKOnly, "Error")
    a = InputBox("input a number Between 0 & 17", "Faktoriel", 0)
Wend
k = 1
For I = 2 To a
    k = k * I
    Print k
Next I
Command1.FontSize = 16
Command1.Caption = "Click to Clear"
Command1.Enabled = False
End Sub
Private Sub Form_Click()
Cls
Command1.Enabled = True
Command1.Font = arial
Command1.FontSize = 16
Command1.Caption = "Start"
End Sub
Private Sub Form_Load()
Form7.FontSize = 13
Form7.FontBold = True
Form7.Width = 5500
Form7.Height = 5500
Form7.AutoRedraw = True
Form7.Caption = "Factoriel"

Command1.Caption = "Start"
Command1.Font = arial
Command1.FontSize = 16
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

