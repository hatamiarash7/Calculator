VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form8"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(100) As Integer, b(100) As Integer, c(100) As Integer
Private Sub Command1_Click()
Cls
n = InputBox("Enter a number for n between 2  &  12 :", "Arash Hatami", 2)
a(1) = 1
b(1) = n
For I = 2 To n
    b(I) = n - I + 1
    a(I) = a(I - 1) * b(I - 1) / (I - 1)
    c(I) = n - b(I)
Next I
Print "a" & "^" & n
For I = 2 To n
    Print a(I) & "  a" & "^" & b(I) & "b" & "^" & c(I)
Next I
Print "b" & "^" & n
End Sub
Private Sub Form_Load()
Form8.Caption = "Bast"
Form8.Font = arial
Form8.FontSize = 15
Form8.AutoRedraw = True

Command1.Caption = "Start"
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

