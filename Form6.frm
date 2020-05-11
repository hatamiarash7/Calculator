VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form6"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2190
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   2190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
x = InputBox("Input Number count", "Arash Hatami")
c = 0
For I = 1 To x
    a = InputBox("input numbers        " & I, "arash hatami")
    c = c + a
Next I
d = c / x
aa = MsgBox("average is :" & " " & d, vbInformation, "result")
End Sub
Private Sub Form_Load()
Form6.Caption = "Average"

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
