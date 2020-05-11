VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form9"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   1785
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   1785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a = InputBox("Please Input Your Number", "First Number")
c = 0
For I = 1 To a
    x = a Mod I
        If x = 0 Then c = c + 1
Next I
If c = 2 Then
    b = MsgBox("Your Number is First", vbInformation)
Else
    b = MsgBox("Your Number isn't First", vbInformation)
End If
End Sub
Private Sub Form_Load()
Form9.Caption = "F number"

Command1.Caption = "Start"
Command1.Font = arial
Command1.FontSize = 15
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
