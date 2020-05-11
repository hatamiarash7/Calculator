VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4110
   ClientLeft      =   3045
   ClientTop       =   2430
   ClientWidth     =   7125
   LinkTopic       =   "Form2"
   ScaleHeight     =   4110
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   650
      Left            =   2640
      Top             =   3360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   3975
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3135
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   3600
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   3600
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3600
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Website : www.arash-hatami.ir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Email : hatamiarash7@gmail.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Arash Hatami"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
Form2.Hide
End Sub
Private Sub Form_Load()
Form2.Height = 4000
Form2.Width = 6850
Form2.Left = 12165
Form2.Top = 6480

Label1.Caption = "Arash Hatami"
Label2.Caption = "Email : hatamiarash7@gmail.com"
Label3.Caption = "Website : www.arash-hatami.ir"
Label1.BackColor = RGB(223, 243, 249)
Label2.BackColor = RGB(223, 243, 249)
Label3.BackColor = RGB(223, 243, 249)
Label1.FontSize = 26
Label2.FontSize = 12
Label3.FontSize = 14
Label2.FontBold = True

Command1.Caption = "Back"
Command1.Font = arial
Command1.FontSize = 15

Image1.Picture = LoadPicture(App.Path & "\picture\Me2.JPG")

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
End
End Sub
Private Sub Timer1_Timer()
Label1.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub

