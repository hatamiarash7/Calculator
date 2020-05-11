VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   5340
   ClientLeft      =   615
   ClientTop       =   615
   ClientWidth     =   2190
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   2190
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command6 
      Caption         =   "Command1"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command1"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command1"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command1"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   2160
      Y1              =   4320
      Y2              =   4320
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Const SND_ASYNC = &H1
Private Const SND_FILENAME = &H20000
Private Sub Form_Load()
Form3.Caption = "Music"

Command1.Caption = "Track    1"
Command2.Caption = "Track    2"
Command3.Caption = "Track    5"
Command4.Caption = "Track    4"
Command5.Caption = "Track    3"
Command6.Caption = "Stop Playing"

Command1.Font = arial
Command2.Font = arial
Command3.Font = arial
Command4.Font = arial
Command5.Font = arial
Command6.Font = arial

Command1.FontSize = 12
Command2.FontSize = 12
Command3.FontSize = 12
Command4.FontSize = 12
Command5.FontSize = 12
Command6.FontSize = 12
End Sub
Private Sub Command1_Click()         'Track 1
sndPlaySound App.Path & "\Sound\Main_1.wav", SND_ASYNC Or SND_FILENAME
End Sub
Private Sub Command2_Click()         'Track 2
sndPlaySound App.Path & "\Sound\Main_2.wav", SND_ASYNC Or SND_FILENAME
End Sub
Private Sub Command3_Click()         'Track 5
sndPlaySound App.Path & "\Sound\Main_3.wav", SND_ASYNC Or SND_FILENAME
End Sub
Private Sub Command4_Click()         'Track 4
sndPlaySound App.Path & "\Sound\Main_4.wav", SND_ASYNC Or SND_FILENAME
End Sub
Private Sub Command5_Click()         'Track 3
sndPlaySound App.Path & "\Sound\Main_5.wav", SND_ASYNC Or SND_FILENAME
End Sub
Private Sub Command6_Click()         'Stop Playing
sndPlaySound App.Path & "\Sound\End.wav", SND_ASYNC Or SND_FILENAME
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
Form3.Hide
sndPlaySound App.Path & "\Sound\End.wav", SND_ASYNC Or SND_FILENAME
End Sub
