VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7035
   ClientLeft      =   3195
   ClientTop       =   960
   ClientWidth     =   5625
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0CCA
   ScaleHeight     =   7035
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   3500
      Left            =   240
      Top             =   1800
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Element-X Software  2001"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lander"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Image Asteroid7 
      Height          =   255
      Left            =   480
      Picture         =   "Form2.frx":81840
      Top             =   3000
      Width           =   285
   End
   Begin VB.Image Asteroid8 
      Height          =   255
      Left            =   2280
      Picture         =   "Form2.frx":81C7E
      Top             =   600
      Width           =   285
   End
   Begin VB.Image Asteroid9 
      Height          =   255
      Left            =   600
      Picture         =   "Form2.frx":820BC
      Top             =   5160
      Width           =   285
   End
   Begin VB.Image Asteroid14 
      Height          =   255
      Left            =   3120
      Picture         =   "Form2.frx":824FA
      Top             =   3240
      Width           =   285
   End
   Begin VB.Image Asteroid13 
      Height          =   255
      Left            =   960
      Picture         =   "Form2.frx":82938
      Top             =   1560
      Width           =   285
   End
   Begin VB.Image Asteroid12 
      Height          =   255
      Left            =   3000
      Picture         =   "Form2.frx":82D76
      Top             =   1920
      Width           =   285
   End
   Begin VB.Image Asteroid11 
      Height          =   255
      Left            =   720
      Picture         =   "Form2.frx":831B4
      Top             =   4080
      Width           =   285
   End
   Begin VB.Image Asteroid10 
      Height          =   255
      Left            =   4320
      Picture         =   "Form2.frx":835F2
      Top             =   4080
      Width           =   285
   End
   Begin VB.Image picLander 
      Height          =   375
      Left            =   1560
      Picture         =   "Form2.frx":83A30
      Top             =   2160
      Width           =   450
   End
   Begin VB.Image picPlanet 
      Height          =   555
      Left            =   0
      Picture         =   "Form2.frx":8436E
      Top             =   6480
      Width           =   5610
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Form1.Show
Unload Me
End Sub

Private Sub Form_Load()
'On Error GoTo err
Dim x
Open App.Path & "\HighScores.txt" For Input As #1
List1.Clear
Do While Not EOF(1)
Input #1, x
List1.AddItem (x)
Loop
Close #1

List1.ListIndex = 0
Form4.lblName1.Caption = List1.Text
List1.ListIndex = List1.ListIndex + 1
Form4.lblScore1.Caption = List1.Text
List1.ListIndex = List1.ListIndex + 1
Form4.lblName2.Caption = List1.Text
List1.ListIndex = List1.ListIndex + 1
Form4.lblScore2.Caption = List1.Text
List1.ListIndex = List1.ListIndex + 1
Form4.lblName3.Caption = List1.Text
List1.ListIndex = List1.ListIndex + 1
Form4.lblScore3.Caption = List1.Text
List1.ListIndex = List1.ListIndex + 1
Form4.lblName4.Caption = List1.Text
List1.ListIndex = List1.ListIndex + 1
Form4.lblScore4.Caption = List1.Text
List1.ListIndex = List1.ListIndex + 1
Form4.lblName5.Caption = List1.Text
Form4.lblScore5.Caption = List1.Text
Exit Sub

err: MsgBox "Couldn't Load High Scores"
Exit Sub
End Sub

Private Sub Timer1_Timer()
Form1.Show
Unload Me
End Sub
