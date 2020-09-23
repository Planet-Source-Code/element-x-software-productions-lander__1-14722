VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "High Scores"
   ClientHeight    =   3195
   ClientLeft      =   3600
   ClientTop       =   3075
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblScore5 
      BackStyle       =   0  'Transparent
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
      Left            =   2640
      TabIndex        =   17
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblScore4 
      BackStyle       =   0  'Transparent
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
      Left            =   2640
      TabIndex        =   16
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblScore3 
      BackStyle       =   0  'Transparent
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
      Left            =   2640
      TabIndex        =   15
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblScore2 
      BackStyle       =   0  'Transparent
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
      Left            =   2640
      TabIndex        =   14
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblScore1 
      BackStyle       =   0  'Transparent
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
      Left            =   2640
      TabIndex        =   13
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   960
      TabIndex        =   11
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblName5 
      BackStyle       =   0  'Transparent
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
      Left            =   720
      TabIndex        =   10
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblName4 
      BackStyle       =   0  'Transparent
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
      Left            =   720
      TabIndex        =   9
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblName3 
      BackStyle       =   0  'Transparent
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
      Left            =   720
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblName2 
      BackStyle       =   0  'Transparent
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
      Left            =   720
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblName1 
      BackStyle       =   0  'Transparent
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
      Left            =   720
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "5."
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
      Left            =   480
      TabIndex        =   5
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "4."
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
      Left            =   480
      TabIndex        =   4
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "3."
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
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "2."
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
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1."
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
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "High Scores:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Form_Unload(Cancel As Integer)
List1.AddItem lblName1.Caption
List1.ListIndex = List1.ListIndex + 1
List1.AddItem lblScore1.Caption
List1.ListIndex = List1.ListIndex + 1
List1.AddItem lblName2.Caption
List1.ListIndex = List1.ListIndex + 1
List1.AddItem lblScore2.Caption
List1.ListIndex = List1.ListIndex + 1
List1.AddItem lblName3.Caption
List1.ListIndex = List1.ListIndex + 1
List1.AddItem lblScore3.Caption
List1.ListIndex = List1.ListIndex + 1
List1.AddItem lblName4.Caption
List1.ListIndex = List1.ListIndex + 1
List1.AddItem lblScore4.Caption
List1.ListIndex = List1.ListIndex + 1
List1.AddItem lblName5.Caption
List1.ListIndex = List1.ListIndex + 1
List1.AddItem lblScore5.Caption

Open App.Path & "\HighScores.txt" For Output As #1
x = 0
List1.ListIndex = 0
While x < List1.ListCount - 1
Write #1, List1.Text
List1.ListIndex = List1.ListIndex + 1
x = x + 1
Wend
Close #1

End Sub
