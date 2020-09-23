VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lander"
   ClientHeight    =   7035
   ClientLeft      =   3120
   ClientTop       =   1095
   ClientWidth     =   5610
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0CCA
   ScaleHeight     =   7035
   ScaleWidth      =   5610
   Begin VB.CommandButton Command2 
      Caption         =   "Continue"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   240
      Top             =   2640
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   240
      Top             =   2160
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   1680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   1200
   End
   Begin VB.Label lblFuel 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Fuel:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblSpeed 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Thrust:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Image Asteroid10 
      Height          =   255
      Left            =   120
      Picture         =   "Form1.frx":81840
      Top             =   4560
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Asteroid11 
      Height          =   255
      Left            =   120
      Picture         =   "Form1.frx":81C7E
      Top             =   4920
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Asteroid12 
      Height          =   255
      Left            =   120
      Picture         =   "Form1.frx":820BC
      Top             =   5280
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Asteroid13 
      Height          =   255
      Left            =   120
      Picture         =   "Form1.frx":824FA
      Top             =   5640
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Asteroid14 
      Height          =   255
      Left            =   120
      Picture         =   "Form1.frx":82938
      Top             =   6000
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Asteroid9 
      Height          =   255
      Left            =   120
      Picture         =   "Form1.frx":82D76
      Top             =   4200
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Asteroid8 
      Height          =   255
      Left            =   120
      Picture         =   "Form1.frx":831B4
      Top             =   3840
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Asteroid7 
      Height          =   255
      Left            =   120
      Picture         =   "Form1.frx":835F2
      Top             =   3480
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Asteroid6 
      Height          =   255
      Left            =   120
      Picture         =   "Form1.frx":83A30
      Top             =   3120
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Asteroid5 
      Height          =   255
      Left            =   120
      Picture         =   "Form1.frx":83E6E
      Top             =   2760
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Asteroid4 
      Height          =   255
      Left            =   120
      Picture         =   "Form1.frx":842AC
      Top             =   2400
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Asteroid3 
      Height          =   255
      Left            =   120
      Picture         =   "Form1.frx":846EA
      Top             =   2040
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Asteroid2 
      Height          =   255
      Left            =   120
      Picture         =   "Form1.frx":84B28
      Top             =   1680
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Asteroid1 
      Height          =   255
      Left            =   120
      Picture         =   "Form1.frx":84F66
      Top             =   1320
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Asteroid 
      Height          =   255
      Left            =   120
      Picture         =   "Form1.frx":853A4
      Top             =   960
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Lander 
      Height          =   375
      Left            =   720
      Picture         =   "Form1.frx":857E2
      Top             =   1680
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image picSmash 
      Height          =   375
      Left            =   720
      Picture         =   "Form1.frx":86120
      Top             =   2040
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image picLander 
      Height          =   375
      Left            =   3960
      Picture         =   "Form1.frx":86A5E
      Top             =   840
      Width           =   450
   End
   Begin VB.Image picPlanet 
      Height          =   555
      Left            =   0
      Picture         =   "Form1.frx":8739C
      Top             =   6480
      Width           =   5610
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
NewGame
Command1.Visible = False
Timer1.Enabled = True
picLander.Picture = Lander.Picture
lblScore.Caption = 0
End Sub

Private Sub Command2_Click()
NewGame
Timer1.Enabled = True
Command2.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If lblFuel.Caption <= 0 Then GoTo less

If KeyCode = vbKeyDown And Timer1.Interval >= 100 And Timer1.Interval <= 200 Then
Timer1.Interval = Timer1.Interval + 4
lblFuel.Caption = lblFuel.Caption - 5
End If

If KeyCode <> vbKeyDown And Timer1.Interval > 100 Then Timer1.Interval = Timer1.Interval - 3

If KeyCode = vbKeyLeft And picLander.Left > 0 Then
picLander.Left = picLander.Left - 25
lblFuel.Caption = lblFuel.Caption - 2
End If

If KeyCode = vbKeyRight And picLander.Left + picLander.Width < Form1.Width Then
picLander.Left = picLander.Left + 25
lblFuel.Caption = lblFuel.Caption - 2
End If

Exit Sub

less: lblFuel.Caption = 0
If Timer1.Interval > 0 Then Timer1.Interval = Timer1.Interval - 4
End Sub

Private Sub Form_Load()
Dim LndLeft As Integer

Randomize Timer
LndLeft = Int(Rnd * 5040) + 1

picLander.Top = 840
picLander.Left = LndLeft

lblFuel.Caption = 300

Counter = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show
Unload Me
End Sub

Private Sub Timer1_Timer()
If Timer1.Interval < 100 Then Timer1.Interval = 100
If Timer1.Interval > 200 Then Timer1.Interval = 200

lblSpeed.Caption = Timer1.Interval

If picLander.Top + picLander.Height >= picPlanet.Top And Timer1.Interval >= 150 Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Command2.Visible = True
If Timer1.Interval > 149 And Timer1.Interval < 163 Then lblScore.Caption = lblScore.Caption + 50
If Timer1.Interval > 163 Then lblScore.Caption = lblScore.Caption + 20
Exit Sub
End If

If picLander.Top + picLander.Height >= picPlanet.Top And Timer1.Interval < 150 Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
picLander.Picture = picSmash.Picture
Command1.Visible = True
HighScores
Exit Sub
Else
picLander.Top = picLander.Top + 30
End If

Form_KeyDown 0, 0
End Sub

Private Sub Timer2_Timer()
If Asteroid.Left > Form1.Width Then Asteroid.Left = 0
If Asteroid1.Left > Form1.Width Then Asteroid1.Left = 0
If Asteroid2.Left > Form1.Width Then Asteroid2.Left = 0
If Asteroid3.Left > Form1.Width Then Asteroid3.Left = 0
If Asteroid4.Left > Form1.Width Then Asteroid4.Left = 0
If Asteroid5.Left > Form1.Width Then Asteroid5.Left = 0
If Asteroid6.Left > Form1.Width Then Asteroid6.Left = 0
If Asteroid7.Left > Form1.Width Then Asteroid7.Left = 0
If Asteroid8.Left > Form1.Width Then Asteroid8.Left = 0
If Asteroid9.Left > Form1.Width Then Asteroid9.Left = 0
If Asteroid10.Left > Form1.Width Then Asteroid10.Left = 0
If Asteroid11.Left > Form1.Width Then Asteroid11.Left = 0
If Asteroid12.Left > Form1.Width Then Asteroid12.Left = 0
If Asteroid13.Left > Form1.Width Then Asteroid13.Left = 0
If Asteroid14.Left > Form1.Width Then Asteroid14.Left = 0

Asteroid.Left = Asteroid.Left + 22
Asteroid1.Left = Asteroid1.Left + 22
Asteroid2.Left = Asteroid2.Left + 22
Asteroid3.Left = Asteroid3.Left + 22
Asteroid4.Left = Asteroid4.Left + 22
Asteroid5.Left = Asteroid5.Left + 22
Asteroid6.Left = Asteroid6.Left + 22
Asteroid7.Left = Asteroid7.Left + 22
Asteroid8.Left = Asteroid8.Left + 22
Asteroid9.Left = Asteroid9.Left + 22
Asteroid10.Left = Asteroid10.Left + 22
Asteroid11.Left = Asteroid11.Left + 22
Asteroid12.Left = Asteroid12.Left + 22
Asteroid13.Left = Asteroid13.Left + 22
Asteroid14.Left = Asteroid14.Left + 22
End Sub

Public Sub NewGame()
Dim LndLeft
Dim A
Randomize Timer
LndLeft = Int(Rnd * 5040) + 1

picLander.Top = 840
picLander.Left = LndLeft

lblFuel.Caption = 300

Asteroid.Visible = False
Asteroid1.Visible = False
Asteroid2.Visible = False
Asteroid3.Visible = False
Asteroid4.Visible = False
Asteroid5.Visible = False
Asteroid6.Visible = False
Asteroid7.Visible = False
Asteroid8.Visible = False
Asteroid9.Visible = False
Asteroid10.Visible = False
Asteroid11.Visible = False
Asteroid12.Visible = False
Asteroid13.Visible = False
Asteroid14.Visible = False

A = Int(Rnd * 15) + 1

Select Case A

Case Is = 1
Asteroid.Visible = True

Case Is = 2
Asteroid.Visible = True
Asteroid1.Visible = True

Case Is = 3
Asteroid.Visible = True
Asteroid1.Visible = True
Asteroid2.Visible = True

Case Is = 4
Asteroid.Visible = True
Asteroid1.Visible = True
Asteroid2.Visible = True
Asteroid3.Visible = True

Case Is = 5
Asteroid.Visible = True
Asteroid1.Visible = True
Asteroid2.Visible = True
Asteroid3.Visible = True
Asteroid4.Visible = True

Case Is = 6
Asteroid.Visible = True
Asteroid1.Visible = True
Asteroid2.Visible = True
Asteroid3.Visible = True
Asteroid4.Visible = True
Asteroid5.Visible = True

Case Is = 7
Asteroid.Visible = True
Asteroid1.Visible = True
Asteroid2.Visible = True
Asteroid3.Visible = True
Asteroid4.Visible = True
Asteroid5.Visible = True
Asteroid6.Visible = True

Case Is = 8
Asteroid.Visible = True
Asteroid1.Visible = True
Asteroid2.Visible = True
Asteroid3.Visible = True
Asteroid4.Visible = True
Asteroid5.Visible = True
Asteroid6.Visible = True
Asteroid7.Visible = True

Case Is = 9
Asteroid.Visible = True
Asteroid1.Visible = True
Asteroid2.Visible = True
Asteroid3.Visible = True
Asteroid4.Visible = True
Asteroid5.Visible = True
Asteroid6.Visible = True
Asteroid7.Visible = True
Asteroid8.Visible = True

Case Is = 10
Asteroid.Visible = True
Asteroid1.Visible = True
Asteroid2.Visible = True
Asteroid3.Visible = True
Asteroid4.Visible = True
Asteroid5.Visible = True
Asteroid6.Visible = True
Asteroid7.Visible = True
Asteroid8.Visible = True
Asteroid9.Visible = True

Case Is = 11
Asteroid.Visible = True
Asteroid1.Visible = True
Asteroid2.Visible = True
Asteroid3.Visible = True
Asteroid4.Visible = True
Asteroid5.Visible = True
Asteroid6.Visible = True
Asteroid7.Visible = True
Asteroid8.Visible = True
Asteroid9.Visible = True
Asteroid10.Visible = True

Case Is = 12
Asteroid.Visible = True
Asteroid1.Visible = True
Asteroid2.Visible = True
Asteroid3.Visible = True
Asteroid4.Visible = True
Asteroid5.Visible = True
Asteroid6.Visible = True
Asteroid7.Visible = True
Asteroid8.Visible = True
Asteroid9.Visible = True
Asteroid10.Visible = True
Asteroid11.Visible = True

Case Is = 13
Asteroid.Visible = True
Asteroid1.Visible = True
Asteroid2.Visible = True
Asteroid3.Visible = True
Asteroid4.Visible = True
Asteroid5.Visible = True
Asteroid6.Visible = True
Asteroid7.Visible = True
Asteroid8.Visible = True
Asteroid9.Visible = True
Asteroid10.Visible = True
Asteroid11.Visible = True
Asteroid12.Visible = True

Case Is = 14
Asteroid.Visible = True
Asteroid1.Visible = True
Asteroid2.Visible = True
Asteroid3.Visible = True
Asteroid4.Visible = True
Asteroid5.Visible = True
Asteroid6.Visible = True
Asteroid7.Visible = True
Asteroid8.Visible = True
Asteroid9.Visible = True
Asteroid10.Visible = True
Asteroid11.Visible = True
Asteroid12.Visible = True
Asteroid13.Visible = True

Case Is = 15
Asteroid.Visible = True
Asteroid1.Visible = True
Asteroid2.Visible = True
Asteroid3.Visible = True
Asteroid4.Visible = True
Asteroid5.Visible = True
Asteroid6.Visible = True
Asteroid7.Visible = True
Asteroid8.Visible = True
Asteroid9.Visible = True
Asteroid10.Visible = True
Asteroid11.Visible = True
Asteroid12.Visible = True
Asteroid13.Visible = True
Asteroid14.Visible = True

End Select

Asteroid.Top = Int(Rnd * 6280) + 1
Asteroid1.Top = Int(Rnd * 6280) + 1
Asteroid2.Top = Int(Rnd * 6280) + 1
Asteroid3.Top = Int(Rnd * 6280) + 1
Asteroid4.Top = Int(Rnd * 6280) + 1
Asteroid5.Top = Int(Rnd * 6280) + 1
Asteroid6.Top = Int(Rnd * 6280) + 1
Asteroid7.Top = Int(Rnd * 6280) + 1
Asteroid8.Top = Int(Rnd * 6280) + 1
Asteroid9.Top = Int(Rnd * 6280) + 1
Asteroid10.Top = Int(Rnd * 6280) + 1
Asteroid11.Top = Int(Rnd * 6280) + 1
Asteroid12.Top = Int(Rnd * 6280) + 1
Asteroid13.Top = Int(Rnd * 6280) + 1
Asteroid14.Top = Int(Rnd * 6280) + 1

Asteroid.Left = Int(Rnd * Form1.Width) + 1
Asteroid1.Left = Int(Rnd * Form1.Width) + 1
Asteroid2.Left = Int(Rnd * Form1.Width) + 1
Asteroid3.Left = Int(Rnd * Form1.Width) + 1
Asteroid4.Left = Int(Rnd * Form1.Width) + 1
Asteroid5.Left = Int(Rnd * Form1.Width) + 1
Asteroid6.Left = Int(Rnd * Form1.Width) + 1
Asteroid7.Left = Int(Rnd * Form1.Width) + 1
Asteroid8.Left = Int(Rnd * Form1.Width) + 1
Asteroid9.Left = Int(Rnd * Form1.Width) + 1
Asteroid10.Left = Int(Rnd * Form1.Width) + 1
Asteroid11.Left = Int(Rnd * Form1.Width) + 1
Asteroid12.Left = Int(Rnd * Form1.Width) + 1
Asteroid13.Left = Int(Rnd * Form1.Width) + 1
Asteroid14.Left = Int(Rnd * Form1.Width) + 1

Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
End Sub


Private Sub Timer3_Timer()
Counter = Counter + 1

If Asteroid.Top + Asteroid.Height >= picLander.Top And Asteroid.Top <= picLander.Top + picLander.Height And Asteroid.Left >= picLander.Left And Asteroid.Left + Asteroid.Width <= picLander.Left + picLander.Width And Asteroid.Visible = True Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
HighScores
picLander.Picture = picSmash.Picture
Command1.Visible = True
End If

If Asteroid1.Top + Asteroid1.Height >= picLander.Top And Asteroid1.Top <= picLander.Top + picLander.Height And Asteroid1.Left >= picLander.Left And Asteroid1.Left + Asteroid1.Width <= picLander.Left + picLander.Width And Asteroid1.Visible = True Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
HighScores
picLander.Picture = picSmash.Picture
Command1.Visible = True
End If

If Asteroid2.Top + Asteroid2.Height >= picLander.Top And Asteroid2.Top <= picLander.Top + picLander.Height And Asteroid2.Left >= picLander.Left And Asteroid2.Left + Asteroid2.Width <= picLander.Left + picLander.Width And Asteroid2.Visible = True Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
HighScores
picLander.Picture = picSmash.Picture
Command1.Visible = True
End If

If Asteroid3.Top + Asteroid3.Height >= picLander.Top And Asteroid3.Top <= picLander.Top + picLander.Height And Asteroid3.Left >= picLander.Left And Asteroid3.Left + Asteroid3.Width <= picLander.Left + picLander.Width And Asteroid3.Visible = True Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
HighScores
picLander.Picture = picSmash.Picture
Command1.Visible = True
End If

If Asteroid4.Top + Asteroid4.Height >= picLander.Top And Asteroid4.Top <= picLander.Top + picLander.Height And Asteroid4.Left >= picLander.Left And Asteroid4.Left + Asteroid4.Width <= picLander.Left + picLander.Width And Asteroid4.Visible = True Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
HighScores
picLander.Picture = picSmash.Picture
Command1.Visible = True
End If

If Asteroid5.Top + Asteroid5.Height >= picLander.Top And Asteroid5.Top <= picLander.Top + picLander.Height And Asteroid5.Left >= picLander.Left And Asteroid5.Left + Asteroid5.Width <= picLander.Left + picLander.Width And Asteroid5.Visible = True Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
HighScores
picLander.Picture = picSmash.Picture
Command1.Visible = True
End If

If Asteroid6.Top + Asteroid6.Height >= picLander.Top And Asteroid6.Top <= picLander.Top + picLander.Height And Asteroid6.Left >= picLander.Left And Asteroid6.Left + Asteroid6.Width <= picLander.Left + picLander.Width And Asteroid6.Visible = True Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
HighScores
picLander.Picture = picSmash.Picture
Command1.Visible = True
End If

If Asteroid7.Top + Asteroid7.Height >= picLander.Top And Asteroid7.Top <= picLander.Top + picLander.Height And Asteroid7.Left >= picLander.Left And Asteroid7.Left + Asteroid7.Width <= picLander.Left + picLander.Width And Asteroid7.Visible = True Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
HighScores
picLander.Picture = picSmash.Picture
Command1.Visible = True
End If

If Asteroid8.Top + Asteroid8.Height >= picLander.Top And Asteroid8.Top <= picLander.Top + picLander.Height And Asteroid8.Left >= picLander.Left And Asteroid8.Left + Asteroid8.Width <= picLander.Left + picLander.Width And Asteroid8.Visible = True Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
HighScores
picLander.Picture = picSmash.Picture
Command1.Visible = True
End If

If Asteroid9.Top + Asteroid9.Height >= picLander.Top And Asteroid9.Top <= picLander.Top + picLander.Height And Asteroid9.Left >= picLander.Left And Asteroid9.Left + Asteroid9.Width <= picLander.Left + picLander.Width And Asteroid9.Visible = True Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
HighScores
picLander.Picture = picSmash.Picture
Command1.Visible = True
End If

If Asteroid10.Top + Asteroid10.Height >= picLander.Top And Asteroid10.Top <= picLander.Top + picLander.Height And Asteroid10.Left >= picLander.Left And Asteroid10.Left + Asteroid10.Width <= picLander.Left + picLander.Width And Asteroid10.Visible = True Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
HighScores
picLander.Picture = picSmash.Picture
Command1.Visible = True
End If

If Asteroid11.Top + Asteroid11.Height >= picLander.Top And Asteroid11.Top <= picLander.Top + picLander.Height And Asteroid11.Left >= picLander.Left And Asteroid11.Left + Asteroid11.Width <= picLander.Left + picLander.Width And Asteroid11.Visible = True Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
HighScores
picLander.Picture = picSmash.Picture
Command1.Visible = True
End If

If Asteroid12.Top + Asteroid12.Height >= picLander.Top And Asteroid12.Top <= picLander.Top + picLander.Height And Asteroid12.Left >= picLander.Left And Asteroid12.Left + Asteroid12.Width <= picLander.Left + picLander.Width And Asteroid12.Visible = True Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
HighScores
picLander.Picture = picSmash.Picture
Command1.Visible = True
End If

If Asteroid13.Top + Asteroid13.Height >= picLander.Top And Asteroid13.Top <= picLander.Top + picLander.Height And Asteroid13.Left >= picLander.Left And Asteroid13.Left + Asteroid13.Width <= picLander.Left + picLander.Width And Asteroid13.Visible = True Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
HighScores
picLander.Picture = picSmash.Picture
Command1.Visible = True
End If

If Asteroid14.Top + Asteroid14.Height >= picLander.Top And Asteroid14.Top <= picLander.Top + picLander.Height And Asteroid14.Left >= picLander.Left And Asteroid14.Left + Asteroid14.Width <= picLander.Left + picLander.Width And Asteroid14.Visible = True Then
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
HighScores
picLander.Picture = picSmash.Picture
Command1.Visible = True
End If
End Sub

Private Sub Timer4_Timer()
lblScore.Caption = lblScore.Caption + 5
End Sub

Private Sub HighScores()
Dim Score
' Check High Scores
If Val(lblScore.Caption) < Val(Form4.lblScore5.Caption) Then Exit Sub

If Val(lblScore.Caption) > Val(Form4.lblScore5.Caption) And Val(lblScore.Caption) < Val(Form4.lblScore4.Caption) Then
Score = InputBox("You have a new High Score of: " & lblScore.Caption & "!  Enter your name:", "High Score")
Form4.lblName5.Caption = Score
Form4.lblScore5.Caption = lblScore
GoTo isscore
End If

If Val(lblScore.Caption) > Val(Form4.lblScore4.Caption) And Val(lblScore.Caption) < Val(Form4.lblScore3.Caption) Then
Score = InputBox("You have a new High Score of: " & lblScore.Caption & "!  Enter your name:", "High Score")
Form4.lblName4.Caption = Score
Form4.lblScore4.Caption = lblScore
GoTo isscore
End If

If Val(lblScore.Caption) > Val(Form4.lblScore3.Caption) And Val(lblScore.Caption) < Val(Form4.lblScore2.Caption) Then
Score = InputBox("You have a new High Score of: " & lblScore.Caption & "!  Enter your name:", "High Score")
Form4.lblName3.Caption = Score
Form4.lblScore3.Caption = lblScore
GoTo isscore
End If

If Val(lblScore.Caption) > Val(Form4.lblScore2.Caption) And Val(lblScore.Caption) < Val(Form4.lblScore1.Caption) Then
Score = InputBox("You have a new High Score of: " & lblScore.Caption & "!  Enter your name:", "High Score")
Form4.lblName2.Caption = Score
Form4.lblScore2.Caption = lblScore
GoTo isscore
End If

If Val(lblScore.Caption) > Val(Form4.lblScore2.Caption) Then
Score = InputBox("You have the Highest Score of: " & lblScore.Caption & "!  Enter your name:", "High Score")
Form4.lblName1.Caption = Score
Form4.lblScore1.Caption = lblScore
Else
Form4.lblName2.Caption = Score
Form4.lblScore2.Caption = lblScore
GoTo isscore
End If

isscore:
If Score = "" Then Score = "Unknown"
Form4.Show
End Sub
