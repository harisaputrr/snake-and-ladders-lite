VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00151294&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "[Lite] Snakes and Ladders"
   ClientHeight    =   10335
   ClientLeft      =   -75
   ClientTop       =   90
   ClientWidth     =   10905
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   689
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   727
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   255
      Left            =   8280
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   8280
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   615
      Left            =   8760
      TabIndex        =   12
      Text            =   "0"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "AvantGarde Bk BT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   11
      Top             =   7560
      Width           =   1935
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   10200
      Top             =   120
   End
   Begin VB.CommandButton lblEnd 
      Caption         =   "Exit Game"
      BeginProperty Font 
         Name            =   "AvantGarde Bk BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   6
      Top             =   8280
      Width           =   1935
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   10200
      Top             =   8880
   End
   Begin VB.Timer tmrAnimateMove 
      Interval        =   200
      Left            =   8880
      Top             =   3600
   End
   Begin VB.Timer tmrAnimatePointer 
      Interval        =   20
      Left            =   8760
      Top             =   8880
   End
   Begin VB.CommandButton cmdroll 
      Caption         =   "Roll Dice"
      BeginProperty Font 
         Name            =   "AvantGarde Bk BT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      Picture         =   "Form1.frx":0BCA
      TabIndex        =   0
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9480
      Top             =   8880
   End
   Begin VB.CommandButton cmdstop 
      Caption         =   "Stop Dice"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "AvantGarde Bk BT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   1
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SCORE"
      BeginProperty Font 
         Name            =   "DK Lemon Yellow Sun"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   13
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DICE"
      BeginProperty Font 
         Name            =   "DK Lemon Yellow Sun"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   10
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "   Joy"
      Height          =   255
      Left            =   9480
      TabIndex        =   9
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Character"
      BeginProperty Font 
         Name            =   "DK Lemon Yellow Sun"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   8
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1005
      Left            =   120
      Picture         =   "Form1.frx":328FD
      Top             =   120
      Width           =   8475
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "21:21:21"
      BeginProperty Font 
         Name            =   "Pro55"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   8640
      TabIndex        =   7
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Appetite"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0041640D&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   9600
      Width           =   10695
   End
   Begin VB.Label lblTurn 
      BackStyle       =   0  'Transparent
      Caption         =   "Turn: Player 1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Top             =   14040
      Width           =   1695
   End
   Begin VB.Image pointer1 
      Height          =   405
      Left            =   9480
      Picture         =   "Form1.frx":416A6
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   525
   End
   Begin VB.Image pointer2 
      Height          =   405
      Left            =   9720
      Picture         =   "Form1.frx":41AFB
      Stretch         =   -1  'True
      Top             =   13440
      Width           =   525
   End
   Begin VB.Label lblPlayer2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player 2: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8640
      TabIndex        =   3
      Top             =   11760
      Width           =   1695
   End
   Begin VB.Label lblPlayer1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player 1: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8640
      TabIndex        =   2
      Top             =   11280
      Width           =   1695
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   30
      Left            =   825
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   29
      Left            =   2505
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   28
      Left            =   4170
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   27
      Left            =   5835
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   26
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   25
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   24
      Left            =   5880
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   23
      Left            =   4230
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   22
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   21
      Left            =   840
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   20
      Left            =   840
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   19
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   18
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   17
      Left            =   5880
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   16
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   15
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   14
      Left            =   5880
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   13
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   12
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   11
      Left            =   840
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   10
      Left            =   855
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   9
      Left            =   2535
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   8
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   7
      Left            =   5865
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   6
      Left            =   7530
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   5
      Left            =   7605
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   4
      Left            =   5850
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   3
      Left            =   4260
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   2
      Left            =   2565
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   285
   End
   Begin VB.Image coordinate 
      Height          =   405
      Index           =   1
      Left            =   840
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   285
   End
   Begin VB.Image Player 
      Height          =   405
      Index           =   1
      Left            =   9600
      Picture         =   "Form1.frx":41F45
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   285
   End
   Begin VB.Image Player 
      Height          =   405
      Index           =   2
      Left            =   9840
      Picture         =   "Form1.frx":422CB
      Stretch         =   -1  'True
      Top             =   13800
      Width           =   285
   End
   Begin VB.Image dice 
      Height          =   750
      Left            =   9360
      Picture         =   "Form1.frx":42651
      Top             =   5160
      Width           =   750
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      BorderWidth     =   5
      FillColor       =   &H0000FFFF&
      Height          =   975
      Left            =   8760
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H009CE430&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   120
      Top             =   9600
      Width           =   10695
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      BorderWidth     =   5
      Height          =   1215
      Left            =   8760
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00DDA531&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H00004080&
      Height          =   8295
      Left            =   8640
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Image Image9 
      Height          =   8265
      Left            =   120
      Picture         =   "Form1.frx":44443
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   8415
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Left            =   8640
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim diceValue, p1value, p1back As Integer
Dim ctr, p1loop As Integer
Dim roll, p1 As Boolean
Dim turn As Integer
Dim abante1, adtras1 As Boolean
Dim defaultTop1, defaultLeft1 As Integer
Dim msgdelay As Integer
Option Explicit
Private Sub cmdroll_Click()
roll = True
cmdstop.Enabled = True
cmdroll.Enabled = False
End Sub

Private Sub cmdstop_Click()
lblTurn.Caption = "Turn: Player " & turn
If turn = 0 Then lblTurn.Caption = "Turn: Player 2"
roll = False
cmdroll.Enabled = True
cmdstop.Enabled = False
turn = turn + 1
If turn > 1 Then turn = 1
'player 1
If turn = 1 Then
p1 = True
abante1 = True
End If

End Sub


Private Sub Command1_Click()
Unload Me
Form2.Show
End Sub

Sub Command2_Click()
p1value = 2
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player 1 Mendapatkan Soal"
Timer2.Enabled = True
Soal1.Show
lblMessage.Visible = True
End Sub


Private Sub Command3_Click()
p1value = 30
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player 1 Mendapatkan Soal"
Timer2.Enabled = True
Soal1.Show
lblMessage.Visible = True
End Sub

Private Sub Form_Load()
roll = False
p1 = False
p1value = 0
ctr = 17
defaultTop1 = Player(1).Top
defaultLeft1 = Player(1).Left

End Sub


Private Sub lblEnd_Click()
End
End Sub

'Private Sub Picture1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' simmple OLE drag/drop example
    
    ' use class to load 1st file that was dropped, if more than one. Unicode compatible
 '   If cImage.LoadPicture_DropedFiles(Data, 1, 256, 256) Then
  '      If Not cShadow Is Nothing Then
   '         CreateNewShadowClass
    '    Else
     '       RefreshImage
      '  End If
       ' ShowImage False, True
    'End If

'End Sub
'End Sub

Private Sub Timer1_Timer()
If roll = True Then
diceValue = Int(Rnd * 6) + 1
dice.Picture = LoadPicture(App.Path & "\Image\" & diceValue & ".bmp")
End If
End Sub

Private Sub Timer2_Timer()
msgdelay = msgdelay + 1
If msgdelay >= 5 Then
lblMessage.Visible = False
msgdelay = 0
Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
Label1.Caption = Time
End Sub

Private Sub tmrAnimateMove_Timer()
'player1
If p1 = True Then
cmdroll.Enabled = False

If abante1 = True Then
adtras1 = False
p1value = p1value + 1
If p1value > 30 Then
p1value = 30
abante1 = False
adtras1 = True
End If
End If

If adtras1 = True Then
p1value = p1value - 1
End If

Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
lblPlayer1.Caption = "Player 1: " & p1value
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
p1loop = p1loop + 1

If p1loop >= diceValue Then
cmdroll.Enabled = True
p1 = False
p1loop = 0
If p1value = 30 Then
MsgBox "Game Selesai", vbOKOnly
Form7.Show
Unload Me
End If
If p1value = 2 Then
Soal21
End If
If p1value = 3 Then
Soal31
End If
If p1value = 4 Then
Soal41
End If
If p1value = 5 Then
Soal51
End If
If p1value = 6 Then
Soal61
End If
'====================
If p1value = 7 Then
Lseven1
End If
'====================
If p1value = 8 Then
Soal81
End If
If p1value = 9 Then
Soal91
End If
'====================
If p1value = 10 Then
Sten1
End If
'====================
If p1value = 11 Then
Soal111
End If
If p1value = 12 Then
Soal121
End If
If p1value = 13 Then
Soal131
End If
'====================
If p1value = 14 Then
Lfourteen1
End If
'====================
If p1value = 15 Then
Soal151
End If
If p1value = 16 Then
Soal161
End If
If p1value = 17 Then
Soal171
End If
If p1value = 18 Then
Soal181
End If
If p1value = 19 Then
Soal191
End If
If p1value = 20 Then
Soal201
End If
If p1value = 21 Then
Soal211
End If
If p1value = 22 Then
Soal221
End If
'====================
If p1value = 23 Then
Ltwentythree1
End If
'====================
'====================
If p1value = 24 Then
Stwentyfour1
End If
'====================
If p1value = 25 Then
Soal251
End If
If p1value = 26 Then
Soal261
End If
If p1value = 27 Then
Soal271
End If
'====================
If p1value = 28 Then
Stwentyeight1
End If
'====================
If p1value = 29 Then
Soal291
End If

End If

End If
End Sub
Private Sub tmrAnimatePointer_Timer()
ctr = ctr + 1
If ctr > 27 Then ctr = 17

pointer1.Height = ctr
End Sub
Sub new_game()
Player(1).Top = defaultTop1
Player(1).Left = defaultLeft1
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
p1value = 0
End Sub
'===============================================================
Sub Soal21()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player 1 Mendapatkan Soal"
Timer2.Enabled = True
Soal1.Show
lblMessage.Visible = True
End Sub

Sub Soal31()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal2.Show
lblMessage.Visible = True
End Sub

Sub Soal41()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal3.Show
lblMessage.Visible = True
End Sub

Sub Soal51()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal4.Show
lblMessage.Visible = True
End Sub

Sub Soal61()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal5.Show
lblMessage.Visible = True
End Sub

Sub Lseven1()
p1value = 13
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player 1 Telah Menggunakan Tangga"
Timer2.Enabled = True
lblMessage.Visible = True
End Sub

Sub Soal71()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal6.Show
lblMessage.Visible = True
End Sub

Sub Soal81()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal7.Show
lblMessage.Visible = True
End Sub

Sub Soal91()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal8.Show
lblMessage.Visible = True
End Sub

Sub Sten1()
p1value = 2
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player 1 Telah Dimakan Ular"
Timer2.Enabled = True
lblMessage.Visible = True
End Sub

Sub Soal111()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal10.Show
lblMessage.Visible = True
End Sub

Sub Soal121()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal11.Show
lblMessage.Visible = True
End Sub

Sub Soal131()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal12.Show
lblMessage.Visible = True
End Sub

Sub Lfourteen1()
p1value = 16
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player 1 Telah Dimakan Ular"
Timer2.Enabled = True
lblMessage.Visible = True
End Sub

Sub Soal151()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal13.Show
lblMessage.Visible = True
End Sub

Sub Soal161()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal14.Show
lblMessage.Visible = True
End Sub

Sub Soal171()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal15.Show
lblMessage.Visible = True
End Sub

Sub Soal181()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal16.Show
lblMessage.Visible = True
End Sub

Sub Soal191()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal17.Show
lblMessage.Visible = True
End Sub

Sub Soal201()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal18.Show
lblMessage.Visible = True
End Sub

Sub Soal211()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal19.Show
lblMessage.Visible = True
End Sub

Sub Soal221()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal20.Show
lblMessage.Visible = True
End Sub

Sub Soal241()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal9.Show
lblMessage.Visible = True
End Sub

Sub Ltwentythree1()
p1value = 27
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player 1 Telah Memakai Tangga"
Timer2.Enabled = True
lblMessage.Visible = True
End Sub

Sub Stwentyfour1()
p1value = 8
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player 1 Telah Dimakan Ular"
Timer2.Enabled = True
lblMessage.Visible = True
End Sub

Sub Soal251()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal22.Show
lblMessage.Visible = True
End Sub

Sub Soal261()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal22.Show
lblMessage.Visible = True
End Sub

Sub Soal271()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal23.Show
lblMessage.Visible = True
End Sub


Sub Stwentyeight1()
p1value = 12
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player 1 Telah Dimakan Ular"
Timer2.Enabled = True
lblMessage.Visible = True
End Sub

Sub Soal291()
Player(1).Top = coordinate(p1value).Top
Player(1).Left = coordinate(p1value).Left
pointer1.Top = Player(1).Top - pointer1.Height
pointer1.Left = Player(1).Left - 8
lblMessage.Caption = "Player1 Mendapatkan Soal"
Timer2.Enabled = True
Soal24.Show
lblMessage.Visible = True
End Sub
