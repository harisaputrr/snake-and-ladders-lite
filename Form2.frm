VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "[Lite] Snakes And Ladders"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   7560
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Hero"
         Size            =   27
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   5
      Top             =   5640
      Width           =   3495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "High Score = 3700"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Hero"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   3
      Top             =   3960
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Hero"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   2
      Top             =   6480
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Hero"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   1
      Top             =   4800
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "Hero"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   0
      Top             =   3120
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "V2.0"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Image Image3 
      Height          =   6405
      Left            =   2400
      Picture         =   "Form2.frx":BE46
      Top             =   6000
      Width           =   11385
   End
   Begin VB.Image Image2 
      Height          =   6405
      Left            =   -960
      Picture         =   "Form2.frx":1E7F2
      Top             =   4680
      Width           =   11385
   End
   Begin VB.Image Image1 
      Height          =   6525
      Left            =   -2040
      Picture         =   "Form2.frx":3119E
      Stretch         =   -1  'True
      Top             =   -1800
      Width           =   11985
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form3.Show
Unload Me
End Sub

Private Sub Command2_Click()
Form4.Show
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Form5.Show
End Sub
