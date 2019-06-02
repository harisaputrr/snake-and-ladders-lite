VERSION 5.00
Begin VB.Form Soal17 
   Caption         =   "Pertanyaan - Snakes And Ladders"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9900
   LinkTopic       =   "Form18"
   ScaleHeight     =   4320
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Amerika"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Eropa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Asia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   2280
      Width           =   1935
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Amerika"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Jawab..."
      Height          =   615
      Left            =   6960
      TabIndex        =   2
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   615
      Left            =   4080
      TabIndex        =   1
      Top             =   3360
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "*Jika Anda Keluar Dari Pertanyaan Score anda Akan Berkurang 100"
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   4080
      Width           =   9615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Silahkan Jawab Pertanyaan Di Bawah Ini"
      BeginProperty Font 
         Name            =   "Letter Gothic Std"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   9615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Indonesia Terletak Di Benua..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   9615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Score :"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   2295
      Left            =   0
      Top             =   960
      Width           =   9975
   End
End
Attribute VB_Name = "Soal17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value = True Then
 MsgBox "Jawaban Anda Salah, Score Anda Berkurang 100", vbOKOnly, "Perhatian"
 Form1.Text1.Text = Val(Form1.Text1.Text) - Val(--100)
ElseIf Option2.Value = True Then
MsgBox "Jawaban Anda Salah, Score Anda Berkurang 100", vbOKOnly, "Perhatian"
 Form1.Text1.Text = Val(Form1.Text1.Text) - Val(--100)
ElseIf Option3.Value = True Then
  MsgBox "Selamat Jawaban Anda Benar, Score Anda Bertambah 100", vbOKOnly, "Perhatian"
 Form1.Text1.Text = Val(Form1.Text1.Text) + Val(100)
 Unload Me
ElseIf Option4.Value = True Then
 MsgBox "Jawaban Anda Salah, Score Anda Berkurang 100", vbOKOnly, "Perhatian"
 Form1.Text1.Text = Val(Form1.Text1.Text) - Val(--100)
End If
End Sub

Private Sub Command2_Click()
MsgBox "Score Anda Berkurang 100", vbOKOnly, "Perhatian"
If vbOK Then Form1.Text1.Text = Val(Form1.Text1.Text) - Val(--100)
Unload Me
End Sub

Private Sub Text1_Change()
Text1.Text = Form1.Text1.Text
End Sub




