VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Hasil - Snakes And Ladders"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   LinkTopic       =   "Form7"
   ScaleHeight     =   2745
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Main Menu"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "THANKS FOR PLAYING !!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   5895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your Skill"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   6255
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6000
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your Score"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF80FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF80FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF80FF&
      Height          =   4815
      Left            =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
Unload Me
Form2.Show
End Sub

Private Sub Form_Load()
Label2.Caption = Form1.Text1.Text
If Label2.Caption >= 2000 Then
Label3.Caption = "Kamu Bukan Orang Rata-Rata!"
ElseIf Label2.Caption >= 1500 And Label2.Caption <= 1999 Then
Label3.Caption = "Kamu Orang Rata-Rata!"
ElseIf Label2.Caption >= 1000 And Label2.Caption <= 1499 Then
Label3.Caption = "Kamu Di Bawah Rata-Rata!"
ElseIf Label2.Caption >= 500 And Label2.Caption <= 999 Then
Label3.Caption = "Bodoh!, Kamu Kurang Ilmu!"
End If
If Label2.Caption <= 499 Then
Label3.Caption = "Kamu Sangat Bodoh!"
End If
If Label2.Caption <= -1 Then
Label3.Caption = "Bodoh Banget Sih Lo!!"
End If
End Sub

