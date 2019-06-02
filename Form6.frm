VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Input Name - Snakes And Ladders"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   ScaleHeight     =   2790
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tulis Nama Anda.."
      BeginProperty Font 
         Name            =   "DK Lemon Yellow Sun"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "Selamat Datang " + Text1.Text, vbOKOnly, "Welcome - Snakes And Ladders"
Unload Form6
End Sub

