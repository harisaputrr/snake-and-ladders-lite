VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "[Lite] About - Snakes And Ladders"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   4590
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buy Full Version ?"
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   1035
      Left            =   0
      OLEDragMode     =   1  'Automatic
      Picture         =   "Form4.frx":0000
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   4620
   End
   Begin VB.Image Image1 
      Height          =   3765
      Left            =   0
      Picture         =   "Form4.frx":31DD
      Top             =   0
      Width           =   4605
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub
