VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   Caption         =   "[Lite] Loading - Snakes And Ladders"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7455
   LinkTopic       =   "Form3"
   ScaleHeight     =   6585
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   4560
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3480
      Top             =   4920
   End
   Begin VB.Image Image2 
      Height          =   1125
      Left            =   0
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   7455
   End
   Begin VB.Image Image1 
      Height          =   7125
      Left            =   -2520
      Picture         =   "Form3.frx":2CB5
      Stretch         =   -1  'True
      Top             =   -960
      Width           =   12465
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    If ProgressBar1.Value >= 100 Then
        Unload Form3
        Form1.Show
    Else
        ProgressBar1.Value = ProgressBar1.Value + 1
    End If
End Sub

