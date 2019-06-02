VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "[Lite] HighScore - Snakes And Ladders"
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   LinkTopic       =   "Form5"
   ScaleHeight     =   9375
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   8520
      TabIndex        =   10
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   8520
      TabIndex        =   9
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   5280
      TabIndex        =   8
      Top             =   6600
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   1920
      TabIndex        =   7
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   975
      Left            =   8520
      TabIndex        =   4
      Top             =   3480
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Your Score"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   7560
      Width           =   7695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   8520
      Width           =   7695
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   7635
      TabIndex        =   11
      Top             =   9480
      Width           =   7695
   End
   Begin VB.PictureBox DataGrid1 
      Height          =   5655
      Left            =   120
      ScaleHeight     =   5595
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   840
      Width           =   7695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ID :"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   6
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   6600
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "High Score"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deklarasikan Koneksi
Dim kon As New ADODB.Connection

Private Sub Command3_Click()

Dim SQLHapus As String
kon.BeginTrans
SQLHapus = "Delete from HSData where Score='" & Text3 & "'"
kon.Execute SQLHapus
kon.CommitTrans
Adodc1.Refresh

End Sub

Private Sub Form_Load()
Dim msql As String
Dim KoneksiData As String

KoneksiData = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\!Harisaputr_\SD Programing\Snake and Ladder [PO]\HighScore Data.mdb;Persist Security Info=False"

kon.Open KoneksiData
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
kon.BeginTrans
SQLSimpan = "INSERT INTO HSData(No,Nama,Score,Skill)VALUES('" & Text2.Text & "','" & Text1.Text & "','" & Text3.Text & "','" & Text4.Text & "')"
kon.Execute SQLSimpan
kon.CommitTrans

'Refresh datagrid
Adodc1.Refresh

'Kosongkan Text
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Text3_Change()
Text3.Text = Form7.Label2.Caption
End Sub

Private Sub Text4_Change()
Text4.Text = Form7.Label3.Caption
End Sub
