VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "TITLE"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9045
   LinkTopic       =   "Form6"
   ScaleHeight     =   4845
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   3600
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SALARY DETAILS"
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DIESEL DETAILS"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ROUTE DETAILS"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DRIVER DETAILS"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BUS DETAILS"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "TRANSPORT MANAGEMENT                SYSTEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   1080
      TabIndex        =   6
      Top             =   1080
      Width           =   6975
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Command3_Click()
Form3.Show
End Sub

Private Sub Command4_Click()
Form4.Show
End Sub

Private Sub Command5_Click()
Form5.Show
End Sub

Private Sub Command6_Click()
End
End Sub
