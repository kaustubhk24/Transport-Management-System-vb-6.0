VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Menu Form"
   ClientHeight    =   8145
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form6"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      Caption         =   " EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   6720
      Width           =   4335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "MAINTAINENCE DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   5520
      Width           =   4335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "REPORTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   6120
      Width           =   4335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SALARY DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   4920
      Width           =   4335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DIESEL DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   4320
      Width           =   4335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ROUTE DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   3720
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DRIVER DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   3120
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BUS DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   2520
      TabIndex        =   9
      Top             =   2160
      Width           =   4815
   End
   Begin VB.Frame Frame2 
      Height          =   5535
      Left            =   2400
      TabIndex        =   10
      Top             =   2040
      Width           =   5055
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   1680
      TabIndex        =   6
      Top             =   480
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
Form8.Show
End Sub

Private Sub Command7_Click()
Form7.Show
End Sub

Private Sub Command8_Click()
End
End Sub
