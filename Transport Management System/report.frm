VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Reports"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   LinkTopic       =   "Form8"
   ScaleHeight     =   6945
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   " NEXT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "MAINTANANCE REPORT"
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
      Left            =   1920
      TabIndex        =   6
      Top             =   5040
      Width           =   4095
   End
   Begin VB.CommandButton Command5 
      Caption         =   " SALARY REPORTS"
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
      Left            =   1920
      TabIndex        =   5
      Top             =   4440
      Width           =   4095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DEISEL REPORTS"
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
      Left            =   1920
      TabIndex        =   4
      Top             =   3840
      Width           =   4095
   End
   Begin VB.CommandButton Command3 
      Caption         =   " ROUTE REPORTS"
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
      Left            =   1920
      TabIndex        =   3
      Top             =   3240
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   " DRIVER REPORTS"
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
      Left            =   1920
      TabIndex        =   2
      Top             =   2640
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BUS REPORTS"
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
      Left            =   1920
      TabIndex        =   1
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   1560
      TabIndex        =   7
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   1440
      TabIndex        =   8
      Top             =   1560
      Width           =   5055
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   6960
      TabIndex        =   10
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   " TRANSPORTATION  REPORTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   5895
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DataEnvironment1.Connection1.Open
DataReport1.Show
End Sub

Private Sub Command2_Click()
DataEnvironment2.Connection1.Open
DataReport2.Show
End Sub

Private Sub Command3_Click()
DataEnvironment3.Connection1.Open
DataReport3.Show
End Sub

Private Sub Command4_Click()
DataEnvironment4.Connection1.Open
DataReport4.Show
End Sub

Private Sub Command5_Click()
DataEnvironment5.Connection1.Open
DataReport5.Show
End Sub

Private Sub Command6_Click()
DataEnvironment6.Connection1.Open
DataReport6.Show
End Sub

Private Sub Command9_Click()
Unload Me
Form6.Show
End Sub
