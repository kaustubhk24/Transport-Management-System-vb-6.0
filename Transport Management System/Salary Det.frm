VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   2520
      TabIndex        =   19
      Top             =   6000
      Width           =   5535
      Begin VB.Frame Frame3 
         Height          =   1215
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   5295
         Begin VB.Frame Frame4 
            Height          =   975
            Left            =   120
            TabIndex        =   21
            Top             =   120
            Width           =   5055
            Begin VB.CommandButton Command4 
               Caption         =   "&CLOSE"
               Height          =   615
               Left            =   3840
               TabIndex        =   25
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton Command3 
               Caption         =   "&EDIT"
               Height          =   615
               Left            =   2640
               TabIndex        =   24
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton Command2 
               Caption         =   "&DELETE"
               Height          =   615
               Left            =   1440
               TabIndex        =   23
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton Command1 
               Caption         =   "&ADD"
               Height          =   615
               Left            =   240
               TabIndex        =   22
               Top             =   240
               Width           =   975
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   5295
      Left            =   3360
      TabIndex        =   0
      Top             =   600
      Width           =   4095
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   2040
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Text            =   "Text3"
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Text            =   "Text4"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Text            =   "Text5"
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Height          =   405
         Left            =   2040
         TabIndex        =   4
         Text            =   "Text6"
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Text            =   "Text7"
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Text            =   "Text8"
         Top             =   4440
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Blood Group"
         Height          =   375
         Left            =   720
         TabIndex        =   18
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Gender"
         Height          =   375
         Left            =   720
         TabIndex        =   17
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Driver Age"
         Height          =   375
         Left            =   720
         TabIndex        =   16
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Driver Name"
         Height          =   375
         Left            =   720
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Driver Identity Number"
         Height          =   375
         Left            =   720
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Driver Address"
         Height          =   375
         Left            =   720
         TabIndex        =   13
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Date of Join"
         Height          =   375
         Left            =   720
         TabIndex        =   12
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Driving Licence Number"
         Height          =   375
         Left            =   720
         TabIndex        =   11
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Driving Licence Expired Date"
         Height          =   375
         Left            =   720
         TabIndex        =   10
         Top             =   4440
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
